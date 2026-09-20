"""
CS Performance dashboard -- core computation engine (pure pandas, no Streamlit).

Sep 2026, per Mahmoud -- built as the successor to the old Carecomm-Customer-Support-Rep
Streamlit tool, adapted to:
  - the live Google Sheet shape (Agents ID / Calls / Chats / Agents Activity / Schedule),
    with Schedule now in LONG format (Employee name, Date, Shift, Is WFH) instead of one
    row per employee with a column per date
  - Chats matched by numeric Assignee -> Agent ID (the live Chats export has no agent
    name column at all, only the numeric id)
  - the Bahrain team excluded entirely from scope (their real chat volume belongs to a
    separate BH support queue, not this report)
  - WFH -> 9-hour shift rule: initially simplified (Sep 2026, per Mahmoud) to drop a
    gender-inferred "female agents on 4 PM - 12 AM are always WFH" special case, for
    lacking a clear basis -- then, Sep 15 2026, Mahmoud explicitly reinstated a version
    of it (standing, every day, not just Fri/Sat/Sun) plus a second new standing rule:
    the overnight 12 AM-9 AM shift is WFH for everyone. See WFH_ALWAYS_SHIFT_ALL /
    WFH_ALWAYS_SHIFT_NON_MALE below for the exact current rules -- they combine with the
    Schedule's own "Is WFH" flag and the Fri/Sat/Sun default (whichever fires first wins;
    multiple can apply to the same day). Sep 16 2026, per Mahmoud: the "non-male" half of
    that rule now reads a live "Gender" column he added to the Agents ID tab (see
    `gender_lookup` in build_roster / compute_adherence) instead of a hardcoded name list,
    and the Fri/Sat/Sun default now only fires when the Schedule's Is WFH cell was blank
    for that day -- an explicit "No" is respected, not overridden.
  - the Agents Activity Timestamp column mixing plain-text and Excel-auto-converted
    datetime cells, which silently swaps day/month for the auto-converted ones unless
    corrected (fix_activity_ts)
  - Calls State always broken out into its own 8 individual states, never collapsed into
    a single Answered/Missed number, so the underlying cause stays visible per agent
  - FCR = no return contact from the same customer (Contact ID) within 7 days of a
    resolved chat's close time
  - Adherence %, REDESIGNED Sep 19 2026 per Mahmoud: after the chat-backed-fill fix
    below started routinely pushing Actual Minutes past Planned Minutes for almost
    every agent (traced to conversations left open for hours/days being credited
    as full continuous presence), Adherence was rebuilt to the plain WFM-standard
    definition -- (actual login time -> actual logout time), clipped to the
    scheduled window, as a share of Planned Minutes -- instead of summed Actual
    Minutes / Planned Minutes. Whether the agent was actually busy vs. idle DURING
    that clipped span is deliberately left to Occupancy % (a separate, existing
    metric) rather than re-checked inside Adherence too. See CHAT_MIN_PER_MESSAGE /
    CHAT_MAX_CREDIT_MINUTES and the 'Adherent Minutes' comment in compute_adherence.
"""
import datetime as dt
import io

import numpy as np
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from openpyxl.styles import Font, PatternFill
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter

# Read-only -- this tool never writes back to the sheet.
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/drive.readonly',
]
GOOGLE_SHEETS_EPOCH = dt.date(1899, 12, 30)

# The only columns this file ever actually reads off the live Calls / Chats tabs (see
# build_roster, compute_calls, compute_chats below) -- used to fetch just these instead
# of every column those two tabs have (see _worksheet_to_df_cols/fetch_sheet_data). Both
# tabs carry several wide free-text columns nothing here uses at all (Call Summary,
# Closing Note Summary, and more) -- fetching those on every load was most of the actual
# payload for no benefit.
CALLS_COLS_USED = ['Agent', 'Created', 'State', 'Handling Duration', 'Holding Duration', 'Type']
CHATS_COLS_USED = [
    'DateTime Conversation Started', 'DateTime Conversation Resolved', 'Contact ID',
    'Assignee', 'First Response Time', 'Resolution Time', 'Conversation Category',
    'First Assignment to First Response Time',
    # Sep 19 2026, per Mahmoud -- needed for the chat-backed-presence credit cap
    # (see CHAT_MIN_PER_MESSAGE / CHAT_MAX_CREDIT_MINUTES below).
    'Number of Incoming Messages', 'Number of Outgoing Messages',
]


def get_client(service_account_info):
    creds = Credentials.from_service_account_info(service_account_info, scopes=SCOPES)
    return gspread.authorize(creds)

# ---------------------------------------------------------------------------
# Roster / scope
# ---------------------------------------------------------------------------
BAHRAIN = {'Sayed Hadi Alwedaei', 'Fatima Hassan', 'Zainab Abbas', 'Mahdi Ali'}

TEAM_OVERRIDE = {
    'Heba Tarek': 'Logistics', 'Mayar Khaled': 'Logistics', 'Nada Esaam': 'Logistics',
}

# canonical Agents-ID name -> known alternate spellings in Calls / Agents Activity
ALIASES = {
    'Mostafa Gomaa': ['mostafa gomaa'],
    'Nariman Shedid': ['nariman shedid', 'nariman ezzat'],
    'Naira Emad': ['naira emad', 'nayera emad', 'nayira emad'],
    'Waad Yassin': ['waad yassin', 'waad mohamed'],
    'Heba Tarek': ['heba tarek'],
    'Mayar Khaled': ['mayar khaled'],
    'Nada Esaam': ['nada esaam'],
    'Duha Younis': ['duha younis'],
    # 'Hagar Shaban' merged in here Sep 19 2026, per Mahmoud (direct confirmation) --
    # Schedule's "Hagar Shaban" and Agents-ID's "Hagar Ahmed" are the SAME person, not
    # two different agents. Previously kept deliberately separate (see the old comment
    # this replaced) because that was an open, unconfirmed question -- while separate,
    # Schedule rows under "Hagar Shaban" had no Agent ID to match against, so Chats
    # could never be attributed to her at all (Assignee is ID-based only) and her
    # Occupancy % read as near-zero for that reason alone, not real idle time.
    'Hagar Ahmed': ['hagar ahmed', 'hagar shaban'],
    'Nada Sayed': ['nada sayed'],
    'Karim Mohamed': ['karim mohamed'],
    'Ramy Maher': ['ramy maher'],
    'Abdallah Mahmoud': ['abdallah mahmoud'],
    'Salma Adel': ['salma adel'],
    'Amr Hazem': ['amr hazem'],
    'Ahmed Ashraf': ['ahmed ashraf', 'ahmed fathy'],
    'Muhammed Hesham': ['muhammed hesham', 'mohamed hesham', 'muhamed hesham'],
    'Mohamed Sayed': ['mohamed sayed', 'mohamed sayed ahmed'],
    'Hassan Badawy': ['hassan badawy'],
    'Sondos Tarik': ['sondos tarik', 'sondos tarek'],
    'Basma Mostafa': ['basma mostafa'],
    'Sara Hussien': ['sara hussien', 'sara hussein'],
    'Samaa Ahmed': ['samaa ahmed', 'sama ahmed'],
}

# Schedule's "Employee name" (full/formal name) -> canonical Agents-ID name
SCHEDULE_NAME_MAP = {
    'Waad Alla El-dein Elsayed Mohamed': 'Waad Yassin',
    'Nariman Ezzat Amin Nagdy': 'Nariman Shedid',
    'Nayira Emad Hamdy Abu-Nar': 'Naira Emad',
    # 'Hagar Shaban' -> 'Hagar Ahmed', confirmed Sep 19 2026 per Mahmoud -- same person
    # (see the ALIASES comment above); this is what actually pulls her Schedule rows
    # onto her real Agents-ID identity/Agent ID, not the alias list alone.
    'Hagar Shaban': 'Hagar Ahmed',
    'Ahmed Ashraf': 'Ahmed Ashraf',
    'Ramy Maher': 'Ramy Maher',
    'Sara Hussien': 'Sara Hussien',
    'Hassan Badawy': 'Hassan Badawy',
    'Samaa Ahmed': 'Samaa Ahmed',
    'Mohamed Sayed Ahmed': 'Mohamed Sayed',
    'Muhammed Hesham': 'Muhammed Hesham',
    'Sondos Tarik': 'Sondos Tarik',
    'Basma Mostafa': 'Basma Mostafa',
}

# ---------------------------------------------------------------------------
# Shift vocabulary
# ---------------------------------------------------------------------------
SHIFT_NORMALIZE = {
    '11:59 PM - 9 AM': '12 AM - 9 AM',
    '12 PM - 9 AM': '12 AM - 9 AM',  # confirmed typo for the night shift
}
NON_WORKING = {'Day Off', 'Ann', 'CL', 'Sick', 'PH', 'Termination'}
UNRESOLVED_LABELS = {'Task', 'Task - Normal Shift'}  # meaning still TBD -- excluded, flagged
SHIFT_PLANNED_MIN = {
    '9 AM - 5 PM': 420, '11 AM - 7 PM': 420, '4 PM - 12 AM': 420, '12 AM - 9 AM': 480,
}
EIGHT_HOUR_TYPES = {'9 AM - 5 PM', '11 AM - 7 PM', '4 PM - 12 AM'}

# Company policy, per Mahmoud (Sep 2026, REVISED Sep 16 2026): on Friday/Saturday/Sunday,
# a real working shift (Day Off/leave excluded, same NON_WORKING check as everywhere
# else) DEFAULTS to work-from-home only when the Schedule's "Is WFH" column was actually
# left blank that day. An explicit "No" in the sheet is respected, not overridden -- the
# original Sep 2026 version forced WFH=Yes on a weekend even over an explicit "No"; Mahmoud
# walked that back once the Schedule started reliably having a real Yes/No either way, not
# just gaps. Every day this default actually changes the flag (i.e. the cell was blank) is
# tracked in daily_audit ('WFH Overridden' column) and surfaced in the Diagnostics tab.
WFH_OVERRIDE_WEEKDAYS = {'Friday', 'Saturday', 'Sunday'}

# Sep 15 2026, per Mahmoud -- re-adds, on his explicit direct confirmation, a narrower
# version of the "female agents on 4 PM-12 AM are always WFH" rule the module docstring
# above says he had previously asked removed for lacking a clear basis. This time it's a
# standing (every day, not just Fri/Sat/Sun) rule -- every roster agent NOT read as 'male'
# from the live Gender lookup (incl. Waad Yassin, confirmed explicitly) on a 4 PM-12 AM
# shift is WFH. Originally (Sep 15) driven by a hardcoded MALE_AGENTS name list Mahmoud
# dictated directly; REVISED Sep 16 2026 to read a live "Gender" column he added to the
# Agents ID tab instead (see `gender_lookup` in build_roster), so a new hire's coverage
# comes from the sheet, not a code change. He also named a second, gender-independent
# standing rule in the same Sep 15 breath: the overnight 12 AM-9 AM shift is WFH for
# EVERYONE, always -- this one doesn't change Planned Minutes (that shift is already 480
# via SHIFT_PLANNED_MIN, not one of EIGHT_HOUR_TYPES), it only makes the WFH flag/tracking
# accurate.
WFH_ALWAYS_SHIFT_ALL = {'12 AM - 9 AM'}
WFH_ALWAYS_SHIFT_NON_MALE = {'4 PM - 12 AM'}

BREAK_STATES = {'Away - short break', 'Away - lunch break', 'Away - gomaa prayer'}
COACHING_STATE = 'Away - coaching'
TRAINING_STATE = 'Away - training'
TECH_STATE = 'Away - technical issue'
OFFLINE_STATE = 'Offline'
RECLASS_STATES = BREAK_STATES | {COACHING_STATE, TRAINING_STATE}
# Max length (minutes) of an Offline gap that can still be bridged into the
# identical state sitting on both sides of it (see the reclassify step in
# compute_adherence). Sep 2026, per Mahmoud, from a real case: Muhammed Hesham,
# 13 Aug -- a genuine ~5-hour Offline gap (09:00-14:12) got bridged into "Away -
# coaching" purely because a few-second coaching blip happened to sit on each
# side of it, which is not remotely a real coaching session. Real bridgeable
# gaps in the Aug 2026 data (a brief reconnect blip mid-break/coaching/training)
# topped out well under this; the one clear outlier was 311 minutes.
RECLASS_MAX_GAP_MINUTES = 90
# A label for Offline time that Chats prove was real work -- see
# _fill_offline_with_chat below. Deliberately distinct from 'Available' (not
# folded into it at the source) so the daily/agent breakdown can still show how
# many minutes came from genuine Maqsam presence vs. this fallback.
CHAT_BACKED_STATE = 'Available (Chat, no Maqsam)'
# Sep 19 2026, per Mahmoud -- a conversation's [Started, Resolved] span is NOT the
# same thing as continuous agent work: a chat can sit open for hours (even days)
# waiting on a slow customer reply, or simply forgotten, with almost no messages
# in it. Confirmed on real Aug 2026 data (Nariman Shedid, 19 Jul: an entire 9-hour
# overnight shift's Offline gap was covered by ONE conversation that stayed open
# 25+ hours with only 3 messages total in it; 26 Jul: a second gap covered by one
# conversation open 40+ hours with 27 messages). Message count vs. duration across
# the whole dataset showed no real relationship past ~1 hour (a 4-24h-open chat
# averaged the same ~4-5 messages as a 15-60min one) -- confirming the extra open
# time is mostly silence, not extra work. So each conversation's credited presence
# window is capped at CHAT_MIN_PER_MESSAGE minutes per message (incoming +
# outgoing), up to CHAT_MAX_CREDIT_MINUTES total, anchored at the conversation's
# own start -- never at more than its own real [Started, Resolved] span either.
# This is a data-cleaning proxy for "evidence some work happened", not a claim
# that the agent worked continuously for exactly that many minutes.
CHAT_MIN_PER_MESSAGE = 10
CHAT_MAX_CREDIT_MINUTES = 60

# Serviced and Successful merged into one "Serviced" bucket, per Mahmoud (Sep 2026) --
# the platform's own docs draw no real distinction Ops needs to track separately.
CALL_STATE_NORMALIZE = {'Successful': 'Serviced'}
CALL_STATES = ['Serviced', 'Dropped', 'No Answer', 'Abandoned', 'Blocked', 'Busy', 'Failed']
CALL_ANSWERED_STATES = {'Serviced'}

# Inbound/Outbound split, added Sep 2026 per Mahmoud -- the raw Calls tab's "Type"
# column is already exactly this (values seen: "Inbound"/"Outbound"), normalized
# the same way State is (strip + title-case) so a stray extra space or lowercase
# entry doesn't silently fall outside both buckets.
CALL_DIRECTIONS = ['Inbound', 'Outbound']

FCR_WINDOW_DAYS = 7

# ---------------------------------------------------------------------------
# CEO Q3 2026 scorecard targets -- the CS-relevant Performance KPIs this tool
# actually has real numbers for (Item 1: Call Answer Rate, weighted 0.2; Item 6:
# First Contact Resolution, weighted 0.1; Draft Order AOV, $90-130 depending on
# market -- only usable once AOV is converted to USD, see compute_aov_by_agent).
# Bands come straight from the "GC KPIs (Younes) Q3 2026 CEO direction"
# scorecard's Below(Red)/Target(Green)/Exceed(Stretch) columns. The remaining
# Performance KPIs (Net Delivery, On-Time Delivery, Delivery Time, Packaging
# Quality, CSAT) belong to Logistics, not this tool.
CEO_TARGETS = {
    'answered_rate': {'red_max': 90.0, 'green_min': 95.0, 'stretch_min': 98.0},
    'fcr_rate': {'red_max': 80.0, 'green_min': 90.0, 'stretch_min': None},
}

TARGET_BADGES = {
    'stretch': '🔵 Exceeds CEO stretch target',
    'green': '🟢 Meets CEO target',
    'amber': '🟡 Below CEO target',
    'red': '🔴 Below CEO red line',
}

# AOV per-market targets, REPLACED Sep 20 2026 per Mahmoud -- there used to be a flat
# $90-130 'aov_usd' entry under CEO_TARGETS, my own guess at reading the scorecard's
# summary text before Mahmoud shared the actual scorecard file. It was wrong on two
# counts: the real target is a DIFFERENT Below/Target/Exceed band per market (scorecard
# item 8, "Draft Order Sales (AOV per Agent, USD)"), not one number for everyone, and
# Target/Exceed are the CEO's own forecast +20%/+30%, not a flat $90/$130. Values below
# are copied directly from that sheet. Iraq is not covered by this KPI in the scorecard
# at all -- deliberately absent, so an Iraq row reads as unbadged ('--') rather than
# silently inheriting someone else's target. This can no longer be checked against the
# blended per-agent AOV table below (one agent can sell across several markets with
# different targets, so no single number is meaningful there) -- only against the
# per-agent-PER-MARKET breakdown, see compute_aov_by_agent_market().
AOV_MARKET_TARGETS = {
    'SA':  {'red_max': 90.0,  'green_min': 108.0, 'stretch_min': 117.0},
    'KW':  {'red_max': 100.0, 'green_min': 120.0, 'stretch_min': 130.0},
    'QA':  {'red_max': 100.0, 'green_min': 120.0, 'stretch_min': 130.0},
    'UAE': {'red_max': 80.0,  'green_min': 96.0,  'stretch_min': 104.0},
    'OM':  {'red_max': 80.0,  'green_min': 96.0,  'stretch_min': 104.0},
}

# Below this many converted orders in a single Agent+Market cell, the average is too
# noisy to badge against a target -- an agent with 1-2 orders in a market would read
# as a flat "missed" or "hit" on what's really a coin flip. Chosen default, per
# Mahmoud (Sep 20 2026) -- raise or lower freely.
MIN_ORDERS_FOR_AOV_TARGET = 5


def target_status(value, red_max, green_min, stretch_min=None):
    """Below(Red) / (implied amber gap) / Target(Green) / Exceed(Stretch), matching
    the CEO scorecard's three named bands plus the unnamed gap between Red and Green
    that the scorecard's own numbers leave open (e.g. 90-94.9% for Answer Rate)."""
    if value is None:
        return None
    if stretch_min is not None and value >= stretch_min:
        return 'stretch'
    if value >= green_min:
        return 'green'
    if value < red_max:
        return 'red'
    return 'amber'


def target_badge(value, key):
    cfg = CEO_TARGETS.get(key)
    if not cfg or value is None:
        return None
    status = target_status(value, cfg['red_max'], cfg['green_min'], cfg.get('stretch_min'))
    return TARGET_BADGES.get(status)


def aov_market_badge(value, market, order_count):
    """Like target_badge, but keyed by market (AOV_MARKET_TARGETS) instead of a single
    CEO_TARGETS entry, and withheld ('--') below MIN_ORDERS_FOR_AOV_TARGET orders or
    for a market the scorecard's AOV KPI doesn't cover at all (e.g. Iraq)."""
    cfg = AOV_MARKET_TARGETS.get(str(market).strip().upper())
    if not cfg or value is None or pd.isna(value) or order_count < MIN_ORDERS_FOR_AOV_TARGET:
        return '—'
    status = target_status(value, cfg['red_max'], cfg['green_min'], cfg.get('stretch_min'))
    return TARGET_BADGES.get(status, '—')


def norm(s):
    return str(s).strip().lower()


def to_timedelta(s):
    """Duration cell -> Timedelta, tolerant of every shape the two data sources hand
    back for the same "00:01:42" cell: an 'H:M:S' string (typed-in / CSV), a
    datetime.time (xlsx via openpyxl -- str() on it happens to already read 'H:M:S',
    so the string branch below catches it too), or a bare fraction-of-a-day float
    (gspread's UNFORMATTED_VALUE for a duration-formatted Sheets cell -- e.g. 102
    seconds comes back as 0.0011805..., NOT text -- silently parsed as 0 by the old
    string-only version, which is exactly why Avg Handling Time read all zeros on
    the live-Sheets path)."""
    if pd.isna(s) or s == '':
        return pd.Timedelta(0)
    if isinstance(s, pd.Timedelta):
        return s
    if isinstance(s, dt.timedelta):
        return pd.Timedelta(s)
    if isinstance(s, bool):
        return pd.Timedelta(0)
    if isinstance(s, (int, float)):
        try:
            return pd.Timedelta(days=float(s))
        except Exception:
            return pd.Timedelta(0)
    try:
        h, m, sec = str(s).split(':')
        return pd.Timedelta(hours=int(h), minutes=int(m), seconds=int(sec))
    except Exception:
        return pd.Timedelta(0)


def fmt_td(td):
    # pd.isna() catches None, a plain float NaN, AND pd.NaT in one shot -- the old
    # check here only caught the first two, so pd.NaT (e.g. _mean_td_nonblank's
    # "nothing qualified" return, whenever EVERY conversation for an agent still
    # lacks a value for that time metric -- a normal case on real data, not rare)
    # slipped through into int(td.total_seconds()), which is nan for NaT -> crashed
    # the whole app with "cannot convert float NaN to integer".
    if pd.isna(td):
        return ''
    total_sec = int(td.total_seconds())
    sign = '-' if total_sec < 0 else ''
    total_sec = abs(total_sec)
    h, rem = divmod(total_sec, 3600)
    m, s = divmod(rem, 60)
    return f"{sign}{h:02d}:{m:02d}:{s:02d}"


def _mean_td_nonblank(df, raw_col, td_col):
    """Average of a duration column, over ONLY the rows that actually have a value in
    raw_col -- to_timedelta() maps a blank cell to Timedelta(0) (a genuinely 0-second
    duration and "no value yet" are otherwise indistinguishable once converted), which
    would silently pull the average toward 0 for every row still waiting on that event
    (e.g. a conversation with no response yet has no 'First Response Time' to average
    in at all -- it isn't a 0-second response). Returns pd.NaT if nothing qualifies."""
    if df.empty:
        return pd.NaT
    mask = ~(df[raw_col].isna() | (df[raw_col].astype(str).str.strip() == ''))
    vals = df.loc[mask, td_col]
    return vals.mean() if len(vals) else pd.NaT


def _merge_intervals(intervals):
    """[(start, end), ...] -> sorted, non-overlapping (start, end) tuples. Used to
    collapse an agent's chat-handling windows before measuring occupied time, so
    two chats worked at once count as ONE stretch of busy time, not double."""
    if not intervals:
        return []
    ordered = sorted(intervals, key=lambda iv: iv[0])
    merged = [list(ordered[0])]
    for s, e in ordered[1:]:
        if s <= merged[-1][1]:
            merged[-1][1] = max(merged[-1][1], e)
        else:
            merged.append([s, e])
    return [(s, e) for s, e in merged]


def _overlap_minutes(a_intervals, b_intervals):
    """Total minutes where a_intervals and b_intervals overlap. Both are assumed
    already merged/non-overlapping within themselves; a plain double loop is fine
    since both lists are short (one shift window's worth of state/chat segments)."""
    total = 0.0
    for a_s, a_e in a_intervals:
        for b_s, b_e in b_intervals:
            cs, ce = max(a_s, b_s), min(a_e, b_e)
            if ce > cs:
                total += (ce - cs).total_seconds() / 60
    return total


def _fill_offline_with_chat(clipped, chat_ivs, label):
    """Real, evidenced chat-handling time inside an Offline stretch is not really
    "not present" -- it means Maqsam/Agents Activity has a gap (no presence log
    at all) while the agent was demonstrably working via chat instead. Confirmed
    Sep 2026, per Mahmoud, on two real cases: Sara Hussien, 17-18 Aug (zero
    Agents Activity rows either day, but 69 and 76 resolved chats spanning the
    whole shift) and Muhammed Hesham, 13 Aug (a ~5-hour Agents Activity gap with
    54 chats resolved inside it). Any portion of an Offline interval that
    overlaps a chat conversation window [Started, Resolved] is relabeled
    `label` instead of staying Offline; the rest of the Offline interval (no
    chat evidence either) is left alone. `chat_ivs` must be pre-sorted,
    non-overlapping (start, end) tuples, already clipped to the shift window --
    see `_merge_intervals` and the `chat_clipped` list built in
    compute_adherence."""
    if not chat_ivs:
        return clipped
    out = []
    for s, e, st in clipped:
        if st != OFFLINE_STATE:
            out.append([s, e, st])
            continue
        cursor = s
        for cs, ce in chat_ivs:
            os_, oe_ = max(cs, s), min(ce, e)
            if oe_ <= os_:
                continue
            if os_ > cursor:
                out.append([cursor, os_, OFFLINE_STATE])
            out.append([os_, oe_, label])
            cursor = max(cursor, oe_)
        if cursor < e:
            out.append([cursor, e, OFFLINE_STATE])
    return out


def fix_activity_ts(v):
    """Agents Activity's Timestamp column mixes plain-text strings (correctly DD-MM-YYYY)
    with Excel-auto-converted datetime cells that get silently misread as MM-DD-YYYY,
    swapping day/month whenever the original day was <=12. Reconstruct the auto-converted
    ones; parse the string ones with the real format."""
    if isinstance(v, str):
        return pd.to_datetime(v, format='%d-%m-%Y %H:%M:%S', errors='coerce')
    if isinstance(v, (dt.datetime, pd.Timestamp)):
        try:
            return pd.Timestamp(year=v.year, month=v.day, day=v.month,
                                 hour=v.hour, minute=v.minute, second=v.second)
        except ValueError:
            return pd.NaT
    return pd.NaT


# ---------------------------------------------------------------------------
# Loading
# ---------------------------------------------------------------------------
def load_workbook(file_obj):
    xl = pd.ExcelFile(file_obj)
    agents = xl.parse('Agents ID')
    schedule = xl.parse('Schedule')
    calls = xl.parse('Calls')
    chats = xl.parse('Chats')
    activity = xl.parse('Agents Activity')
    return build_roster(agents, schedule, calls, chats, activity)


def build_roster(agents, schedule, calls, chats, activity):
    agents = agents.copy()
    agents['Agent Name'] = agents['Agent Name'].astype(str).str.strip()
    agents = agents[~agents['Agent Name'].isin(BAHRAIN)].reset_index(drop=True)
    # Gender column, added Sep 2026 per Mahmoud directly on the Agents ID tab -- read
    # BEFORE the Agent ID cleanup below so someone without an ID yet still gets a gender
    # entry. Replaces the old hardcoded MALE_AGENTS list for the 4 PM-12 AM WFH rule (see
    # compute_adherence) -- a new hire's gender now comes from the sheet, no code change
    # needed. An agent missing from this lookup entirely (e.g. off-roster, not on the
    # Agents ID tab at all -- see off_roster below) falls back to the same "not male"
    # default the old hardcoded list already implied for anyone it didn't name.
    if 'Gender' in agents.columns:
        gender_lookup = {
            name: str(g).strip().lower()
            for name, g in zip(agents['Agent Name'], agents['Gender'])
            if pd.notna(g) and str(g).strip()
        }
    else:
        gender_lookup = {}
    agents['Team'] = agents.get('Team')
    if agents['Team'].isna().all():
        agents['Team'] = agents['Agent Name'].map(lambda n: TEAM_OVERRIDE.get(n, 'CS'))
    else:
        agents['Team'] = agents.apply(
            lambda r: r['Team'] if pd.notna(r['Team']) else TEAM_OVERRIDE.get(r['Agent Name'], 'CS'), axis=1)
    # A blank/non-numeric Agent ID (e.g. a new hire added to the Agents ID tab before
    # HR/Ops assigned them one) used to crash the whole app on load ("cannot convert
    # float NaN to integer") -- coerce instead, drop those rows from the ID-keyed
    # roster, and keep their names so they aren't silently lost (see off_roster below).
    agents['Agent ID'] = pd.to_numeric(agents['Agent ID'], errors='coerce')
    missing_id_mask = agents['Agent ID'].isna()
    agents_missing_id = sorted(agents.loc[missing_id_mask, 'Agent Name'].tolist())
    agents = agents.loc[~missing_id_mask].reset_index(drop=True)
    agents['Agent ID'] = agents['Agent ID'].astype(int)
    id_to_name = dict(zip(agents['Agent ID'], agents['Agent Name']))
    roster = agents['Agent Name'].tolist()

    # add off-roster-but-scheduled people AND anyone on the Agents ID tab with a
    # blank Agent ID, so Schedule/Calls/Activity data isn't silently dropped just
    # because they have no Agent ID yet -- Chats can't be attributed to them though
    # (Assignee is ID-based), which is why they're flagged separately
    # (agents_missing_id) rather than treated as fully resolved. ("Hagar Shaban" used
    # to be the standing example here -- no longer applies, see SCHEDULE_NAME_MAP.)
    sched_names = set(SCHEDULE_NAME_MAP.values())
    off_roster = sorted((sched_names - set(roster)) | set(agents_missing_id))
    full_roster = roster + off_roster

    calls = calls.copy()
    calls['Agent_n'] = calls['Agent'].map(norm)
    calls['Created'] = pd.to_datetime(calls['Created'], errors='coerce')
    calls['State'] = calls['State'].astype(str).str.strip().map(lambda s: CALL_STATE_NORMALIZE.get(s, s))
    # Inbound/Outbound, added Sep 2026 per Mahmoud -- see CALL_DIRECTIONS above. Anything
    # that doesn't normalize to exactly "Inbound"/"Outbound" (blank, a typo, a value this
    # export never had before) stays as whatever it normalized to rather than getting
    # silently dropped -- it just won't match either direction filter in compute_calls,
    # so it's still visible in the combined (direction=None) totals, only missing from
    # both split blocks. Worth a look if the Inbound+Outbound counts stop summing to Total.
    # Opportunistic like several other columns in this file (e.g. Delivery Date, chats'
    # First Response Time): the live-sheets path only fetches the columns listed in
    # CALLS_COLS_USED (see that list above -- 'Type' was missing from it at first, which
    # crashed the whole app with KeyError: 'Type' instead of degrading gracefully; fixed
    # by adding it there AND here, so a future rename/removal on the live Calls tab fails
    # soft -- Direction just comes back blank -- rather than taking the app down again).
    if 'Type' in calls.columns:
        calls['Direction'] = calls['Type'].astype(str).str.strip().str.title()
    else:
        calls['Direction'] = None

    chats = chats.copy()
    chats['DateTime Conversation Started'] = pd.to_datetime(chats['DateTime Conversation Started'], errors='coerce')
    chats['DateTime Conversation Resolved'] = pd.to_datetime(chats['DateTime Conversation Resolved'], errors='coerce')
    # Opportunistic -- these columns aren't in the minimal Chats shape this tool was
    # first built against, but the raw Carecomm export can carry them. Parsed with the
    # same to_timedelta() used for Calls (tolerant of "H:M:S" text AND the Sheets
    # duration-serial float) whenever the column exists, so richer per-agent chat KPIs
    # (see compute_chats) switch on automatically the moment the live Chats tab has
    # them -- no crash, no extra KPI, if it doesn't.
    for col in ('First Response Time', 'Resolution Time', 'First Assignment to First Response Time'):
        if col in chats.columns:
            chats[col + ' (td)'] = chats[col].map(to_timedelta)
    # Message counts, added Sep 19 2026 per Mahmoud -- the activity proxy behind the
    # chat-backed-presence credit cap (see CHAT_MIN_PER_MESSAGE / CHAT_MAX_CREDIT_MINUTES
    # and compute_adherence's chat_intervals below). Opportunistic like the columns
    # above: missing on an older export just reads as 0 messages, so those conversations
    # simply credit no presence (a total absence of a proxy signal, not proof of no work,
    # but the safer default given no other evidence is available).
    for col in ('Number of Incoming Messages', 'Number of Outgoing Messages'):
        if col in chats.columns:
            chats[col] = pd.to_numeric(chats[col], errors='coerce').fillna(0)
        else:
            chats[col] = 0

    activity = activity.copy()
    activity['Agent_n'] = activity['Agent Name'].map(norm)
    if 'ts' not in activity.columns:
        # xlsx path only -- the live-sheet path (fetch_sheet_data) already fills 'ts'
        # with correctly-parsed timestamps via _serial_to_ts, which has no day/month
        # ambiguity to begin with, so it must NOT be run back through fix_activity_ts
        # (that would blindly swap day/month a second time and corrupt every date).
        activity['ts'] = activity['Timestamp'].map(fix_activity_ts)

    schedule = schedule.copy()
    schedule['sched_name'] = schedule['Employee name'].astype(str).str.strip()
    schedule['canonical'] = schedule['sched_name'].map(lambda n: SCHEDULE_NAME_MAP.get(n, n))
    schedule['Date'] = pd.to_datetime(schedule['Date'], errors='coerce')
    schedule['Shift'] = schedule['Shift'].astype(str).str.strip().map(lambda s: SHIFT_NORMALIZE.get(s, s))
    if 'Is WFH' in schedule.columns:
        schedule['is_wfh_flag'] = schedule['Is WFH'].astype(str).str.strip().str.lower().eq('yes')
        # Raw tri-state alongside the boolean above -- '' means the cell was blank/missing,
        # distinct from an explicit 'no'. Needed for the Fri/Sat/Sun default below (Sep 16
        # 2026, per Mahmoud): blank on a weekend defaults to WFH=Yes, but an explicit "No"
        # in the sheet is respected, not overridden -- the plain boolean above can't tell
        # those two apart (both read as False), so this keeps the distinction available.
        schedule['wfh_raw'] = schedule['Is WFH'].apply(
            lambda v: '' if pd.isna(v) else str(v).strip().lower())
    else:
        schedule['is_wfh_flag'] = False
        schedule['wfh_raw'] = ''

    # Excuse Minutes / Note, added Sep 15 2026 per Mahmoud -- two columns he asked to sit
    # directly on Schedule, next to Is WFH: a manager-approved excuse/lost-time number in
    # minutes for that agent+day, plus a plain-text reason. Opportunistic like Is WFH
    # above -- older Schedule exports without these columns just get 0/blank, no crash.
    # Per the methodology already established for this (Aug 2026 manual reconciliation,
    # reused directly rather than re-derived): approved excuse minutes REDUCE that day's
    # Planned Minutes -- see compute_adherence() below. A blank/non-numeric cell reads as
    # 0 (no excuse that day), never NaN, so it can't silently poison a sum.
    if 'Excuse Minutes' in schedule.columns:
        schedule['excuse_minutes'] = pd.to_numeric(schedule['Excuse Minutes'], errors='coerce').fillna(0.0)
    else:
        schedule['excuse_minutes'] = 0.0
    if 'Note' in schedule.columns:
        # fillna('') BEFORE astype(str) -- with pandas' newer string dtype, a NaN cell
        # converted via .astype(str) can come back as an actual missing value again
        # rather than the literal text "nan", which then slips past the isin() cleanup
        # below and reappears downstream (str(NaN) == "nan") as a bogus "nan" note on
        # every ordinary no-excuse day. Filling first sidesteps that regardless of dtype.
        schedule['excuse_note'] = schedule['Note'].fillna('').astype(str).str.strip()
        schedule.loc[schedule['excuse_note'].isin(['None', 'nan', 'NaN', '<NA>']), 'excuse_note'] = ''
    else:
        schedule['excuse_note'] = ''

    alias_lookup = {}
    for canon, al in ALIASES.items():
        for a in al:
            alias_lookup[a] = canon
    # anyone in scope without an explicit alias list matches on their own lowercased name
    for name in full_roster:
        alias_lookup.setdefault(norm(name), name)

    calls['canonical'] = calls['Agent_n'].map(alias_lookup)
    activity['canonical'] = activity['Agent_n'].map(alias_lookup)
    chats['canonical'] = chats['Assignee'].map(
        lambda v: id_to_name.get(int(v)) if pd.notna(v) else None)

    return {
        'agents': agents, 'schedule': schedule, 'calls': calls, 'chats': chats,
        'activity': activity, 'roster': roster, 'off_roster': off_roster,
        'agents_missing_id': agents_missing_id,
        'full_roster': full_roster, 'id_to_name': id_to_name,
        'gender': gender_lookup,
    }


# ---------------------------------------------------------------------------
# Live Google Sheets loading (reuses build_roster above -- same code path as the
# xlsx upload once the 5 tabs are turned into DataFrames, so nothing about the
# actual computation differs between "upload a file" and "read live").
# ---------------------------------------------------------------------------
def _serial_to_ts(value):
    """A gspread UNFORMATTED_VALUE read gives real dates/datetimes back as plain
    numbers (days since the sheets epoch, like Excel) -- NOT text, so there's no
    string-vs-datetime ambiguity here the way there was with the xlsx export's
    Agents Activity Timestamp column. Falls back to text parsing for a cell someone
    typed in by hand as plain text instead of a real date/time value."""
    if value in (None, ''):
        return pd.NaT
    if isinstance(value, bool):
        return pd.NaT
    if isinstance(value, (int, float)):
        try:
            return pd.Timestamp(GOOGLE_SHEETS_EPOCH) + pd.Timedelta(days=value)
        except (OverflowError, ValueError):
            return pd.NaT
    s = str(value).strip()
    if not s:
        return pd.NaT
    try:
        return pd.Timestamp(pd.to_datetime(s, dayfirst=True))
    except Exception:
        return pd.NaT


def _worksheet_to_df(ws):
    """Raw values -> DataFrame, header row included, tolerant of short trailing rows
    (Sheets omits fully-blank trailing cells) the way a plain get_all_records() isn't."""
    values = ws.get_values(value_render_option='UNFORMATTED_VALUE')
    if not values or len(values) < 2:
        return pd.DataFrame(columns=values[0] if values else [])
    header = [str(h).strip() for h in values[0]]
    width = len(header)
    body = [list(r) + [None] * (width - len(r)) if len(r) < width else r[:width] for r in values[1:]]
    return pd.DataFrame(body, columns=header)



# Calls and Chats are the two tabs that grow fastest (every call/chat ever handled, plus
# now the Customer Support native-export uploads append directly into them too) and each
# carries several wide free-text columns (Call Summary, Closing Note Summary, and more)
# that nothing in this file actually reads -- see compute_calls/compute_chats/
# build_roster above for the full list of what's really used. Fetching every column of
# a tab with tens of thousands of rows just to throw most of it away was the main reason
# a fresh load ("Reading the live sheet...") could take a long time (confirmed Sep 2026,
# per Mahmoud -- the live sheet's own xlsx export is 14+ MB). _worksheet_to_df_cols pulls
# ONLY the named columns, via ONE batched Sheets API request (gspread's batch_get, all
# ranges in a single HTTP round trip -- not one request per column), instead of the full
# row width. Agents ID / Schedule / Agents Activity stay on the full-width read: they're
# either small (Agents ID, Schedule) or already minimal at just 3 columns (Agents
# Activity) with a chats/calls-style growth curve.
def _worksheet_to_df_cols(ws, wanted_cols):
    """Same shape of result as _worksheet_to_df (header row's worth of columns -> a
    DataFrame), but reads ONLY the columns in wanted_cols that actually exist in the
    live header -- silently skipping ones that don't (same 'opportunistic' tolerance as
    the rest of this file already has for a Chats/Calls tab that's missing some column),
    rather than every column the tab happens to have."""
    header = ws.row_values(1)
    if not header:
        return pd.DataFrame(columns=wanted_cols)
    header = [str(h).strip() for h in header]
    present = [c for c in wanted_cols if c in header]
    if not present:
        return pd.DataFrame(columns=wanted_cols)

    ranges = [get_column_letter(header.index(c) + 1) + ':' + get_column_letter(header.index(c) + 1)
              for c in present]
    results = ws.batch_get(ranges, value_render_option='UNFORMATTED_VALUE')

    columns = {}
    max_len = 0
    for col_name, value_range in zip(present, results):
        col_values = [row[0] if row else None for row in value_range]
        # First cell of each range is that column's own header -- drop it here, same as
        # _worksheet_to_df does via values[1:].
        col_values = col_values[1:] if col_values else []
        columns[col_name] = col_values
        max_len = max(max_len, len(col_values))

    for col_name in columns:
        columns[col_name] += [None] * (max_len - len(columns[col_name]))
    return pd.DataFrame(columns)


def fetch_sheet_data(gc, spreadsheet_id):
    """Reads the live Agents ID / Schedule / Calls / Chats / Agents Activity tabs and
    returns them shaped exactly like pd.ExcelFile(...).parse(tab) would, so they can go
    straight into build_roster() unchanged."""
    sh = gc.open_by_key(spreadsheet_id)

    agents = _worksheet_to_df(sh.worksheet('Agents ID'))
    if 'Agent ID' in agents.columns:
        agents['Agent ID'] = pd.to_numeric(agents['Agent ID'], errors='coerce')

    schedule = _worksheet_to_df(sh.worksheet('Schedule'))
    if 'Date' in schedule.columns:
        schedule['Date'] = schedule['Date'].map(_serial_to_ts)

    calls = _worksheet_to_df_cols(sh.worksheet('Calls'), CALLS_COLS_USED)
    if 'Created' in calls.columns:
        calls['Created'] = calls['Created'].map(_serial_to_ts)

    chats = _worksheet_to_df_cols(sh.worksheet('Chats'), CHATS_COLS_USED)
    for c in ('DateTime Conversation Started', 'DateTime Conversation Resolved'):
        if c in chats.columns:
            chats[c] = chats[c].map(_serial_to_ts)
    if 'Assignee' in chats.columns:
        chats['Assignee'] = pd.to_numeric(chats['Assignee'], errors='coerce')
    if 'Contact ID' in chats.columns:
        chats['Contact ID'] = pd.to_numeric(chats['Contact ID'], errors='coerce')

    activity = _worksheet_to_df(sh.worksheet('Agents Activity'))
    if 'Timestamp' in activity.columns:
        activity['Timestamp'] = activity['Timestamp'].map(_serial_to_ts)
        # Set 'ts' directly here (bypassing fix_activity_ts) -- _serial_to_ts already
        # gives an unambiguous, correctly-parsed timestamp, unlike the xlsx export's
        # mixed text/Excel-auto-converted Timestamp column. build_roster() only falls
        # back to running fix_activity_ts itself when 'ts' isn't already present.
        activity['ts'] = activity['Timestamp']

    return agents, schedule, calls, chats, activity


def load_from_sheet(gc, spreadsheet_id):
    agents, schedule, calls, chats, activity = fetch_sheet_data(gc, spreadsheet_id)
    return build_roster(agents, schedule, calls, chats, activity)


def _in_range(series, start, end):
    return (series >= start) & (series <= pd.Timestamp(end) + pd.Timedelta(hours=23, minutes=59, seconds=59))


# ---------------------------------------------------------------------------
# Adherence (per agent, per scheduled day)
# ---------------------------------------------------------------------------
def shift_window(shift_label, day_dt):
    if shift_label == '9 AM - 5 PM':
        return day_dt + pd.Timedelta(hours=9), day_dt + pd.Timedelta(hours=17)
    if shift_label == '11 AM - 7 PM':
        return day_dt + pd.Timedelta(hours=11), day_dt + pd.Timedelta(hours=19)
    if shift_label == '4 PM - 12 AM':
        return day_dt + pd.Timedelta(hours=16), day_dt + pd.Timedelta(hours=24)
    if shift_label == '12 AM - 9 AM':
        # Overnight Convention (verified against the old manually-built report's
        # Methodology sheet, Sep 2026): a "12 AM - 9 AM" shift dated D on the
        # Schedule actually runs from 12:00 AM to 9:00 AM on D+1, not on D itself
        # -- Maqsam logs the login under the calendar day it actually happens,
        # which for a midnight-start shift is always the NEXT day relative to the
        # Schedule's "Date" cell. Anchoring the window to D instead (the old,
        # buggy behavior) queried an empty stretch of D 00:00-09:00 and silently
        # orphaned the agent's real D+1 00:00-09:00 session -- misattributing the
        # first day of an overnight streak as ~0 Actual Minutes, and dropping the
        # true activity of the streak's last day entirely if D+1 wasn't itself
        # scheduled as a working day. The row is still labeled/reported under
        # the original Schedule date D (see 'Date': day_dt below) -- only the
        # window used to pull real activity is shifted to D+1.
        next_day = day_dt + pd.Timedelta(days=1)
        return next_day, next_day + pd.Timedelta(hours=9)
    return None, None


def compute_adherence(data, start, end):
    schedule = data['schedule']
    activity = data['activity']
    chats = data['chats']
    roster = data['full_roster']
    gender_lookup = data.get('gender', {})

    sched_win = schedule[_in_range(schedule['Date'], start, end)].copy()
    act_win = activity[activity['ts'].notna()].copy()
    if act_win.empty:
        cutoff = pd.Timestamp(end) + pd.Timedelta(hours=23, minutes=59)
    else:
        cutoff = act_win['ts'].max()

    agent_intervals = {}
    for agent, grp in act_win.groupby('canonical'):
        grp = grp.sort_values('ts').reset_index(drop=True)
        intervals = []
        for i in range(len(grp)):
            s = grp.loc[i, 'ts']
            state = grp.loc[i, 'State']
            e = grp.loc[i + 1, 'ts'] if i + 1 < len(grp) else cutoff
            if e > s:
                intervals.append([s, e, state])
        agent_intervals[agent] = intervals

    # Occupancy needs to see chat-handling time too, not just call "Busy" state --
    # per Mahmoud (Sep 2026), Busy only fires for calls, so an agent working chats
    # all day still reads as "Available" and Occupancy comes out implausibly low.
    # Approximated as [DateTime Conversation Started, DateTime Conversation
    # Resolved] per conversation assigned to the agent -- the closest thing to an
    # "I was handling this" window the Chats export actually has -- but capped per
    # conversation (see CHAT_MIN_PER_MESSAGE / CHAT_MAX_CREDIT_MINUTES above): a
    # conversation only credits min(its own real duration, messages * 10 min, 60
    # min), anchored at its own start, since a long-open conversation is not proof
    # of continuous work the whole time it stayed open. This feeds BOTH the
    # chat-backed Offline fill below and Occupancy's Chat Occupied Minutes, so
    # fixing it here cleans both at once. Two conversations worked at once (or
    # whose capped credit windows overlap) are merged into one stretch first so
    # simultaneous/overlapping chats don't double-count.
    chat_intervals = {}
    chats_scoped = chats[chats['canonical'].isin(roster)]
    for agent, grp in chats_scoped.groupby('canonical'):
        ivs = []
        for s, e, n_in, n_out in zip(
            grp['DateTime Conversation Started'], grp['DateTime Conversation Resolved'],
            grp['Number of Incoming Messages'], grp['Number of Outgoing Messages'],
        ):
            if pd.isna(s) or pd.isna(e) or e <= s:
                continue
            n_msgs = (n_in or 0) + (n_out or 0)
            cap_min = min(n_msgs * CHAT_MIN_PER_MESSAGE, CHAT_MAX_CREDIT_MINUTES)
            credited_end = min(e, s + pd.Timedelta(minutes=cap_min))
            if credited_end > s:
                ivs.append((s, credited_end))
        chat_intervals[agent] = _merge_intervals(ivs)

    def clip(intervals, ws, we):
        out = []
        for s, e, st in intervals:
            cs, ce = max(s, ws), min(e, we)
            if ce > cs:
                out.append([cs, ce, st])
        return out

    daily_rows = []
    unclassified_days = []
    for agent in roster:
        a_sched = sched_win[sched_win['canonical'] == agent].sort_values('Date')
        if a_sched.empty:
            continue
        intervals = agent_intervals.get(agent, [])

        for _, srow in a_sched.iterrows():
            day_dt = srow['Date']
            shift_label = srow['Shift']
            is_wfh = bool(srow['is_wfh_flag'])
            wfh_raw = srow.get('wfh_raw', '')

            if shift_label in UNRESOLVED_LABELS:
                unclassified_days.append({'Agent': agent, 'Date': day_dt, 'Shift': shift_label})
                continue
            if shift_label in NON_WORKING or shift_label not in SHIFT_PLANNED_MIN:
                daily_rows.append({
                    'Agent': agent, 'Date': day_dt, 'Shift': shift_label,
                    'Working Day': False, 'Planned Minutes': 0, 'Actual Minutes': 0,
                    'Adherent Minutes': 0, 'Scheduled Minutes': 0,
                    'Late Minutes': 0, 'Early Logout Minutes': 0, 'Break Minutes': 0,
                    'Data Status': 'Complete', 'WFH Overridden': False,
                    'Excuse Minutes': 0, 'Excuse Note': '',
                })
                continue

            # Fri/Sat/Sun WFH default -- see WFH_OVERRIDE_WEEKDAYS above. REVISED Sep 16
            # 2026, per Mahmoud: only defaults to WFH when the Schedule's Is WFH cell was
            # actually blank for that day -- an explicit "No" is now respected, not
            # overridden (the old version forced WFH=Yes on a weekend even over an
            # explicit "No"). `wfh_raw` is '' for blank/missing, distinct from 'no'. Only
            # reached for a genuine working shift (Day Off/leave already excluded above).
            wfh_overridden = False
            if not is_wfh and day_dt.day_name() in WFH_OVERRIDE_WEEKDAYS and wfh_raw == '':
                is_wfh = True
                wfh_overridden = True
            # Standing shift-based WFH overrides, per Mahmoud (Sep 15 2026) -- see
            # WFH_ALWAYS_SHIFT_* above. Independent of the weekday override above (either
            # can flip it; wfh_overridden just records that SOME override fired when the
            # raw sheet said No).
            if not is_wfh and shift_label in WFH_ALWAYS_SHIFT_ALL:
                is_wfh = True
                wfh_overridden = True
            # Gender-driven, not a hardcoded name list -- revised Sep 16 2026, per Mahmoud,
            # who added a live "Gender" column to the Agents ID tab specifically so this
            # rule reads from the sheet instead of needing a code change per new hire. An
            # agent missing from gender_lookup entirely (off-roster, not on Agents ID at
            # all) falls back to the same "not male" default the old hardcoded list
            # already implied for anyone it didn't explicitly name.
            if not is_wfh and shift_label in WFH_ALWAYS_SHIFT_NON_MALE and gender_lookup.get(agent) != 'male':
                is_wfh = True
                wfh_overridden = True

            win_start, win_end = shift_window(shift_label, day_dt)
            planned = SHIFT_PLANNED_MIN[shift_label]
            # WFH -> 9-hour rule, per Mahmoud (Sep 2026): purely driven by that day's Is WFH
            # flag now (the old "female agent on 4 PM-12 AM = always WFH" special case is
            # gone -- no real basis for it was ever found, see module docstring). WFH = 9
            # hours (480 planned), not WFH = 8 hours (420, the SHIFT_PLANNED_MIN default
            # already set above) -- for any of the three 8-hour-type shifts, not just 4 PM-12 AM.
            nine_hour = False
            if shift_label in EIGHT_HOUR_TYPES and is_wfh:
                nine_hour = True
                win_end = win_end + pd.Timedelta(hours=1)
                planned = 480

            # Excuse Minutes, per Mahmoud (Sep 15 2026) -- an approved excuse for this
            # agent+day reduces their obligation for the day, same methodology already
            # reconciled manually for Aug 2026 (reused directly, not re-derived): approved
            # excuses reduce PLANNED minutes, they are not added to Actual. Floored at 0 so
            # a bad/oversized manual entry can't flip Planned negative. This intentionally
            # runs AFTER the WFH nine_hour bump above, so an excuse on a WFH day reduces
            # from 480, not the un-bumped 420/8-hour figure. (Sep 16 2026: a same-day cap
            # at the real Actual-vs-Planned shortfall was tried and then reverted, per
            # Mahmoud -- back to reducing Planned by the excuse's full stated value.)
            excuse_min = float(srow.get('excuse_minutes', 0.0) or 0.0)
            _raw_note = srow.get('excuse_note', '')
            excuse_note = '' if pd.isna(_raw_note) else str(_raw_note).strip()
            planned = max(0.0, planned - excuse_min)

            # Scheduled Minutes, added Sep 19 2026 per Mahmoud -- the DENOMINATOR for the
            # redesigned Adherence % (see 'Adherent Minutes' below), deliberately separate
            # from Planned Minutes above. Planned Minutes has the standing 60-minute break
            # already subtracted out (so Actual Minutes, which also excludes Break-state
            # time, compares apples to apples against it for Shrinkage % etc.) -- but
            # Adherent Minutes is a plain login-to-logout SPAN, which naturally includes
            # any break time that happened in the middle of it. Comparing that span against
            # the break-reduced Planned Minutes was mismatched by design (a fully-present,
            # perfectly on-time agent would already read ~114%, capped at 100% for
            # everyone) -- confirmed on real data (all 13 agents hit exactly 100.0%).
            # Scheduled Minutes is the raw shift window span instead (win_end - win_start,
            # already reflecting the WFH nine-hour bump above), reduced only by the same
            # Excuse Minutes -- matching the plain textbook Adherence definition (Scheduled
            # Shift Duration in minutes, no break subtracted either side).
            scheduled_min = max(0.0, (win_end - win_start).total_seconds() / 60 - excuse_min)

            data_status = 'Complete' if win_end <= cutoff else 'Incomplete*'
            window_clip_end = min(win_end + pd.Timedelta(hours=12), win_end + pd.Timedelta(hours=1))
            clipped = clip(intervals, win_start, window_clip_end)

            # Chat conversations [Started, Resolved] clipped to this shift window --
            # needed BELOW (chat-backed presence) as well as further down (Occupancy).
            agent_chat_ivs = chat_intervals.get(agent, [])
            chat_clipped = [
                (max(s, win_start), min(e, window_clip_end))
                for s, e in agent_chat_ivs
                if min(e, window_clip_end) > max(s, win_start)
            ]

            # Fill Offline gaps with real, evidenced chat-handling time BEFORE the
            # short-gap reclassification below -- see _fill_offline_with_chat.
            # Confirmed Sep 2026, per Mahmoud: Sara Hussien 17-18 Aug and Muhammed
            # Hesham 13 Aug both had multi-hour Agents Activity gaps that were
            # actually full shifts of real chat work, not real absence.
            clipped = _fill_offline_with_chat(clipped, chat_clipped, CHAT_BACKED_STATE)
            chat_inferred_min = sum(
                (e - s).total_seconds() / 60 for s, e, st in clipped if st == CHAT_BACKED_STATE
            )

            # Reclassify short Offline gaps sandwiched between identical break/
            # coaching/training states (e.g. a brief reconnect blip mid-break).
            # Capped at RECLASS_MAX_GAP_MINUTES -- Sep 2026, per Mahmoud, after a
            # real case (Muhammed Hesham, 13 Aug) bridged a ~5-hour Offline gap
            # into "Away - coaching" purely because a few-second coaching blip
            # happened to sit on each side of it. That gap is now mostly covered
            # by the chat fill above; this cap guards whatever's left.
            for i in range(1, len(clipped) - 1):
                if clipped[i][2] == OFFLINE_STATE:
                    prev_s, next_s = clipped[i - 1][2], clipped[i + 1][2]
                    gap_min = (clipped[i][1] - clipped[i][0]).total_seconds() / 60
                    if prev_s == next_s and prev_s in RECLASS_STATES and gap_min <= RECLASS_MAX_GAP_MINUTES:
                        clipped[i][2] = prev_s

            def mins(states):
                return sum((e - s).total_seconds() / 60 for s, e, st in clipped if st in states)

            online_min = mins({'Available', CHAT_BACKED_STATE})
            busy_min = mins({'Busy'})
            break_min = mins(BREAK_STATES)
            coaching_min = mins({COACHING_STATE})
            training_min = mins({TRAINING_STATE})
            tech_min = mins({TECH_STATE})
            actual = online_min + busy_min + coaching_min + training_min + tech_min

            # Chat time that fell inside an "Available" (or chat-backed) stretch --
            # NOT inside "Busy", since that's already call time and would
            # double-count. This is a reclassification within online_min, not extra
            # minutes, so Actual Minutes above is unaffected. A CHAT_BACKED_STATE
            # segment is, by construction, 100% chat-covered already.
            avail_intervals = [(s, e) for s, e, st in clipped if st in ('Available', CHAT_BACKED_STATE)]
            chat_occupied_min = _overlap_minutes(avail_intervals, chat_clipped)

            active = [iv for iv in clipped if iv[2] != OFFLINE_STATE]
            login = active[0][0] if active else None
            logout = active[-1][1] if active else None
            late_min = max(0, (login - win_start).total_seconds() / 60) if login else planned
            early_min = max(0, (win_end - logout).total_seconds() / 60) if logout else planned

            # Adherence, REDESIGNED Sep 19 2026 per Mahmoud -- no longer Actual
            # Minutes / Planned Minutes (that summed every state minute, including
            # chat-backed fill and reclassified gaps, so it could -- and after the
            # chat-fill fix, routinely did -- exceed Planned, making Adherence
            # meaningless for most agents). Adherence is now the plain WFM-standard
            # definition: how much of the scheduled window falls between actual
            # login and actual logout, clipped to the window itself so arriving
            # early or leaving late (overtime) earns no extra credit. It deliberately
            # does NOT re-check what state the agent was in minute-by-minute inside
            # that span (Break vs Available etc.) -- that's already Occupancy's job
            # (see 'Occupancy %' below), and doubling it into Adherence too was
            # solving the same problem twice. login/logout above already come from
            # the SAME clipped, chat-fill-capped intervals as everything else in
            # this loop, so a stale open chat can no longer make a no-show day read
            # as a full on-time shift (see CHAT_MIN_PER_MESSAGE above).
            if login and logout:
                adherent_min = max(0.0, (min(logout, win_end) - max(login, win_start)).total_seconds() / 60)
            else:
                adherent_min = 0.0

            daily_rows.append({
                'Agent': agent, 'Date': day_dt, 'Shift': shift_label,
                'Working Day': True, 'Nine Hour': nine_hour, 'Is WFH': is_wfh,
                'WFH Overridden': wfh_overridden,
                'Planned Minutes': planned, 'Actual Minutes': round(actual, 1),
                # Login-to-logout span, clipped to the scheduled window -- see the
                # comment above. This, not Actual Minutes, is what Adherence % is
                # built from now, against Scheduled Minutes (not Planned Minutes).
                'Adherent Minutes': round(adherent_min, 1),
                'Scheduled Minutes': round(scheduled_min, 1),
                'Online Minutes': round(online_min, 1), 'Busy Minutes': round(busy_min, 1),
                'Chat Occupied Minutes': round(chat_occupied_min, 1),
                # Of Online Minutes above, how much came from chat evidence filling an
                # Agents Activity gap rather than a real Maqsam "Available" state --
                # see _fill_offline_with_chat. 0 on a normal day; the whole point is
                # this stays visible instead of silently blending into Online Minutes.
                'Chat-Inferred Minutes': round(chat_inferred_min, 1),
                'Break Minutes': round(break_min, 1), 'Coaching Minutes': round(coaching_min, 1),
                'Training Minutes': round(training_min, 1), 'Technical Minutes': round(tech_min, 1),
                'Late Minutes': round(late_min, 1), 'Early Logout Minutes': round(early_min, 1),
                'Excuse Minutes': round(excuse_min, 1), 'Excuse Note': excuse_note,
                'Data Status': data_status,
            })

    daily_df = pd.DataFrame(daily_rows)
    unclassified_df = pd.DataFrame(unclassified_days)

    rows = []
    for agent in roster:
        a = daily_df[daily_df['Agent'] == agent] if not daily_df.empty else pd.DataFrame()
        if a.empty:
            continue
        working = a[a['Working Day']]
        complete = working[working['Data Status'] == 'Complete']
        incomplete_n = int((working['Data Status'] == 'Incomplete*').sum())
        planned = working['Planned Minutes'].sum()
        actual = complete['Actual Minutes'].sum()
        adherent = complete.get('Adherent Minutes', pd.Series(dtype=float)).sum()
        scheduled = complete.get('Scheduled Minutes', pd.Series(dtype=float)).sum()
        online = complete.get('Online Minutes', pd.Series(dtype=float)).sum()
        busy = complete.get('Busy Minutes', pd.Series(dtype=float)).sum()
        chat_occ = complete.get('Chat Occupied Minutes', pd.Series(dtype=float)).sum()
        occupied = busy + chat_occ
        reachable = online + busy
        occupancy = (occupied / reachable * 100) if reachable else np.nan
        shrink_min = (complete.get('Break Minutes', pd.Series(dtype=float)).sum()
                      + complete.get('Coaching Minutes', pd.Series(dtype=float)).sum()
                      + complete.get('Training Minutes', pd.Series(dtype=float)).sum()
                      + complete.get('Technical Minutes', pd.Series(dtype=float)).sum())
        planned_complete = complete['Planned Minutes'].sum()
        shrinkage = (shrink_min / planned_complete * 100) if planned_complete else np.nan
        # Adherence %, REDESIGNED Sep 19 2026 per Mahmoud -- Adherent Minutes (a
        # login-to-logout span, clipped to the scheduled window) / Scheduled Minutes
        # (the raw window span, not the break-reduced Planned Minutes -- see the
        # 'Scheduled Minutes' comment in the loop above for why those two can't be
        # mixed). The min(100, ...) below is doing real work, not just a formality:
        # Adherent Minutes is clipped to the raw window (win_start/win_end), but
        # Scheduled Minutes is that same window MINUS Excuse Minutes -- so on a day
        # with an approved excuse where the agent stayed present anyway (e.g. Sara
        # Hussien working through a Maqsam outage that was excused as Technical
        # Issue), Adherent Minutes can legitimately be a few minutes more than
        # Scheduled Minutes. That's the correct outcome (full credit, not >100%),
        # not a bug -- confirmed on real Aug 2026 data, 22 such day-rows, all with
        # a real Excuse Minutes entry behind them.
        adherence_pct = min(100.0, adherent / scheduled * 100) if scheduled else np.nan
        rows.append({
            'Agent': agent, 'Team': TEAM_OVERRIDE.get(agent, 'CS'),
            'Working Days': int(len(working)), 'Days Off': int((~a['Working Day']).sum()),
            'Planned Minutes': round(planned, 0), 'Actual Minutes': round(actual, 0),
            'Scheduled Minutes': round(scheduled, 0), 'Adherent Minutes': round(adherent, 0),
            'Adherence %': round(adherence_pct, 1) if pd.notna(adherence_pct) else None,
            'Late Minutes': round(complete['Late Minutes'].sum(), 0),
            'Early Logout Minutes': round(complete['Early Logout Minutes'].sum(), 0),
            # Total approved-excuse minutes this period, already subtracted out of
            # Planned Minutes above (Sep 15 2026, per Mahmoud) -- shown here too so the
            # size of that adjustment stays visible next to the number it changed.
            'Excuse Minutes': round(working['Excuse Minutes'].sum(), 0),
            # The raw components behind Occupancy % (now (Busy + Chat) / (Available +
            # Busy), not just call-Busy -- per Mahmoud (Sep 2026), Busy only fires for
            # calls, so an agent who's mostly on chats read as "Available" nearly all
            # day and Occupancy came out implausibly low even though she was working.
            # All three shown separately, not just the ratio, so a number that looks
            # off can be traced to WHICH side is unexpected straight from this table.
            'Available Minutes': round(online, 0), 'Busy Minutes (Calls)': round(busy, 0),
            'Chat Minutes': round(chat_occ, 0),
            'Occupancy %': round(occupancy, 1) if pd.notna(occupancy) else None,
            'Shrinkage %': round(shrinkage, 1) if pd.notna(shrinkage) else None,
            'Incomplete Days': incomplete_n,
            # Of Available Minutes above, how much is chat-evidenced fill rather than
            # a real Maqsam login -- see the per-day column of the same name. A large
            # number here is worth a manual look (agent not toggling Maqsam correctly,
            # or a real presence-tracking gap), not proof of anything wrong on its own.
            'Chat-Inferred Minutes': round(working.get('Chat-Inferred Minutes', pd.Series(dtype=float)).sum(), 0),
        })
    adherence_cols = ['Agent', 'Team', 'Working Days', 'Days Off', 'Planned Minutes', 'Actual Minutes',
                       'Scheduled Minutes', 'Adherent Minutes', 'Adherence %', 'Late Minutes',
                       'Early Logout Minutes', 'Excuse Minutes', 'Available Minutes',
                       'Busy Minutes (Calls)', 'Chat Minutes', 'Occupancy %', 'Shrinkage %',
                       'Incomplete Days', 'Chat-Inferred Minutes']
    adherence_df = pd.DataFrame(rows) if rows else pd.DataFrame(columns=adherence_cols)
    return adherence_df, daily_df, unclassified_df


# ---------------------------------------------------------------------------
# Chats
# ---------------------------------------------------------------------------
def compute_fcr(chats_win):
    """Per resolved conversation: FCR = the same Contact ID does not start another
    conversation within FCR_WINDOW_DAYS after this one's resolution. Computed within the
    loaded/filtered Chats data only (a contact's very next conversation might fall outside
    the file's date range and won't be seen)."""
    resolved = chats_win[chats_win['DateTime Conversation Resolved'].notna()].copy()
    resolved = resolved.sort_values(['Contact ID', 'DateTime Conversation Started'])
    fcr_flags = []
    for contact_id, grp in resolved.groupby('Contact ID'):
        grp = grp.sort_values('DateTime Conversation Started').reset_index()
        starts = grp['DateTime Conversation Started'].tolist()
        resolves = grp['DateTime Conversation Resolved'].tolist()
        for i in range(len(grp)):
            has_repeat = False
            if i + 1 < len(grp):
                gap = (starts[i + 1] - resolves[i]).total_seconds() / 86400
                has_repeat = gap <= FCR_WINDOW_DAYS
            fcr_flags.append((grp.loc[i, 'index'], not has_repeat))
    fcr_series = pd.Series({idx: val for idx, val in fcr_flags})
    resolved['FCR'] = resolved.index.map(fcr_series)
    return resolved


def compute_chats(data, start, end, denom_days):
    chats = data['chats']
    roster = data['full_roster']
    win = chats[_in_range(chats['DateTime Conversation Started'], start, end)].copy()
    win_scoped = win[win['canonical'].isin(roster)].copy()
    fcr_df = compute_fcr(win_scoped)

    # Opportunistic richer KPIs -- only switch on if the live Chats tab actually has
    # these columns (see build_roster). Kept separate from the required columns so a
    # sheet without them just gets the original, smaller table -- no crash either way.
    has_frt = 'First Response Time (td)' in win_scoped.columns
    has_restime = 'Resolution Time (td)' in win_scoped.columns
    has_atfr = 'First Assignment to First Response Time (td)' in win_scoped.columns
    has_category = 'Conversation Category' in win_scoped.columns

    rows = []
    for agent in roster:
        a = win_scoped[win_scoped['canonical'] == agent]
        if a.empty:
            continue
        closed = a[a['DateTime Conversation Resolved'].notna()]
        unique_contacts = closed['Contact ID'].nunique()
        a_fcr = fcr_df[fcr_df['canonical'] == agent] if not fcr_df.empty else pd.DataFrame()
        fcr_rate = (a_fcr['FCR'].mean() * 100) if not a_fcr.empty else None
        days = denom_days.get(agent)
        row = {
            'Agent': agent, 'Team': TEAM_OVERRIDE.get(agent, 'CS'),
            'Assigned': int(len(a)), 'Closed': int(len(closed)),
            'Unique Contacts': int(unique_contacts),
            'Avg per day': round(len(closed) / days, 2) if days else None,
            'FCR %': round(fcr_rate, 1) if fcr_rate is not None else None,
        }
        if has_frt:
            row['Avg First Response Time'] = fmt_td(_mean_td_nonblank(a, 'First Response Time', 'First Response Time (td)'))
        if has_atfr:
            row['Avg First Assignment to First Response'] = fmt_td(
                _mean_td_nonblank(a, 'First Assignment to First Response Time', 'First Assignment to First Response Time (td)'))
        if has_restime:
            row['Avg Resolution Time'] = fmt_td(_mean_td_nonblank(closed, 'Resolution Time', 'Resolution Time (td)'))
        rows.append(row)
    chats_cols = ['Agent', 'Team', 'Assigned', 'Closed', 'Unique Contacts', 'Avg per day', 'FCR %']
    if has_frt:
        chats_cols.append('Avg First Response Time')
    if has_atfr:
        chats_cols.append('Avg First Assignment to First Response')
    if has_restime:
        chats_cols.append('Avg Resolution Time')
    chats_df = (pd.DataFrame(rows).sort_values('Closed', ascending=False).reset_index(drop=True)
                if rows else pd.DataFrame(columns=chats_cols))

    unmatched = win[win['canonical'].isna() | ~win['canonical'].isin(roster)]
    unmatched_ids = (unmatched['Assignee'].dropna().astype(int).value_counts()
                     if not unmatched.empty else pd.Series(dtype=int))

    category_totals = None
    if has_category:
        vc = win_scoped['Conversation Category'].dropna()
        vc = vc[vc.astype(str).str.strip() != '']
        if not vc.empty:
            category_totals = vc.value_counts().head(12).to_dict()

    return chats_df, unmatched_ids, category_totals


# ---------------------------------------------------------------------------
# Calls
# ---------------------------------------------------------------------------
def compute_calls(data, start, end, denom_days, direction=None):
    calls = data['calls']
    roster = data['full_roster']
    win = calls[_in_range(calls['Created'], start, end)].copy()
    # Inbound/Outbound split, added Sep 2026 per Mahmoud -- direction=None (the default,
    # used for the combined Overall cards and the AHT card) keeps every call; pass
    # 'Inbound' or 'Outbound' to scope everything below (per-agent table AND the
    # company-wide totals dict) to just that direction. Same function either way, so the
    # two split blocks in the app can never drift from the logic the combined view uses.
    if direction is not None:
        win = win[win['Direction'] == direction]
    win['handle_td'] = win['Handling Duration'].map(to_timedelta)
    # Holding Duration, added Sep 20 2026 per Mahmoud -- how long an answered call sat on
    # hold, same "opportunistic" treatment as elsewhere in this file: missing on an older
    # export just reads as 0 holding time everywhere below, no crash. Averaged over the
    # same Serviced-calls scope as Avg Handling Time (a call that was never answered has
    # no real hold time to average in either).
    if 'Holding Duration' in win.columns:
        win['hold_td'] = win['Holding Duration'].map(to_timedelta)
    else:
        win['hold_td'] = pd.Timedelta(0)
    win_scoped = win[win['canonical'].isin(roster)].copy()

    rows = []
    for agent in roster:
        a = win_scoped[win_scoped['canonical'] == agent]
        if a.empty:
            continue
        total = len(a)
        state_counts = {s: int((a['State'] == s).sum()) for s in CALL_STATES}
        answered = sum(state_counts[s] for s in CALL_ANSWERED_STATES)
        answered_calls = a[a['State'].isin(CALL_ANSWERED_STATES)]
        avg_handle = answered_calls['handle_td'].mean() if len(answered_calls) else pd.Timedelta(0)
        avg_hold = answered_calls['hold_td'].mean() if len(answered_calls) else pd.Timedelta(0)
        days = denom_days.get(agent)
        row = {
            'Agent': agent, 'Team': TEAM_OVERRIDE.get(agent, 'CS'),
            'Total Calls': total, 'Answered %': round(answered / total * 100, 1) if total else None,
            'Avg per day': round(total / days, 2) if days else None,
            'Avg Handling Time': fmt_td(avg_handle), 'Avg Holding Time': fmt_td(avg_hold),
        }
        row.update(state_counts)
        rows.append(row)
    calls_cols = (['Agent', 'Team', 'Total Calls', 'Answered %', 'Avg per day',
                    'Avg Handling Time', 'Avg Holding Time'] + CALL_STATES)
    calls_df = (pd.DataFrame(rows).sort_values('Total Calls', ascending=False).reset_index(drop=True)
                if rows else pd.DataFrame(columns=calls_cols))

    # Company-wide totals from EVERY call in the period, not just the ones a single
    # agent can be credited with. A call that never reached anyone -- essentially all
    # Dropped calls, and some Abandoned ones -- has no Agent value at all, so it can
    # never appear in the per-agent table above (correctly -- there's nobody to
    # attribute it to). But deriving the Overall "Calls by state" totals by summing
    # that per-agent table (the old approach) silently made those calls invisible
    # company-wide too, which is exactly the "Dropped always reads 0" bug Mahmoud
    # flagged -- the sheet has the Dropped rows, they just have a blank Agent column.
    state_totals_all = {s: int((win['State'] == s).sum()) for s in CALL_STATES}
    total_calls_all = int(len(win))
    answered_all = sum(state_totals_all[s] for s in CALL_ANSWERED_STATES)
    answered_rate_all = round(answered_all / total_calls_all * 100, 1) if total_calls_all else None
    unattributed = int((~win['canonical'].isin(roster)).sum())

    # Company-wide AHT (added Sep 2026, per Mahmoud -- new top-level card) -- same
    # "every call in the period, not just the per-agent table" scope as the other
    # _all figures above, and the same averaged-over-Serviced-calls-only definition
    # already used per-agent (see 'Avg Handling Time' in the per-agent loop above).
    # Avg Holding Time (added Sep 20 2026, per Mahmoud) computed the same way,
    # alongside it -- see 'Avg Holding Time' in the per-agent loop above.
    answered_mask_all = win['State'].isin(CALL_ANSWERED_STATES)
    avg_handle_all = win.loc[answered_mask_all, 'handle_td'].mean() if answered_mask_all.any() else None
    aht_all = fmt_td(avg_handle_all) if avg_handle_all is not None else None
    avg_hold_all = win.loc[answered_mask_all, 'hold_td'].mean() if answered_mask_all.any() else None
    aholdt_all = fmt_td(avg_hold_all) if avg_hold_all is not None else None

    calls_totals = {
        'state_totals': state_totals_all, 'total_calls': total_calls_all,
        'answered_rate': answered_rate_all, 'unattributed_calls': unattributed,
        'avg_handling_time': aht_all, 'avg_holding_time': aholdt_all,
    }
    return calls_df, calls_totals


# ---------------------------------------------------------------------------
# Overall roll-up (for the top metric cards)
# ---------------------------------------------------------------------------
def compute_overall(chats_df, calls_df, calls_totals, adherence_df):
    o = {}
    o['total_chats_closed'] = int(chats_df['Closed'].sum()) if not chats_df.empty else 0
    o['total_chats_assigned'] = int(chats_df['Assigned'].sum()) if not chats_df.empty else 0
    valid_fcr = chats_df['FCR %'].dropna()
    o['fcr_rate'] = round(valid_fcr.mean(), 1) if len(valid_fcr) else None

    # From calls_totals (the full period, unattributed calls included) rather than
    # summed off the per-agent table -- see the comment in compute_calls.
    o['total_calls'] = calls_totals['total_calls']
    o['answered_rate'] = calls_totals['answered_rate']
    o['call_state_totals'] = calls_totals['state_totals']
    o['unattributed_calls'] = calls_totals['unattributed_calls']
    # AHT card, added Sep 2026 per Mahmoud -- one company-wide number (not split by
    # Inbound/Outbound, unlike the Calls section below -- confirmed with Mahmoud).
    # Avg Holding Time added Sep 20 2026, same scope/treatment.
    o['avg_handling_time'] = calls_totals.get('avg_handling_time')
    o['avg_holding_time'] = calls_totals.get('avg_holding_time')
    state_totals = calls_totals['state_totals']
    o['dropped_rate'] = (round(state_totals['Dropped'] / o['total_calls'] * 100, 1)
                          if o['total_calls'] else None)
    o['abandoned_rate'] = (round(state_totals['Abandoned'] / o['total_calls'] * 100, 1)
                            if o['total_calls'] else None)

    o['agents_in_scope'] = int(len(set(chats_df['Agent']) | set(calls_df['Agent']) | set(adherence_df['Agent'] if not adherence_df.empty else [])))
    if not adherence_df.empty:
        o['avg_adherence'] = round(adherence_df['Adherence %'].dropna().mean(), 1) if adherence_df['Adherence %'].notna().any() else None
        o['avg_occupancy'] = round(adherence_df['Occupancy %'].dropna().mean(), 1) if adherence_df['Occupancy %'].notna().any() else None
        o['avg_shrinkage'] = round(adherence_df['Shrinkage %'].dropna().mean(), 1) if adherence_df['Shrinkage %'].notna().any() else None
        # Company-wide split of the "occupied" side of Occupancy % between calls and
        # chats -- Mahmoud wants to see, not just the blended ratio, how much of it
        # is calls vs. chats (surfaced as a popover on the Avg. Occupancy card).
        calls_min = adherence_df['Busy Minutes (Calls)'].sum()
        chat_min = adherence_df['Chat Minutes'].sum()
        occupied_total = calls_min + chat_min
        if occupied_total:
            o['occupancy_calls_min'] = round(calls_min, 0)
            o['occupancy_chat_min'] = round(chat_min, 0)
            o['occupancy_calls_pct'] = round(calls_min / occupied_total * 100, 1)
            o['occupancy_chat_pct'] = round(chat_min / occupied_total * 100, 1)
        else:
            o['occupancy_calls_min'] = o['occupancy_chat_min'] = 0
            o['occupancy_calls_pct'] = o['occupancy_chat_pct'] = None
    else:
        o['avg_adherence'] = o['avg_occupancy'] = o['avg_shrinkage'] = None
        o['occupancy_calls_min'] = o['occupancy_chat_min'] = 0
        o['occupancy_calls_pct'] = o['occupancy_chat_pct'] = None
    return o


def working_days_lookup(adherence_df):
    if adherence_df.empty:
        return {}
    return dict(zip(adherence_df['Agent'], adherence_df['Working Days']))


def build_report_from_data(data, start, end):
    adherence_df, daily_df, unclassified_df = compute_adherence(data, start, end)
    days_lookup = working_days_lookup(adherence_df)
    calendar_days = max(1, (pd.Timestamp(end) - pd.Timestamp(start)).days + 1)
    denom_days = {a: days_lookup.get(a, calendar_days) or calendar_days for a in data['full_roster']}
    chats_df, unmatched_chat_ids, chat_category_totals = compute_chats(data, start, end, denom_days)
    calls_df, calls_totals = compute_calls(data, start, end, denom_days)
    # Inbound/Outbound split, added Sep 2026 per Mahmoud -- both blocks always computed
    # and shown together in the app (no toggle), on top of the combined calls_df/
    # calls_totals above which stay as-is for the existing Overall cards/chart.
    calls_inbound_df, calls_inbound_totals = compute_calls(data, start, end, denom_days, direction='Inbound')
    calls_outbound_df, calls_outbound_totals = compute_calls(data, start, end, denom_days, direction='Outbound')
    overall = compute_overall(chats_df, calls_df, calls_totals, adherence_df)
    return {
        'data': data, 'overall': overall, 'chats': chats_df, 'calls': calls_df,
        'calls_inbound': calls_inbound_df, 'calls_inbound_totals': calls_inbound_totals,
        'calls_outbound': calls_outbound_df, 'calls_outbound_totals': calls_outbound_totals,
        'adherence': adherence_df, 'daily_audit': daily_df, 'unclassified_shifts': unclassified_df,
        'unmatched_chat_ids': unmatched_chat_ids, 'unattributed_calls': calls_totals['unattributed_calls'],
        'chat_category_totals': chat_category_totals,
    }


def build_report(file_obj, start, end):
    """From an uploaded xlsx workbook."""
    return build_report_from_data(load_workbook(file_obj), start, end)


def build_report_from_sheet(gc, spreadsheet_id, start, end):
    """From the live Google Sheet."""
    return build_report_from_data(load_from_sheet(gc, spreadsheet_id), start, end)


# ---------------------------------------------------------------------------
# Period comparison -- same concept as the existing Ops Pulse comparison report:
# pick two date ranges (Period A / Period B), see every metric side by side with
# a delta, both company-wide (Overall) and per agent. Operates on two already-
# built report dicts (from build_report_from_sheet/build_report), so it works
# the same way regardless of data source and needs no extra sheet access.
# ---------------------------------------------------------------------------
# (data key, label, unit, higher-is-better) -- Shrinkage is the one metric here where
# a NEGATIVE delta is the improvement, so the highlight narrative below needs this to
# classify "improved" vs "declined" correctly instead of just reading the delta's sign.
OVERALL_COMPARISON_METRICS = [
    ('total_chats_closed', 'Chats Closed', 'count', True),
    ('fcr_rate', 'Chats FCR Rate', 'pp', True),
    ('total_calls', 'Total Calls', 'count', True),
    ('answered_rate', 'Calls Answered Rate', 'pp', True),
    ('avg_adherence', 'Avg. Adherence', 'pp', True),
    ('avg_occupancy', 'Avg. Occupancy', 'pp', True),
    ('avg_shrinkage', 'Avg. Shrinkage', 'pp', False),
]


def _delta(a, b):
    if a is None or b is None or pd.isna(a) or pd.isna(b):
        return None
    return round(b - a, 2)


def compare_overall(overall_a, overall_b):
    rows = []
    for key, label, unit, higher_better in OVERALL_COMPARISON_METRICS:
        a, b = overall_a.get(key), overall_b.get(key)
        rows.append({'Metric': label, 'Period A': a, 'Period B': b, 'Delta': _delta(a, b),
                     'Unit': unit, 'Higher Is Better': higher_better})
    return pd.DataFrame(rows)


def _safe_select(df, cols):
    """df[['Agent', ...]], but tolerant of a completely empty/columnless DataFrame
    (an empty period) instead of raising a KeyError."""
    if df.empty or any(c not in df.columns for c in cols):
        return pd.DataFrame(columns=cols)
    return df[cols].copy()


def _compare_agents(df_a, df_b, metric_cols):
    cols = ['Agent'] + metric_cols
    a, b = _safe_select(df_a, cols), _safe_select(df_b, cols)
    merged = a.merge(b, on='Agent', how='outer', suffixes=(' (A)', ' (B)'))
    ordered = ['Agent']
    for col in metric_cols:
        ca, cb, cd = f'{col} (A)', f'{col} (B)', f'{col} Δ'
        merged[cd] = pd.to_numeric(merged[cb], errors='coerce') - pd.to_numeric(merged[ca], errors='coerce')
        ordered += [ca, cb, cd]
    return merged[ordered].sort_values('Agent').reset_index(drop=True)


def compare_chats(chats_a, chats_b):
    return _compare_agents(chats_a, chats_b, ['Closed', 'FCR %'])


def compare_calls(calls_a, calls_b):
    return _compare_agents(calls_a, calls_b, ['Total Calls', 'Answered %'])


def compare_adherence(adherence_a, adherence_b):
    return _compare_agents(adherence_a, adherence_b, ['Adherence %', 'Occupancy %'])


def build_highlights(overall_cmp, chats_cmp, calls_cmp, adherence_cmp, top_n=3):
    """Short "what changed most" narrative -- same spirit as Ops Pulse's Summary tab:
    biggest movers first, split into improved / declined. Only rate metrics (pp) are
    compared this way -- a raw count moving is just volume, not a rate improving or
    declining, so it's left to the tables rather than the highlight list.

    Each mover is (goodness, text): goodness is the delta re-signed so positive always
    means "this got better" -- Shrinkage is the one metric here where a falling number
    is the improvement, so its delta gets flipped before ranking; the text below always
    shows the real, unflipped delta so the number itself never lies."""
    movers = []
    for _, r in overall_cmp.iterrows():
        if r['Unit'] == 'pp' and r['Delta'] is not None:
            goodness = r['Delta'] if r['Higher Is Better'] else -r['Delta']
            movers.append((goodness, f"Company-wide: {r['Metric']} moved {r['Delta']:+.1f}pp "
                                      f"({r['Period A']:.1f}% -> {r['Period B']:.1f}%)."))
    for label, df, col in [
        ('Chats FCR', chats_cmp, 'FCR % Δ'), ('Calls Answered rate', calls_cmp, 'Answered % Δ'),
        ('Adherence', adherence_cmp, 'Adherence % Δ'), ('Occupancy', adherence_cmp, 'Occupancy % Δ'),
    ]:
        if col not in df.columns:
            continue
        for _, r in df.dropna(subset=[col]).iterrows():
            movers.append((r[col], f"{r['Agent']} -- {label} moved {r[col]:+.1f}pp."))
    movers.sort(key=lambda m: m[0])
    declined = [m[1] for m in movers if m[0] < 0][:top_n]
    improved = [m[1] for m in movers[::-1] if m[0] > 0][:top_n]
    return {'improved': improved, 'declined': declined}


def build_comparison(report_a, report_b):
    overall_cmp = compare_overall(report_a['overall'], report_b['overall'])
    chats_cmp = compare_chats(report_a['chats'], report_b['chats'])
    calls_cmp = compare_calls(report_a['calls'], report_b['calls'])
    adherence_cmp = compare_adherence(report_a['adherence'], report_b['adherence'])
    highlights = build_highlights(overall_cmp, chats_cmp, calls_cmp, adherence_cmp)
    return {
        'overall': overall_cmp, 'chats': chats_cmp, 'calls': calls_cmp,
        'adherence': adherence_cmp, 'highlights': highlights,
    }


# ---------------------------------------------------------------------------
# Excel export -- the whole report, or a hand-picked subset of sections, as one
# .xlsx with a sheet per section. Sheet names are kept under Excel's 31-char cap.
#
# Visual style matches the existing Ops Pulse comparison export (same workbook
# Mahmoud already has): a bold title + grey period/generated-at lines above each
# table, a dark-green (#1F4E3D) bold-white header row, percent-formatted rate
# columns, sized columns, a frozen header row, and a couple of embedded bar
# charts -- rather than a bare pandas dump.
# ---------------------------------------------------------------------------
_HEADER_FILL = PatternFill('solid', fgColor='1F4E3D')
_HEADER_FONT = Font(bold=True, color='FFFFFF')
_TITLE_FONT = Font(bold=True, size=13)
_SUBTITLE_FONT = Font(color='555555')
_PCT_FORMAT = '0.0"%"'  # values are already 0-100 scale in this codebase, not 0-1 fractions

# Column names (as they appear in each exported table) that should render with a
# "%" suffix instead of a bare number.
_PCT_COLS_BY_SHEET = {
    'Overall': set(),  # handled per-row below -- Value column mixes % and non-% metrics
    'Chats': {'FCR %'},
    'Calls': {'Answered %'},
    'Adherence': {'Adherence %', 'Occupancy %', 'Shrinkage %'},
    'AOV per Agent': set(),
}


def _write_titled_sheet(writer, df, sheet_name, title, subtitle, pct_cols=None):
    """Writes df starting a few rows down (leaving room for a title block), then
    styles the header row, sizes columns, freezes the header, and applies percent
    formatting to the given column names. Returns (worksheet, header_row, first_data_row,
    last_data_row) -- 1-indexed Excel rows -- for callers that want to add a chart."""
    startrow = 3  # 0-indexed -- header lands on Excel row 4
    df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=startrow)
    ws = writer.sheets[sheet_name]
    ws.cell(row=1, column=1, value=title).font = _TITLE_FONT
    ws.cell(row=2, column=1, value=subtitle).font = _SUBTITLE_FONT

    header_row = startrow + 1
    ncols = max(len(df.columns), 1)
    for col_idx, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(row=header_row, column=col_idx)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        content_width = max((len(str(v)) for v in df[col_name]), default=0) if len(df) else 0
        width = min(40, max(12, len(str(col_name)) + 4, content_width + 2))
        ws.column_dimensions[get_column_letter(col_idx)].width = width
        if pct_cols and col_name in pct_cols:
            for row_idx in range(header_row + 1, header_row + 1 + len(df)):
                ws.cell(row=row_idx, column=col_idx).number_format = _PCT_FORMAT
    ws.freeze_panes = ws.cell(row=header_row + 1, column=1)
    return ws, header_row, header_row + 1, header_row + len(df)


def _add_bar_chart(ws, title, cat_col, val_col, first_data_row, last_data_row, anchor_row):
    if last_data_row < first_data_row:
        return  # nothing to chart
    chart = BarChart()
    chart.type = 'bar'
    chart.title = title
    chart.y_axis.title = None
    chart.x_axis.title = None
    cats = Reference(ws, min_col=cat_col, min_row=first_data_row, max_row=last_data_row)
    vals = Reference(ws, min_col=val_col, min_row=first_data_row - 1, max_row=last_data_row)  # include header for series name
    chart.add_data(vals, titles_from_data=True)
    chart.set_categories(cats)
    chart.height, chart.width = 8, 16
    ws.add_chart(chart, f'A{anchor_row}')


def export_excel(result, comparison=None, sections=None, aov_df=None, aov_market_df=None,
                  period_label=None, currency_note=None):
    sections = sections or ['Overall', 'Chats', 'Calls', 'Adherence']
    generated = dt.datetime.now().strftime('%Y-%m-%d %H:%M')
    subtitle = f"{period_label or ''}    |    Generated: {generated}".strip(' |')

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        if 'Overall' in sections:
            overall = result['overall']
            scalar_rows = [{'Metric': k, 'Value': v} for k, v in overall.items()
                            if not isinstance(v, dict)]
            overall_df = pd.DataFrame(scalar_rows)
            ws, hdr, first, last = _write_titled_sheet(
                writer, overall_df, 'Overall', 'CS Dashboard -- Overall', subtitle)
            # Value column mixes % and non-% metrics -- format only the '_rate' rows.
            for row_idx, rec in zip(range(first, last + 1), scalar_rows):
                if 'rate' in str(rec['Metric']).lower():
                    ws.cell(row=row_idx, column=2).number_format = _PCT_FORMAT

            state_df = pd.DataFrame(list(overall['call_state_totals'].items()), columns=['State', 'Calls'])
            ws2, hdr2, first2, last2 = _write_titled_sheet(
                writer, state_df, 'Calls by State', 'CS Dashboard -- Calls by State', subtitle)
            _add_bar_chart(ws2, 'Calls by state', cat_col=1, val_col=2, first_data_row=first2,
                            last_data_row=last2, anchor_row=last2 + 3)

        if 'Chats' in sections:
            _write_titled_sheet(writer, result['chats'], 'Chats', 'CS Dashboard -- Chats per Agent',
                                 subtitle, pct_cols=_PCT_COLS_BY_SHEET['Chats'])
        if 'Calls' in sections:
            _write_titled_sheet(writer, result['calls'], 'Calls', 'CS Dashboard -- Calls per Agent',
                                 subtitle, pct_cols=_PCT_COLS_BY_SHEET['Calls'])
            # Inbound/Outbound split, added Sep 2026 per Mahmoud -- same per-agent shape
            # as the combined 'Calls' sheet above, just pre-filtered by Direction, so the
            # export mirrors the two always-shown blocks in the app.
            _write_titled_sheet(writer, result['calls_inbound'], 'Calls (Inbound)',
                                 'CS Dashboard -- Calls per Agent, Inbound', subtitle,
                                 pct_cols=_PCT_COLS_BY_SHEET['Calls'])
            _write_titled_sheet(writer, result['calls_outbound'], 'Calls (Outbound)',
                                 'CS Dashboard -- Calls per Agent, Outbound', subtitle,
                                 pct_cols=_PCT_COLS_BY_SHEET['Calls'])
        if 'Adherence' in sections:
            _write_titled_sheet(writer, result['adherence'], 'Adherence', 'CS Dashboard -- Adherence per Agent',
                                 subtitle, pct_cols=_PCT_COLS_BY_SHEET['Adherence'])
        if 'AOV' in sections and (aov_df is not None or aov_market_df is not None):
            aov_subtitle = subtitle + (f"    |    {currency_note}" if currency_note else '')
        if aov_df is not None and 'AOV' in sections:
            ws3, hdr3, first3, last3 = _write_titled_sheet(
                writer, aov_df, 'AOV per Agent', 'CS Dashboard -- AOV per Agent', aov_subtitle)
            value_col_name = 'Total Value (USD)' if 'Total Value (USD)' in aov_df.columns else 'Total Value'
            val_col_idx = list(aov_df.columns).index(value_col_name) + 1
            _add_bar_chart(ws3, f'Total order value by agent ({value_col_name})', cat_col=1,
                            val_col=val_col_idx, first_data_row=first3, last_data_row=last3,
                            anchor_row=last3 + 3)
        # Per-agent-per-market breakdown, added Sep 20 2026 per Mahmoud -- separate
        # sheet, same 'AOV' export section, since it's the table that carries the
        # (correct, per-market) 'vs AOV Target' badge -- see compute_aov_by_agent_market.
        if aov_market_df is not None and not aov_market_df.empty and 'AOV' in sections:
            _write_titled_sheet(
                writer, aov_market_df, 'AOV per Agent per Market',
                'CS Dashboard -- AOV per Agent per Market (vs CEO Target)', aov_subtitle)
        if comparison is not None and 'Comparison' in sections:
            _write_titled_sheet(writer, comparison['overall'], 'Comparison Overall',
                                 'CS Dashboard -- Comparison, Overall', subtitle)
            _write_titled_sheet(writer, comparison['chats'], 'Comparison Chats',
                                 'CS Dashboard -- Comparison, Chats', subtitle)
            _write_titled_sheet(writer, comparison['calls'], 'Comparison Calls',
                                 'CS Dashboard -- Comparison, Calls', subtitle)
            _write_titled_sheet(writer, comparison['adherence'], 'Comparison Adherence',
                                 'CS Dashboard -- Comparison, Adherence', subtitle)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# AOV per agent -- from the separate Orders "Clean" sheet's Orders tab, Salesman
# column (Mahmoud, Sep 2026). This is a DIFFERENT spreadsheet from the CS sheet --
# needs its own Viewer share for the same service account, see the README.
#
# Only the Salesman column's position/values were confirmed directly ("last column,
# either 'Created by customer' or the agent's name -- we want the agent name").
# The order-value, date and market columns were NOT confirmed, so this is written
# defensively: it guesses from a short list of likely names and returns a clear
# diagnostic (not a crash, not a silent wrong number) if it can't find them, so the
# UI can show exactly what went wrong instead of a raw traceback.
# ---------------------------------------------------------------------------
AOV_DATE_CANDIDATES = ['Created', 'Order Date', 'Date', 'Created At', 'Order Created']
AOV_VALUE_CANDIDATES = ['Subtotal', 'Order Value', 'Total', 'Amount', 'Value', 'Order Total']
AOV_MARKET_CANDIDATES = ['Country', 'Shipping Country', 'Market']
NOT_AGENT_SALESMAN = {'created by customer'}


def _find_column(columns, candidates):
    lower = {str(c).strip().lower(): c for c in columns}
    for cand in candidates:
        if cand.lower() in lower:
            return lower[cand.lower()]
    for c in columns:
        for cand in candidates:
            if cand.lower() in str(c).lower():
                return c
    return None


def load_orders_clean(gc, spreadsheet_id, tab_name='Orders'):
    sh = gc.open_by_key(spreadsheet_id)
    return _worksheet_to_df(sh.worksheet(tab_name))


def compute_aov_by_agent(orders_df, start, end, fx_rates=None):
    """Returns (aov_df, diagnostic). aov_df is None (with diagnostic explaining why)
    if the sheet's column names couldn't be confidently identified -- see the module
    docstring above this function for why that's the fallback here rather than a
    best-effort guess that could quietly be wrong.

    fx_rates, if given, is a {market_code: rate} dict where rate means "local
    currency units per 1 USD" (e.g. {'IQ': 1310} for 1,310 IQD = $1) -- entered by
    Mahmoud in the app's sidebar, never guessed here (exchange rates move and
    Ops Pulse itself doesn't convert currency, so there's no rate to inherit from
    it). Conversion happens PER ORDER, before aggregation -- an agent whose orders
    span more than one market/currency can't be correctly converted after the fact
    by dividing an already-mixed-currency sum by a single rate. Orders in a market
    with no rate supplied are simply left out of the USD columns (counted in the
    native-currency Orders/AOV/Total Value columns as before) -- never silently
    assigned somebody else's rate."""
    if orders_df.empty:
        return None, "The Orders tab came back empty."
    columns = list(orders_df.columns)
    salesman_col = 'Salesman' if 'Salesman' in columns else columns[-1]
    date_col = _find_column(columns, AOV_DATE_CANDIDATES)
    value_col = _find_column(columns, AOV_VALUE_CANDIDATES)
    market_col = _find_column(columns, AOV_MARKET_CANDIDATES)
    missing = [name for name, col in [('a date', date_col), ('an order-value', value_col)] if col is None]
    if missing:
        return None, (f"Couldn't confidently find {' or '.join(missing)} column. "
                       f"Columns actually seen in the Orders tab: {columns}")

    work = orders_df.copy()
    if work[date_col].map(lambda v: isinstance(v, (int, float)) and not isinstance(v, bool)).any():
        work[date_col] = work[date_col].map(_serial_to_ts)
    else:
        work[date_col] = pd.to_datetime(work[date_col], errors='coerce')
    work[value_col] = pd.to_numeric(work[value_col], errors='coerce')

    win = work[_in_range(work[date_col], start, end)]
    salesman_norm = win[salesman_col].astype(str).str.strip()
    agent_orders = win[~salesman_norm.str.lower().isin(NOT_AGENT_SALESMAN) & (salesman_norm != '') & win[value_col].notna()].copy()
    if agent_orders.empty:
        return pd.DataFrame(columns=['Agent', 'Orders', 'AOV', 'Total Value']), None

    fx_rates = {str(k).strip().upper(): v for k, v in (fx_rates or {}).items() if v}
    has_usd = bool(fx_rates) and market_col is not None
    if has_usd:
        rate = agent_orders[market_col].astype(str).str.strip().str.upper().map(fx_rates)
        agent_orders['_usd_value'] = agent_orders[value_col] / rate
        agent_orders['_unconverted'] = rate.isna()

    # One row per agent (this is "AOV per agent", not per agent-per-market) --
    # Market, if present, is folded in as an informational "Markets" column (the
    # distinct markets that agent's orders touched) rather than splitting the agent
    # across multiple rows, which would make the headline AOV number impossible to
    # read at a glance.
    grouped = agent_orders.groupby(salesman_col)[value_col].agg(['count', 'mean', 'sum']).reset_index()
    grouped = grouped.rename(columns={salesman_col: 'Agent', 'count': 'Orders', 'mean': 'AOV', 'sum': 'Total Value'})
    if market_col:
        markets = agent_orders.groupby(salesman_col)[market_col].apply(
            lambda s: ', '.join(sorted(set(str(v).strip() for v in s if str(v).strip())))
        ).reset_index().rename(columns={salesman_col: 'Agent', market_col: 'Markets'})
        grouped = grouped.merge(markets, on='Agent', how='left')

    if has_usd:
        converted = agent_orders[~agent_orders['_unconverted']]
        if not converted.empty:
            usd_agg = converted.groupby(salesman_col)['_usd_value'].agg(['count', 'mean', 'sum']).reset_index()
            usd_agg = usd_agg.rename(columns={
                salesman_col: 'Agent', 'count': 'Orders (converted)', 'mean': 'AOV (USD)', 'sum': 'Total Value (USD)',
            })
            grouped = grouped.merge(usd_agg, on='Agent', how='left')
        else:
            grouped['Orders (converted)'] = 0
            grouped['AOV (USD)'] = np.nan
            grouped['Total Value (USD)'] = np.nan
        unconverted_n = int(agent_orders['_unconverted'].sum())
        if unconverted_n:
            missing_markets = sorted(set(
                agent_orders.loc[agent_orders['_unconverted'], market_col].astype(str).str.strip().str.upper()
            ) - {''})
        else:
            missing_markets = []
        grouped['AOV (USD)'] = grouped['AOV (USD)'].round(2)
        grouped['Total Value (USD)'] = grouped['Total Value (USD)'].round(2)
        # No 'vs AOV Target' badge here any more, REMOVED Sep 20 2026 per Mahmoud --
        # this table is blended across every market an agent sold in (see the
        # 'Markets' column), but the CEO scorecard's AOV target is set PER MARKET
        # and varies up to 2x (UAE/OM $80 floor vs KW/QA/SA $90-100) -- no single
        # number here could ever compare honestly. See compute_aov_by_agent_market()
        # below for the per-agent-per-market breakdown that carries this badge.

    grouped['AOV'] = grouped['AOV'].round(2)
    grouped['Total Value'] = grouped['Total Value'].round(2)
    grouped = grouped.sort_values('Total Value', ascending=False).reset_index(drop=True)

    diagnostic = None
    if has_usd and unconverted_n:
        diagnostic = (f"⚠️ {unconverted_n:,} order(s) in market(s) without a rate set "
                       f"({', '.join(missing_markets)}) aren't included in the USD columns -- "
                       f"add a rate for {'them' if len(missing_markets) > 1 else 'it'} in the sidebar to include them.")
    return grouped, diagnostic


# ---------------------------------------------------------------------------
# AOV per agent, PER MARKET -- added Sep 20 2026 per Mahmoud. Same source/columns
# as compute_aov_by_agent above, just grouped by [Agent, Market] instead of Agent
# alone, so each cell can be checked against that market's own CEO target
# (AOV_MARKET_TARGETS) rather than one blended number that can't be right for
# every market an agent touches at once. See the 'vs AOV Target' removal note in
# compute_aov_by_agent for the full reasoning.
# ---------------------------------------------------------------------------
def compute_aov_by_agent_market(orders_df, start, end, fx_rates=None):
    """Returns (df, diagnostic), same None-with-explanation convention as
    compute_aov_by_agent. Requires a market/country column (compute_aov_by_agent
    doesn't) since that's the whole point of this table."""
    if orders_df.empty:
        return None, "The Orders tab came back empty."
    columns = list(orders_df.columns)
    salesman_col = 'Salesman' if 'Salesman' in columns else columns[-1]
    date_col = _find_column(columns, AOV_DATE_CANDIDATES)
    value_col = _find_column(columns, AOV_VALUE_CANDIDATES)
    market_col = _find_column(columns, AOV_MARKET_CANDIDATES)
    missing = [name for name, col in [('a date', date_col), ('an order-value', value_col)] if col is None]
    if missing:
        return None, (f"Couldn't confidently find {' or '.join(missing)} column. "
                       f"Columns actually seen in the Orders tab: {columns}")
    if market_col is None:
        return None, "Couldn't confidently find a market/country column, so a per-market breakdown isn't possible."

    work = orders_df.copy()
    if work[date_col].map(lambda v: isinstance(v, (int, float)) and not isinstance(v, bool)).any():
        work[date_col] = work[date_col].map(_serial_to_ts)
    else:
        work[date_col] = pd.to_datetime(work[date_col], errors='coerce')
    work[value_col] = pd.to_numeric(work[value_col], errors='coerce')

    win = work[_in_range(work[date_col], start, end)]
    salesman_norm = win[salesman_col].astype(str).str.strip()
    agent_orders = win[~salesman_norm.str.lower().isin(NOT_AGENT_SALESMAN) & (salesman_norm != '') & win[value_col].notna()].copy()
    if agent_orders.empty:
        return pd.DataFrame(columns=['Agent', 'Market', 'Orders', 'AOV', 'Total Value']), None
    agent_orders[market_col] = agent_orders[market_col].astype(str).str.strip().str.upper()

    grouped = agent_orders.groupby([salesman_col, market_col])[value_col].agg(['count', 'mean', 'sum']).reset_index()
    grouped = grouped.rename(columns={
        salesman_col: 'Agent', market_col: 'Market', 'count': 'Orders', 'mean': 'AOV', 'sum': 'Total Value',
    })
    grouped['AOV'] = grouped['AOV'].round(2)
    grouped['Total Value'] = grouped['Total Value'].round(2)

    fx_rates = {str(k).strip().upper(): v for k, v in (fx_rates or {}).items() if v}
    has_usd = bool(fx_rates)
    unconverted_n = 0
    missing_markets = []
    if has_usd:
        rate = agent_orders[market_col].map(fx_rates)
        agent_orders['_usd_value'] = agent_orders[value_col] / rate
        agent_orders['_unconverted'] = rate.isna()
        converted = agent_orders[~agent_orders['_unconverted']]
        if not converted.empty:
            usd_agg = converted.groupby([salesman_col, market_col])['_usd_value'].agg(['count', 'mean', 'sum']).reset_index()
            usd_agg = usd_agg.rename(columns={
                salesman_col: 'Agent', market_col: 'Market',
                'count': 'Orders (converted)', 'mean': 'AOV (USD)', 'sum': 'Total Value (USD)',
            })
            grouped = grouped.merge(usd_agg, on=['Agent', 'Market'], how='left')
        else:
            grouped['Orders (converted)'] = 0
            grouped['AOV (USD)'] = np.nan
            grouped['Total Value (USD)'] = np.nan
        grouped['Orders (converted)'] = grouped['Orders (converted)'].fillna(0).astype(int)
        grouped['AOV (USD)'] = grouped['AOV (USD)'].round(2)
        grouped['Total Value (USD)'] = grouped['Total Value (USD)'].round(2)
        grouped['vs AOV Target'] = grouped.apply(
            lambda r: aov_market_badge(r['AOV (USD)'], r['Market'], r['Orders (converted)']), axis=1
        )
        unconverted_n = int(agent_orders['_unconverted'].sum())
        if unconverted_n:
            missing_markets = sorted(set(
                agent_orders.loc[agent_orders['_unconverted'], market_col]
            ) - {''})

    grouped = grouped.sort_values(['Agent', 'Total Value'], ascending=[True, False]).reset_index(drop=True)

    diagnostic = None
    if has_usd and unconverted_n:
        diagnostic = (f"⚠️ {unconverted_n:,} order(s) in market(s) without a rate set "
                       f"({', '.join(missing_markets)}) aren't included in the USD columns -- "
                       f"add a rate for {'them' if len(missing_markets) > 1 else 'it'} in the sidebar to include them.")
    return grouped, diagnostic

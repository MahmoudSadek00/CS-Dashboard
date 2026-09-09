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
  - company policy: on the "4 PM - 12 AM" shift, female agents work from home and it is
    always counted as a 9-hour shift ending 1 AM, whether or not that day's "Is WFH" flag
    is set -- gender is inferred from each agent's first name (see FEMALE below)
  - the Agents Activity Timestamp column mixing plain-text and Excel-auto-converted
    datetime cells, which silently swaps day/month for the auto-converted ones unless
    corrected (fix_activity_ts)
  - Calls State always broken out into its own 8 individual states, never collapsed into
    a single Answered/Missed number, so the underlying cause stays visible per agent
  - FCR = no return contact from the same customer (Contact ID) within 7 days of a
    resolved chat's close time
"""
import datetime as dt

import numpy as np
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# Read-only -- this tool never writes back to the sheet.
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/drive.readonly',
]
GOOGLE_SHEETS_EPOCH = dt.date(1899, 12, 30)


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

# Gender inferred from first name, per Mahmoud -- used only for the 4PM-12AM WFH rule.
FEMALE = {
    'Nariman Shedid', 'Naira Emad', 'Waad Yassin', 'Heba Tarek', 'Mayar Khaled',
    'Nada Esaam', 'Duha Younis', 'Hagar Ahmed', 'Nada Sayed', 'Salma Adel',
    'Sondos Tarik', 'Basma Mostafa', 'Sara Hussien', 'Samaa Ahmed', 'Hagar Shaban',
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
    'Hagar Ahmed': ['hagar ahmed'],
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
    # Open question (flagged, not merged): Schedule's "Hagar Shaban" reads as a DIFFERENT
    # person from Agents-ID's "Hagar Ahmed" -- kept separate here on purpose.
    'Hagar Shaban': ['hagar shaban'],
}

# Schedule's "Employee name" (full/formal name) -> canonical Agents-ID name
SCHEDULE_NAME_MAP = {
    'Waad Alla El-dein Elsayed Mohamed': 'Waad Yassin',
    'Nariman Ezzat Amin Nagdy': 'Nariman Shedid',
    'Nayira Emad Hamdy Abu-Nar': 'Naira Emad',
    'Hagar Shaban': 'Hagar Shaban',
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

BREAK_STATES = {'Away - short break', 'Away - lunch break', 'Away - gomaa prayer'}
COACHING_STATE = 'Away - coaching'
TRAINING_STATE = 'Away - training'
TECH_STATE = 'Away - technical issue'
OFFLINE_STATE = 'Offline'
RECLASS_STATES = BREAK_STATES | {COACHING_STATE, TRAINING_STATE}

# Serviced and Successful merged into one "Serviced" bucket, per Mahmoud (Sep 2026) --
# the platform's own docs draw no real distinction Ops needs to track separately.
CALL_STATE_NORMALIZE = {'Successful': 'Serviced'}
CALL_STATES = ['Serviced', 'Dropped', 'No Answer', 'Abandoned', 'Blocked', 'Busy', 'Failed']
CALL_ANSWERED_STATES = {'Serviced'}

FCR_WINDOW_DAYS = 7


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
    if td is None or (isinstance(td, float) and pd.isna(td)):
        return ''
    total_sec = int(td.total_seconds())
    sign = '-' if total_sec < 0 else ''
    total_sec = abs(total_sec)
    h, rem = divmod(total_sec, 3600)
    m, s = divmod(rem, 60)
    return f"{sign}{h:02d}:{m:02d}:{s:02d}"


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
    agents['Team'] = agents.get('Team')
    if agents['Team'].isna().all():
        agents['Team'] = agents['Agent Name'].map(lambda n: TEAM_OVERRIDE.get(n, 'CS'))
    else:
        agents['Team'] = agents.apply(
            lambda r: r['Team'] if pd.notna(r['Team']) else TEAM_OVERRIDE.get(r['Agent Name'], 'CS'), axis=1)
    agents['Agent ID'] = agents['Agent ID'].astype(int)
    id_to_name = dict(zip(agents['Agent ID'], agents['Agent Name']))
    roster = agents['Agent Name'].tolist()

    # add off-roster-but-scheduled people (e.g. "Hagar Shaban") so Schedule/Calls/
    # Activity data for them isn't silently dropped just because they have no Agent ID yet
    sched_names = set(SCHEDULE_NAME_MAP.values())
    off_roster = sorted(sched_names - set(roster))
    full_roster = roster + off_roster

    calls = calls.copy()
    calls['Agent_n'] = calls['Agent'].map(norm)
    calls['Created'] = pd.to_datetime(calls['Created'], errors='coerce')
    calls['State'] = calls['State'].astype(str).str.strip().map(lambda s: CALL_STATE_NORMALIZE.get(s, s))

    chats = chats.copy()
    chats['DateTime Conversation Started'] = pd.to_datetime(chats['DateTime Conversation Started'], errors='coerce')
    chats['DateTime Conversation Resolved'] = pd.to_datetime(chats['DateTime Conversation Resolved'], errors='coerce')
    # Opportunistic -- these columns aren't in the minimal Chats shape this tool was
    # first built against, but the raw Carecomm export can carry them. Parsed with the
    # same to_timedelta() used for Calls (tolerant of "H:M:S" text AND the Sheets
    # duration-serial float) whenever the column exists, so richer per-agent chat KPIs
    # (see compute_chats) switch on automatically the moment the live Chats tab has
    # them -- no crash, no extra KPI, if it doesn't.
    for col in ('First Response Time', 'Resolution Time'):
        if col in chats.columns:
            chats[col + ' (td)'] = chats[col].map(to_timedelta)

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
    else:
        schedule['is_wfh_flag'] = False

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
        'full_roster': full_roster, 'id_to_name': id_to_name,
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

    calls = _worksheet_to_df(sh.worksheet('Calls'))
    if 'Created' in calls.columns:
        calls['Created'] = calls['Created'].map(_serial_to_ts)

    chats = _worksheet_to_df(sh.worksheet('Chats'))
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
        return day_dt, day_dt + pd.Timedelta(hours=9)
    return None, None


def compute_adherence(data, start, end):
    schedule = data['schedule']
    activity = data['activity']
    roster = data['full_roster']

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
        is_female = agent in FEMALE

        for _, srow in a_sched.iterrows():
            day_dt = srow['Date']
            shift_label = srow['Shift']
            is_wfh = bool(srow['is_wfh_flag'])

            if shift_label in UNRESOLVED_LABELS:
                unclassified_days.append({'Agent': agent, 'Date': day_dt, 'Shift': shift_label})
                continue
            if shift_label in NON_WORKING or shift_label not in SHIFT_PLANNED_MIN:
                daily_rows.append({
                    'Agent': agent, 'Date': day_dt, 'Shift': shift_label,
                    'Working Day': False, 'Planned Minutes': 0, 'Actual Minutes': 0,
                    'Late Minutes': 0, 'Early Logout Minutes': 0, 'Break Minutes': 0,
                    'Data Status': 'Complete',
                })
                continue

            win_start, win_end = shift_window(shift_label, day_dt)
            planned = SHIFT_PLANNED_MIN[shift_label]
            nine_hour = False
            if shift_label in EIGHT_HOUR_TYPES and (is_wfh or (is_female and shift_label == '4 PM - 12 AM')):
                nine_hour = True
                win_end = win_end + pd.Timedelta(hours=1)
                planned = 480

            data_status = 'Complete' if win_end <= cutoff else 'Incomplete*'
            clipped = clip(intervals, win_start, min(win_end + pd.Timedelta(hours=12), win_end + pd.Timedelta(hours=1)))
            # reclassify short Offline gaps sandwiched between identical break/coaching/training states
            for i in range(1, len(clipped) - 1):
                if clipped[i][2] == OFFLINE_STATE:
                    prev_s, next_s = clipped[i - 1][2], clipped[i + 1][2]
                    if prev_s == next_s and prev_s in RECLASS_STATES:
                        clipped[i][2] = prev_s

            def mins(states):
                return sum((e - s).total_seconds() / 60 for s, e, st in clipped if st in states)

            online_min = mins({'Available'})
            busy_min = mins({'Busy'})
            break_min = mins(BREAK_STATES)
            coaching_min = mins({COACHING_STATE})
            training_min = mins({TRAINING_STATE})
            tech_min = mins({TECH_STATE})
            actual = online_min + busy_min + coaching_min + training_min + tech_min

            active = [iv for iv in clipped if iv[2] != OFFLINE_STATE]
            login = active[0][0] if active else None
            logout = active[-1][1] if active else None
            late_min = max(0, (login - win_start).total_seconds() / 60) if login else planned
            early_min = max(0, (win_end - logout).total_seconds() / 60) if logout else planned

            daily_rows.append({
                'Agent': agent, 'Date': day_dt, 'Shift': shift_label,
                'Working Day': True, 'Nine Hour': nine_hour, 'Is WFH': is_wfh,
                'Planned Minutes': planned, 'Actual Minutes': round(actual, 1),
                'Online Minutes': round(online_min, 1), 'Busy Minutes': round(busy_min, 1),
                'Break Minutes': round(break_min, 1), 'Coaching Minutes': round(coaching_min, 1),
                'Training Minutes': round(training_min, 1), 'Technical Minutes': round(tech_min, 1),
                'Late Minutes': round(late_min, 1), 'Early Logout Minutes': round(early_min, 1),
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
        online = complete.get('Online Minutes', pd.Series(dtype=float)).sum()
        busy = complete.get('Busy Minutes', pd.Series(dtype=float)).sum()
        reachable = online + busy
        occupancy = (busy / reachable * 100) if reachable else np.nan
        shrink_min = (complete.get('Break Minutes', pd.Series(dtype=float)).sum()
                      + complete.get('Coaching Minutes', pd.Series(dtype=float)).sum()
                      + complete.get('Training Minutes', pd.Series(dtype=float)).sum()
                      + complete.get('Technical Minutes', pd.Series(dtype=float)).sum())
        planned_complete = complete['Planned Minutes'].sum()
        shrinkage = (shrink_min / planned_complete * 100) if planned_complete else np.nan
        adherence_pct = min(100.0, actual / planned * 100) if planned else np.nan
        rows.append({
            'Agent': agent, 'Team': TEAM_OVERRIDE.get(agent, 'CS'),
            'Working Days': int(len(working)), 'Days Off': int((~a['Working Day']).sum()),
            'Planned Minutes': round(planned, 0), 'Actual Minutes': round(actual, 0),
            'Adherence %': round(adherence_pct, 1) if pd.notna(adherence_pct) else None,
            'Late Minutes': round(complete['Late Minutes'].sum(), 0),
            'Early Logout Minutes': round(complete['Early Logout Minutes'].sum(), 0),
            'Occupancy %': round(occupancy, 1) if pd.notna(occupancy) else None,
            'Shrinkage %': round(shrinkage, 1) if pd.notna(shrinkage) else None,
            'Incomplete Days': incomplete_n,
        })
    adherence_cols = ['Agent', 'Team', 'Working Days', 'Days Off', 'Planned Minutes', 'Actual Minutes',
                       'Adherence %', 'Late Minutes', 'Early Logout Minutes', 'Occupancy %',
                       'Shrinkage %', 'Incomplete Days']
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
            row['Avg First Response Time'] = fmt_td(a['First Response Time (td)'].mean())
        if has_restime:
            row['Avg Resolution Time'] = fmt_td(closed['Resolution Time (td)'].mean()) if not closed.empty else ''
        rows.append(row)
    chats_cols = ['Agent', 'Team', 'Assigned', 'Closed', 'Unique Contacts', 'Avg per day', 'FCR %']
    if has_frt:
        chats_cols.append('Avg First Response Time')
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
def compute_calls(data, start, end, denom_days):
    calls = data['calls']
    roster = data['full_roster']
    win = calls[_in_range(calls['Created'], start, end)].copy()
    win['handle_td'] = win['Handling Duration'].map(to_timedelta)
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
        days = denom_days.get(agent)
        row = {
            'Agent': agent, 'Team': TEAM_OVERRIDE.get(agent, 'CS'),
            'Total Calls': total, 'Answered %': round(answered / total * 100, 1) if total else None,
            'Avg per day': round(total / days, 2) if days else None,
            'Avg Handling Time': fmt_td(avg_handle),
        }
        row.update(state_counts)
        rows.append(row)
    calls_cols = ['Agent', 'Team', 'Total Calls', 'Answered %', 'Avg per day', 'Avg Handling Time'] + CALL_STATES
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

    calls_totals = {
        'state_totals': state_totals_all, 'total_calls': total_calls_all,
        'answered_rate': answered_rate_all, 'unattributed_calls': unattributed,
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

    o['agents_in_scope'] = int(len(set(chats_df['Agent']) | set(calls_df['Agent']) | set(adherence_df['Agent'] if not adherence_df.empty else [])))
    if not adherence_df.empty:
        o['avg_adherence'] = round(adherence_df['Adherence %'].dropna().mean(), 1) if adherence_df['Adherence %'].notna().any() else None
        o['avg_occupancy'] = round(adherence_df['Occupancy %'].dropna().mean(), 1) if adherence_df['Occupancy %'].notna().any() else None
        o['avg_shrinkage'] = round(adherence_df['Shrinkage %'].dropna().mean(), 1) if adherence_df['Shrinkage %'].notna().any() else None
    else:
        o['avg_adherence'] = o['avg_occupancy'] = o['avg_shrinkage'] = None
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
    overall = compute_overall(chats_df, calls_df, calls_totals, adherence_df)
    return {
        'data': data, 'overall': overall, 'chats': chats_df, 'calls': calls_df,
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

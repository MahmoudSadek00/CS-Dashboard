import datetime as dt
import json

import pandas as pd
import streamlit as st
import plotly.graph_objects as go
import plotly.io as pio

import logic

st.set_page_config(page_title="CS Pulse", layout="wide")

# Same dark, bordered-tile look as Ops Pulse (see .streamlit/config.toml for the base
# theme colors) -- kept identical on purpose so this reads as the same product family.
st.markdown("""
<style>
    div[data-testid="stMetric"] {
        background: linear-gradient(180deg, #171A21 0%, #14161B 100%);
        border: 1px solid #262B35;
        border-radius: 12px;
        padding: 14px 18px 10px 18px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.35);
    }
    div[data-testid="stMetricLabel"] p {
        color: #9AA4B2 !important;
        font-size: 0.82rem !important;
        font-weight: 500 !important;
        letter-spacing: .02em;
        text-transform: uppercase;
    }
    div[data-testid="stMetric"] { min-width: 0; }
    div[data-testid="stMetricValue"] {
        color: #F4F6F8;
        font-weight: 700;
        font-size: 1.55rem;
        line-height: 1.25;
        white-space: normal !important;
        overflow: visible !important;
        text-overflow: clip !important;
        word-break: break-word;
    }
    h1, h2, h3 { letter-spacing: -0.01em; }
    h2, h3 { border-bottom: 1px solid #262B35; padding-bottom: 6px; margin-top: 1.6rem; }
    div[data-testid="stDataFrame"] { border: 1px solid #262B35; border-radius: 10px; overflow: hidden; }
    section[data-testid="stSidebar"] { border-right: 1px solid #262B35; }
</style>
""", unsafe_allow_html=True)

pio.templates.default = "plotly_dark"

st.title("CS Pulse")
st.caption("Customer Support performance -- Chats, Calls and Adherence, company-wide and per agent.")

# Serviced and Successful are merged into one "Serviced" bucket in logic.py.
STATE_COLORS = {
    'Serviced': '#2E7D32', 'Dropped': '#C62828',
    'No Answer': '#F9A825', 'Abandoned': '#EF6C00', 'Blocked': '#8D6E63',
    'Busy': '#5C6BC0', 'Failed': '#B71C1C',
}


def _pct(v):
    return f"{v:.1f}%" if v is not None else "—"


# ---------------------------------------------------------------------------
# Data source: live Google Sheet only (needs the service account secret
# configured on this deployment, see README). logic.py still exposes
# build_report() for a raw xlsx workbook -- kept for local debugging/tests --
# but the UI no longer offers it as a source now that the live sheet is trusted.
# ---------------------------------------------------------------------------
DEFAULT_SPREADSHEET_ID = '1Lz9OaWLpEM-m9w-5bxITTPKs9e00ZfuGtCl3m1Iicpw'

# The Orders "Clean" sheet AOV is pulled from -- a DIFFERENT spreadsheet from the CS
# sheet above, per Mahmoud (Sep 2026). Needs its own Viewer share for the same service
# account -- see the README. Tab is "Orders", last column "Salesman".
CLEAN_SHEET_ID = '1dZMqtqvnxe6GspH0C10AvXECB74NP-ZjDG_BihMOkmg'
CLEAN_SHEET_ORDERS_TAB = 'Orders'


def _load_creds_info():
    try:
        if 'gcp_service_account' in st.secrets:
            return dict(st.secrets['gcp_service_account'])
        if 'gcp_service_account_json' in st.secrets:
            raw = st.secrets['gcp_service_account_json']
            return json.loads(raw) if isinstance(raw, str) else dict(raw)
    except Exception:
        pass
    return None


creds_info = _load_creds_info()

with st.sidebar:
    st.header("Data")
    if not creds_info:
        st.error("No Google credential configured on this deployment -- see the README.")
        st.stop()
    with st.expander("Advanced (data source)"):
        spreadsheet_id = st.text_input("Spreadsheet ID", value=DEFAULT_SPREADSHEET_ID)
    if 'cache_bump' not in st.session_state:
        st.session_state['cache_bump'] = 0
    if st.button("🔄 Refresh from Google Sheets"):
        st.session_state['cache_bump'] += 1
        st.cache_data.clear()

    st.divider()
    st.header("Period")
    default_start = dt.date(2026, 7, 1)
    default_end = dt.date(2026, 8, 31)

    period_mode = st.radio(
        "Mode", ["Single period", "Compare two periods"], index=0,
        help="\"Compare two periods\" is the same concept as the Ops Pulse comparison report -- "
             "pick Period A and Period B, see every metric side by side with a delta, company-wide "
             "and per agent. The dashboard below always reflects Period B (or the single period).",
    )
    enable_comparison = period_mode == "Compare two periods"
    range_a = range_b = None

    if not enable_comparison:
        start = st.date_input("Start date", value=default_start, key="single_start")
        end = st.date_input("End date", value=default_end, key="single_end")
    else:
        period_len = (default_end - default_start).days + 1
        default_a_end = default_start - dt.timedelta(days=1)
        default_a_start = default_a_end - dt.timedelta(days=period_len - 1)
        st.caption("Period B is the main period this dashboard shows below; Period A defaults to "
                   "the same-length period right before it -- all four dates are editable.")
        a_start = st.date_input("Period A start", value=default_a_start, key="a_start")
        a_end = st.date_input("Period A end", value=default_a_end, key="a_end")
        b_start = st.date_input("Period B start", value=default_start, key="b_start")
        b_end = st.date_input("Period B end", value=default_end, key="b_end")
        range_a = (a_start, a_end)
        range_b = (b_start, b_end)
        start, end = b_start, b_end

    st.divider()
    st.caption("Bahrain team is always excluded from this report.")

if not (isinstance(start, dt.date) and isinstance(end, dt.date)) or start > end:
    st.warning("Pick a valid date range (start on or before end).")
    st.stop()


@st.cache_resource(show_spinner=False)
def _client(_creds_info):
    return logic.get_client(_creds_info)


@st.cache_data(ttl=600, show_spinner=False)
def _cached_sheet_report(_gc, spreadsheet_id, start, end, cache_bump):
    # _gc / leading-underscore args aren't hashed by Streamlit's cache; cache_bump
    # (no underscore) IS part of the cache key, so the sidebar Refresh button forces
    # a fresh read even before the 600s TTL expires.
    return logic.build_report_from_sheet(_gc, spreadsheet_id, start, end)


@st.cache_data(ttl=600, show_spinner=False)
def _cached_orders_df(_gc, spreadsheet_id, tab_name, cache_bump):
    return logic.load_orders_clean(_gc, spreadsheet_id, tab_name)


if not spreadsheet_id:
    st.stop()
try:
    gc = _client(creds_info)
except Exception as e:
    st.error(f"Couldn't connect to Google Sheets with the configured credential: {e}")
    st.stop()
with st.spinner("Reading the live sheet and crunching the numbers..."):
    try:
        result = _cached_sheet_report(gc, spreadsheet_id, pd.Timestamp(start), pd.Timestamp(end), st.session_state['cache_bump'])
    except Exception as e:
        st.error(
            f"Couldn't read the spreadsheet: {e}\n\nMost likely cause: the service "
            "account isn't shared as a Viewer on this specific sheet yet -- see the README."
        )
        st.stop()

comparison = None
if enable_comparison:
    start_a, end_a = range_a
    if start_a > end_a:
        st.sidebar.warning("Pick a valid Period A (start on or before end) to see the comparison.")
    else:
        # Period B is exactly the main `start`/`end` above, so `result` (already fetched)
        # IS report_b -- no need to fetch it a second time, just fetch Period A.
        with st.spinner("Building the comparison (reading Period A)..."):
            try:
                report_a = _cached_sheet_report(gc, spreadsheet_id, pd.Timestamp(start_a), pd.Timestamp(end_a), st.session_state['cache_bump'])
                comparison = logic.build_comparison(report_a, result)
            except Exception as e:
                st.warning(f"Couldn't build the comparison: {e}")

overall = result['overall']
chats_df = result['chats']
calls_df = result['calls']
adherence_df = result['adherence']
unclassified = result['unclassified_shifts']
daily_audit = result['daily_audit']
unmatched_ids = result['unmatched_chat_ids']
unattributed_calls = result['unattributed_calls']
chat_category_totals = result['chat_category_totals']

# ---------------------------------------------------------------------------
# Overall cards
# ---------------------------------------------------------------------------
st.header("Overall")

def _badge_delta(value, key):
    # st.metric's delta only accepts a short string, and we don't want the usual
    # green-up/red-down arrow semantics (this isn't a "change" value) -- delta_color="off"
    # renders it as plain grey text next to the number instead.
    return logic.target_badge(value, key)


c1, c2, c3, c4 = st.columns(4)
c1.metric("Chats Closed", f"{overall['total_chats_closed']:,}")
c2.metric(
    "Chats FCR Rate", _pct(overall['fcr_rate']),
    delta=_badge_delta(overall['fcr_rate'], 'fcr_rate'), delta_color="off",
    help="No return contact from the same customer within 7 days of the chat closing. "
         "Badge shows this against the CEO Q3 2026 scorecard target (90%).",
)
c3.metric("Total Calls", f"{overall['total_calls']:,}")
c4.metric(
    "Calls Answered Rate", _pct(overall['answered_rate']),
    delta=_badge_delta(overall['answered_rate'], 'answered_rate'), delta_color="off",
    help="Serviced, as a share of ALL calls in the period -- including calls that never reached any agent "
         "(e.g. Dropped), which is why this is computed against the full Calls tab rather than the per-agent table. "
         "Badge shows this against the CEO Q3 2026 scorecard target (95%, stretch 98%).",
)

c5, c6, c7, c8 = st.columns(4)
c5.metric("Dropped Rate", _pct(overall['dropped_rate']), help="Dropped calls as a share of ALL calls in the period.")
c6.metric("Abandoned Rate", _pct(overall['abandoned_rate']), help="Abandoned calls as a share of ALL calls in the period.")
c7.metric("Avg. Adherence", _pct(overall['avg_adherence']))
with c8:
    st.metric(
        "Avg. Occupancy", _pct(overall['avg_occupancy']),
        help="(Busy call time + chat-handling time) / (Available + Busy) time, averaged "
             "across agents. Click below for the calls-vs-chats split.",
    )
    with st.popover("Calls vs. chats split"):
        if overall['occupancy_calls_pct'] is None:
            st.caption("No occupied time (calls or chats) in this period.")
        else:
            st.write(f"📞 Calls: **{overall['occupancy_calls_pct']:.1f}%** "
                     f"({overall['occupancy_calls_min']:,.0f} min)")
            st.write(f"💬 Chats: **{overall['occupancy_chat_pct']:.1f}%** "
                     f"({overall['occupancy_chat_min']:,.0f} min)")
            st.caption("Share of the combined Busy-calls + chat-handling minutes behind "
                       "the Occupancy % above (company-wide, this period).")

c9, c10, _c11, _c12 = st.columns(4)
c9.metric("Avg. Shrinkage", _pct(overall['avg_shrinkage']))
c10.metric("Agents in Scope", f"{overall['agents_in_scope']}")

st.subheader("Calls by state")
state_totals = overall['call_state_totals']
state_total_n = sum(state_totals.values())
state_labels = [
    f"{v:,} ({v / state_total_n * 100:.1f}%)" if state_total_n else f"{v:,}"
    for v in state_totals.values()
]
fig = go.Figure(go.Bar(
    x=list(state_totals.values()), y=list(state_totals.keys()), orientation='h',
    marker_color=[STATE_COLORS[s] for s in state_totals], text=state_labels,
    textposition='outside', hovertemplate='%{y}: %{x:,}<extra></extra>',
))
fig.update_layout(height=300, margin=dict(l=10, r=10, t=10, b=10), xaxis_title="Calls", yaxis_title=None)
st.plotly_chart(fig, use_container_width=True)

# Opportunistic -- only appears when the live Chats tab has a "Conversation Category"
# column (see compute_chats in logic.py). Silently absent otherwise, no error shown.
if chat_category_totals:
    st.subheader("Chats by category")
    cats = dict(sorted(chat_category_totals.items(), key=lambda kv: kv[1]))
    cat_total_n = sum(cats.values())
    cat_labels = [
        f"{v:,} ({v / cat_total_n * 100:.1f}%)" if cat_total_n else f"{v:,}"
        for v in cats.values()
    ]
    fig_cat = go.Figure(go.Bar(
        x=list(cats.values()), y=list(cats.keys()), orientation='h',
        marker_color='#5B8DEF', text=cat_labels,
        textposition='outside', hovertemplate='%{y}: %{x:,}<extra></extra>',
    ))
    fig_cat.update_layout(height=max(220, 28 * len(cats)), margin=dict(l=10, r=10, t=10, b=10),
                           xaxis_title="Chats", yaxis_title=None)
    st.plotly_chart(fig_cat, use_container_width=True)

# ---------------------------------------------------------------------------
# Agent filter -- applies to the three per-agent tables below only. The Overall
# cards above stay company-wide on purpose (so they always read as "everyone"),
# but everything from here down can be narrowed to a hand-picked set of agents.
# ---------------------------------------------------------------------------
all_agents = sorted(set(chats_df['Agent']) | set(calls_df['Agent']) | set(adherence_df['Agent']))
with st.sidebar:
    st.header("Agents")
    selected_agents = st.multiselect(
        "Show in the tables below", options=all_agents, default=all_agents,
        help="Filters Chats / Calls / Adherence below. Leave everyone selected (default) to see the full team; pick a few to compare specific agents.",
    )
    sel_all = st.button("Select all", use_container_width=True)
if sel_all:
    selected_agents = all_agents

if not selected_agents:
    st.warning("No agents selected in the sidebar -- pick at least one to see the tables below.")
    selected_agents = all_agents

# ---------------------------------------------------------------------------
# Currency -- for the AOV per Agent section below. Per Mahmoud, GC's markets
# settle in one of three currencies, not one-per-country: UAE and Oman orders
# are both in AED; Saudi, Kuwait and Qatar orders are all in BHD; Iraq is its
# own IQD. So this is three rate fields, not seven -- entered by hand (Ops
# Pulse itself never converts currency anywhere in its own exports, so there's
# no existing method to inherit, and rates move day to day so this app never
# fetches or guesses one on its own). AED and BHD are both long-standing hard
# pegs to the dollar (unchanged for decades), so their defaults below are safe
# to leave as-is; IQD drifts, so that one is worth checking periodically --
# these are starting defaults only, not live rates, and always editable.
# ---------------------------------------------------------------------------
MARKET_CURRENCY = {
    'UAE': 'AED', 'OM': 'AED',
    'SA': 'BHD', 'KW': 'BHD', 'QA': 'BHD', 'BH': 'BHD',
    'IQ': 'IQD',
}
CURRENCY_DEFAULTS = {'AED': 3.6725, 'BHD': 0.3760, 'IQD': 1310.0}

with st.sidebar:
    st.divider()
    st.header("Currency")
    currency_mode = st.radio(
        "Currency", ["Original (as recorded)", "USD"], index=1, label_visibility="collapsed",
    )
    fx_rates = {}
    if currency_mode == "USD":
        st.caption(
            "1 USD = how many of each currency -- NOT live rates, starting defaults only, edit "
            "freely. AED and BHD are both long-standing hard pegs to the Dollar; IQD drifts more, "
            "so it's worth checking that one periodically."
        )
        c_aed, c_bhd, c_iqd = st.columns(3)
        aed_rate = c_aed.number_input("1 USD = _\nAED", min_value=0.0, value=CURRENCY_DEFAULTS['AED'], step=0.0001, format="%.4f")
        bhd_rate = c_bhd.number_input("1 USD = _\nBHD", min_value=0.0, value=CURRENCY_DEFAULTS['BHD'], step=0.0001, format="%.4f")
        iqd_rate = c_iqd.number_input("1 USD = _\nIQD", min_value=0.0, value=CURRENCY_DEFAULTS['IQD'], step=1.0, format="%.1f")
        currency_rates = {'AED': aed_rate, 'BHD': bhd_rate, 'IQD': iqd_rate}
        for market, currency in MARKET_CURRENCY.items():
            rate = currency_rates.get(currency)
            if rate:
                fx_rates[market] = rate

# ---------------------------------------------------------------------------
# AOV per agent data -- fetched once here (before Export, which needs it) and
# rendered further down the page. From the SEPARATE "Orders Clean" spreadsheet's
# Orders tab, Salesman column (last column: either "Created by customer" or an
# agent's name -- only the agent-name rows count). Own try/except since this is
# a different sheet that may not be shared with the service account yet.
# ---------------------------------------------------------------------------
aov_df = aov_diag = aov_error = None
try:
    orders_df = _cached_orders_df(gc, CLEAN_SHEET_ID, CLEAN_SHEET_ORDERS_TAB, st.session_state['cache_bump'])
    aov_df, aov_diag = logic.compute_aov_by_agent(orders_df, pd.Timestamp(start), pd.Timestamp(end), fx_rates=fx_rates)
except Exception as e:
    aov_error = (
        f"Couldn't read the Orders (Clean) sheet: {e}\n\nMost likely cause: the service "
        "account isn't shared as a Viewer on this specific sheet yet -- see the README."
    )

# ---------------------------------------------------------------------------
# Export -- whole report or a hand-picked set of sections, as one .xlsx with a
# sheet per section. Always exports the FULL agent roster (not narrowed by the
# Agents filter above) so the downloaded file reads as a complete report.
# ---------------------------------------------------------------------------
EXPORT_SECTIONS = (
    ['Overall', 'Chats', 'Calls', 'Adherence']
    + (['AOV'] if aov_df is not None and not aov_df.empty else [])
    + (['Comparison'] if comparison is not None else [])
)
with st.sidebar:
    st.divider()
    st.header("Export")
    export_sections = st.multiselect(
        "Sections to include", options=EXPORT_SECTIONS, default=EXPORT_SECTIONS,
        help="Leave everything selected for the whole report, or pick just the section(s) you need.",
    )
    period_label = f"{start:%Y-%m-%d} → {end:%Y-%m-%d}"
    currency_note = (
        "AOV converted to USD using: " + ", ".join(f"{k}={v:g}" for k, v in currency_rates.items())
        if currency_mode == "USD" else "AOV shown in original currency (not converted)"
    )
    export_bytes = logic.export_excel(
        result, comparison=comparison, sections=export_sections,
        aov_df=aov_df if 'AOV' in export_sections else None,
        period_label=period_label, currency_note=currency_note,
    ) if export_sections else None
    st.download_button(
        "⬇️ Download report (.xlsx)", data=export_bytes or b"",
        file_name=f"cs_pulse_{start:%Y%m%d}_to_{end:%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        disabled=not export_sections, use_container_width=True,
    )


def _filter_agents(df):
    return df[df['Agent'].isin(selected_agents)].reset_index(drop=True) if not df.empty else df


chats_df_view = _filter_agents(chats_df)
calls_df_view = _filter_agents(calls_df)
adherence_df_view = _filter_agents(adherence_df)

# ---------------------------------------------------------------------------
# Chats -- per agent
# ---------------------------------------------------------------------------
st.header("Chats -- per agent")
if chats_df_view.empty:
    st.warning("No chat data matched the roster (or the agent filter) in this period.")
else:
    st.dataframe(
        chats_df_view, use_container_width=True, hide_index=True,
        column_config={
            'FCR %': st.column_config.NumberColumn(format="%.1f%%"),
            'Avg per day': st.column_config.NumberColumn(format="%.2f"),
        },
    )

# ---------------------------------------------------------------------------
# Calls -- per agent
# ---------------------------------------------------------------------------
st.header("Calls -- per agent")
if calls_df_view.empty:
    st.warning("No call data matched the roster (or the agent filter) in this period.")
else:
    st.dataframe(
        calls_df_view, use_container_width=True, hide_index=True,
        column_config={
            'Answered %': st.column_config.NumberColumn(format="%.1f%%"),
            'Avg per day': st.column_config.NumberColumn(format="%.2f"),
        },
    )

# ---------------------------------------------------------------------------
# Adherence -- per agent
# ---------------------------------------------------------------------------
st.header("Adherence -- per agent")
st.caption("Only agents with rows in the Schedule tab for this period can have Adherence computed.")
if adherence_df_view.empty:
    st.warning("No agents in the Schedule tab fall inside this period (or match the agent filter).")
else:
    st.dataframe(
        adherence_df_view, use_container_width=True, hide_index=True,
        column_config={
            'Adherence %': st.column_config.NumberColumn(format="%.1f%%"),
            'Occupancy %': st.column_config.NumberColumn(format="%.1f%%"),
            'Shrinkage %': st.column_config.NumberColumn(format="%.1f%%"),
        },
    )

# ---------------------------------------------------------------------------
# AOV per agent -- rendering only; aov_df/aov_diag/aov_error were fetched
# earlier (before Export, which also needs them) using the Currency rates set
# in the sidebar above.
# ---------------------------------------------------------------------------
st.header("AOV per Agent")
st.caption(
    "From the Orders (Clean) sheet's \"Salesman\" column -- orders created by the customer "
    "themselves are excluded, only orders attributed to an agent's own sales count."
)
if aov_error:
    st.info(aov_error)
elif aov_df is None:
    st.caption("Couldn't compute AOV for this period.")
elif aov_df.empty:
    st.caption("No agent-attributed orders matched the roster in this period.")
else:
    has_usd = 'AOV (USD)' in aov_df.columns
    if not fx_rates:
        st.warning(
            "⚠️ Currency is set to \"Original (as recorded)\" -- these AOV figures are shown "
            "in the Orders sheet's original currency, as-is, not converted to USD. Switch the "
            "**Currency** toggle in the sidebar to \"USD\" to compare against the CEO "
            "scorecard's AOV target ($90-130, market-dependent).",
            icon="⚠️",
        )
    else:
        st.caption(
            "AOV (USD) columns use the rate(s) set in the sidebar's **Currency** section "
            "(UAE + Oman via AED, Saudi + Kuwait + Qatar + Bahrain via BHD, Iraq via IQD). "
            "Orders in any other market stay out of those columns (shown in the "
            "Orders/AOV/Total Value columns in local currency only)."
        )
    if aov_diag:
        st.info(aov_diag)
    aov_df_view = aov_df[aov_df['Agent'].isin(selected_agents)].reset_index(drop=True)
    col_config = {'AOV': st.column_config.NumberColumn(format="%.2f")}
    if has_usd:
        col_config['AOV (USD)'] = st.column_config.NumberColumn(format="$%.2f")
        col_config['Total Value (USD)'] = st.column_config.NumberColumn(format="$%.2f")
    st.dataframe(aov_df_view, use_container_width=True, hide_index=True, column_config=col_config)

# ---------------------------------------------------------------------------
# Comparison -- Period A vs Period B, same concept as the Ops Pulse comparison
# report. Only rendered when enabled + successfully built in the sidebar above.
# ---------------------------------------------------------------------------
if comparison is not None:
    st.header("Comparison -- Period A vs Period B")
    st.caption(
        f"Period A: {range_a[0]:%Y-%m-%d} → {range_a[1]:%Y-%m-%d}    |    "
        f"Period B: {range_b[0]:%Y-%m-%d} → {range_b[1]:%Y-%m-%d}"
    )

    hl = comparison['highlights']
    h1, h2 = st.columns(2)
    with h1:
        st.subheader("✅ What's working")
        if hl['improved']:
            for line in hl['improved']:
                st.markdown(f"- {line}")
        else:
            st.caption("No notable improvements found.")
    with h2:
        st.subheader("⚠️ What's not")
        if hl['declined']:
            for line in hl['declined']:
                st.markdown(f"- {line}")
        else:
            st.caption("No notable declines found.")

    st.subheader("Overall")
    overall_cmp = comparison['overall'].drop(columns=['Higher Is Better'])
    st.dataframe(
        overall_cmp, use_container_width=True, hide_index=True,
        column_config={
            'Period A': st.column_config.NumberColumn(format="%.1f"),
            'Period B': st.column_config.NumberColumn(format="%.1f"),
            'Delta': st.column_config.NumberColumn(format="%+.1f"),
        },
    )

    with st.expander("Chats -- per agent, Period A vs B"):
        st.dataframe(comparison['chats'], use_container_width=True, hide_index=True)
    with st.expander("Calls -- per agent, Period A vs B"):
        st.dataframe(comparison['calls'], use_container_width=True, hide_index=True)
    with st.expander("Adherence -- per agent, Period A vs B"):
        st.dataframe(comparison['adherence'], use_container_width=True, hide_index=True)

# ---------------------------------------------------------------------------
# Diagnostics -- data-quality notes, not hidden away
# ---------------------------------------------------------------------------
st.header("Diagnostics")
d1, d2, d3, d4 = st.columns(4)
with d1:
    st.subheader("Unclassified shift days")
    st.caption('"Task" / "Task - Normal Shift" -- meaning still unconfirmed, excluded from Adherence.')
    if unclassified.empty:
        st.caption("None in this period.")
    else:
        st.dataframe(unclassified, use_container_width=True, hide_index=True)
with d2:
    st.subheader("Unmatched Chat Assignee IDs")
    st.caption("Chat volume under an Agent ID not found on the (Bahrain-excluded) roster.")
    if unmatched_ids.empty:
        st.caption("None in this period.")
    else:
        st.dataframe(unmatched_ids.rename("Chats").to_frame(), use_container_width=True)
with d3:
    st.subheader("Calls with no agent attributed")
    st.caption("Included in the Overall cards/chart above, but can't appear in the per-agent Calls table -- "
               "there's no Agent value to attribute them to (almost all Dropped calls are like this by nature).")
    st.metric("Unattributed calls", f"{unattributed_calls:,}")
with d4:
    st.subheader("Fri/Sat/Sun WFH override")
    st.caption('Any working shift on Friday/Saturday/Sunday is treated as WFH regardless of the '
               'sheet\'s "Is WFH" value -- rows below are the ones where the sheet actually said '
               'something else (blank or "No") and got overridden.')
    wfh_overrides = daily_audit[daily_audit.get('WFH Overridden', False) == True] if not daily_audit.empty else daily_audit
    if wfh_overrides.empty:
        st.caption("None in this period.")
    else:
        st.dataframe(wfh_overrides[['Agent', 'Date', 'Shift']], use_container_width=True, hide_index=True)

agents_missing_id = result['data'].get('agents_missing_id', [])
if agents_missing_id:
    st.warning(
        f"⚠️ {len(agents_missing_id)} name(s) on the **Agents ID** tab have no Agent ID filled in "
        f"yet: {', '.join(agents_missing_id)}. They're still included via name-matching (Calls/Activity/"
        "Adherence), but can't receive Chats -- the Chats tab's Assignee column is ID-based, not name-"
        "based. Fill in their Agent ID on the sheet to fix this. (This used to crash the whole app on "
        "load instead -- a blank Agent ID cell is now skipped rather than fatal.)"
    )

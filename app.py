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
    date_range = st.date_input("Date range", value=(default_start, default_end))
    st.divider()
    st.caption("Bahrain team is always excluded from this report.")

if isinstance(date_range, tuple) and len(date_range) == 2:
    start, end = date_range
else:
    st.warning("Pick a full date range (start and end).")
    st.stop()

with st.sidebar:
    st.divider()
    st.header("Comparison")
    enable_comparison = st.checkbox(
        "Compare two periods", value=False,
        help="Same concept as the Ops Pulse comparison report -- pick Period A and Period B, "
             "see every metric side by side with a delta, company-wide and per agent.",
    )
    range_a = range_b = None
    if enable_comparison:
        period_len = (end - start).days + 1
        default_b_start, default_b_end = start, end
        default_a_end = start - dt.timedelta(days=1)
        default_a_start = default_a_end - dt.timedelta(days=period_len - 1)
        st.caption("Period B defaults to the main date range above; Period A defaults to the "
                   "same-length period right before it -- both are editable.")
        range_a = st.date_input("Period A", value=(default_a_start, default_a_end), key="period_a")
        range_b = st.date_input("Period B", value=(default_b_start, default_b_end), key="period_b")


@st.cache_resource(show_spinner=False)
def _client(_creds_info):
    return logic.get_client(_creds_info)


@st.cache_data(ttl=600, show_spinner=False)
def _cached_sheet_report(_gc, spreadsheet_id, start, end, cache_bump):
    # _gc / leading-underscore args aren't hashed by Streamlit's cache; cache_bump
    # (no underscore) IS part of the cache key, so the sidebar Refresh button forces
    # a fresh read even before the 600s TTL expires.
    return logic.build_report_from_sheet(_gc, spreadsheet_id, start, end)


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
    valid_a = isinstance(range_a, tuple) and len(range_a) == 2
    valid_b = isinstance(range_b, tuple) and len(range_b) == 2
    if not (valid_a and valid_b):
        st.sidebar.warning("Pick a full date range for both Period A and Period B to see the comparison.")
    else:
        start_a, end_a = range_a
        start_b, end_b = range_b
        with st.spinner("Building the comparison (reading both periods)..."):
            try:
                report_a = _cached_sheet_report(gc, spreadsheet_id, pd.Timestamp(start_a), pd.Timestamp(end_a), st.session_state['cache_bump'])
                report_b = _cached_sheet_report(gc, spreadsheet_id, pd.Timestamp(start_b), pd.Timestamp(end_b), st.session_state['cache_bump'])
                comparison = logic.build_comparison(report_a, report_b)
            except Exception as e:
                st.warning(f"Couldn't build the comparison: {e}")

overall = result['overall']
chats_df = result['chats']
calls_df = result['calls']
adherence_df = result['adherence']
unclassified = result['unclassified_shifts']
unmatched_ids = result['unmatched_chat_ids']
unattributed_calls = result['unattributed_calls']
chat_category_totals = result['chat_category_totals']

# ---------------------------------------------------------------------------
# Overall cards
# ---------------------------------------------------------------------------
st.header("Overall")

c1, c2, c3, c4 = st.columns(4)
c1.metric("Chats Closed", f"{overall['total_chats_closed']:,}")
c2.metric("Chats FCR Rate", _pct(overall['fcr_rate']), help="No return contact from the same customer within 7 days of the chat closing.")
c3.metric("Total Calls", f"{overall['total_calls']:,}")
c4.metric(
    "Calls Answered Rate", _pct(overall['answered_rate']),
    help="Serviced, as a share of ALL calls in the period -- including calls that never reached any agent "
         "(e.g. Dropped), which is why this is computed against the full Calls tab rather than the per-agent table.",
)

c5, c6, c7, c8 = st.columns(4)
c5.metric("Avg. Adherence", _pct(overall['avg_adherence']))
c6.metric("Avg. Occupancy", _pct(overall['avg_occupancy']), help="Calls-only metric: Busy time / (Busy + Online) time. Chat handling shows as Online, same as idle time, so this will always read low.")
c7.metric("Avg. Shrinkage", _pct(overall['avg_shrinkage']))
c8.metric("Agents in Scope", f"{overall['agents_in_scope']}")

st.subheader("Calls by state")
state_totals = overall['call_state_totals']
fig = go.Figure(go.Bar(
    x=list(state_totals.values()), y=list(state_totals.keys()), orientation='h',
    marker_color=[STATE_COLORS[s] for s in state_totals], text=[f"{v:,}" for v in state_totals.values()],
    textposition='outside', hovertemplate='%{y}: %{x:,}<extra></extra>',
))
fig.update_layout(height=300, margin=dict(l=10, r=10, t=10, b=10), xaxis_title="Calls", yaxis_title=None)
st.plotly_chart(fig, use_container_width=True)

# Opportunistic -- only appears when the live Chats tab has a "Conversation Category"
# column (see compute_chats in logic.py). Silently absent otherwise, no error shown.
if chat_category_totals:
    st.subheader("Chats by category")
    cats = dict(sorted(chat_category_totals.items(), key=lambda kv: kv[1]))
    fig_cat = go.Figure(go.Bar(
        x=list(cats.values()), y=list(cats.keys()), orientation='h',
        marker_color='#5B8DEF', text=[f"{v:,}" for v in cats.values()],
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
d1, d2, d3 = st.columns(3)
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

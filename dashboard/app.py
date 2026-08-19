"""
TVS Motor Company — Equity Research Dashboard
Reads pipeline-generated data/outputs.json and renders an interactive,
dark-themed research dashboard. Never crashes on missing/malformed data.
"""

from datetime import datetime, timezone
from pathlib import Path
import json

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

ROOT = Path(__file__).resolve().parent.parent
OUTPUTS_JSON = ROOT / "data" / "outputs.json"

# Fixed categorical palette, distinct hues in a set order (never cycled/rainbow)
CATEGORICAL_PALETTE = [
    "#3498DB", "#9B59B6", "#1ABC9C", "#F39C12",
    "#E74C3C", "#2ECC71", "#34495E",
]
GREEN = "#2ECC71"
RED = "#E74C3C"
BLUE = "#3498DB"
MUTED = "#8892A6"
SURFACE = "#12161F"
SURFACE_2 = "#1A1F2B"
TEXT_PRIMARY = "#E8EAED"
TEXT_MUTED = "#9AA3B2"

# Operating-case scenario -> badge CSS class (badge-buy=green, badge-hold=grey, badge-sell=red)
SCENARIO_BADGE_CLASS = {"Bull": "badge-buy", "Base": "badge-hold", "Bear": "badge-sell"}
SCENARIO_DISPLAY_ORDER = ["Bull", "Base", "Bear"]

st.set_page_config(
    page_title="TVS Motor — Equity Research Dashboard",
    page_icon="\U0001F3CD",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    f"""
    <style>
    .stApp {{ background-color: {SURFACE}; color: {TEXT_PRIMARY}; }}
    section[data-testid="stSidebar"] {{ background-color: {SURFACE_2}; }}
    div[data-testid="stMetric"] {{
        background-color: {SURFACE_2};
        border: 1px solid #2A3040;
        border-radius: 10px;
        padding: 14px 16px 10px 16px;
    }}
    div[data-testid="stMetricLabel"] {{ color: {TEXT_MUTED}; }}
    .badge {{
        display: inline-block; padding: 6px 18px; border-radius: 8px;
        font-weight: 700; font-size: 1.1rem; text-align: center; width: 100%;
    }}
    .badge-buy {{ background-color: rgba(46,204,113,0.18); color: {GREEN}; border: 1px solid {GREEN}; }}
    .badge-sell {{ background-color: rgba(231,76,60,0.18); color: {RED}; border: 1px solid {RED}; }}
    .badge-hold {{ background-color: rgba(154,163,178,0.18); color: {MUTED}; border: 1px solid {MUTED}; }}
    .footer-text {{ color: {TEXT_MUTED}; font-size: 0.85rem; }}
    </style>
    """,
    unsafe_allow_html=True,
)


def format_inr(value, decimals=2):
    """Indian number formatting. Crores/lakhs for large magnitudes, else lakh-style commas."""
    if value is None:
        return "N/A"
    try:
        value = float(value)
    except (TypeError, ValueError):
        return "N/A"

    sign = "-" if value < 0 else ""
    v = abs(value)

    if v >= 1e7:
        return f"{sign}₹{v / 1e7:,.{decimals}f} Cr"
    if v >= 1e5:
        return f"{sign}₹{v / 1e5:,.{decimals}f} L"

    s = f"{v:,.{decimals}f}"
    int_part, _, dec_part = s.partition(".")
    int_part = int_part.replace(",", "")
    if len(int_part) > 3:
        last3 = int_part[-3:]
        rest = int_part[:-3]
        groups = []
        while len(rest) > 2:
            groups.insert(0, rest[-2:])
            rest = rest[:-2]
        if rest:
            groups.insert(0, rest)
        int_part = ",".join(groups) + "," + last3
    formatted = int_part + (("." + dec_part) if dec_part else "")
    return f"{sign}₹{formatted}"


@st.cache_data(ttl=3600)
def load_data():
    with open(OUTPUTS_JSON, "r", encoding="utf-8") as f:
        return json.load(f)


try:
    data = load_data()
except Exception as e:
    st.error(
        "Couldn't load data/outputs.json. Run `python pipeline/refresh.py` first "
        f"to generate it.\n\nDetails: {e}"
    )
    st.stop()

if not isinstance(data, dict) or not data:
    st.error("data/outputs.json is empty or malformed. Re-run the pipeline to regenerate it.")
    st.stop()

market_inputs = data.get("market_inputs", {}) or {}
cmp_val = market_inputs.get("cmp")

# Backward compatible with a pre-scenario outputs.json (flat structure, no
# "scenarios" key) — treat the whole file as a single implicit Base scenario.
scenarios = data.get("scenarios") or {"Base": data}
scenario_order = [s for s in SCENARIO_DISPLAY_ORDER if s in scenarios] or list(scenarios.keys())
default_active = data.get("active_scenario") or "Base"
if scenario_order and default_active not in scenario_order:
    default_active = scenario_order[0]

# ---------------------------------------------------------------------------
# Sidebar — scenario controls
# ---------------------------------------------------------------------------
st.sidebar.header("Scenario Controls")

if scenario_order:
    selected_scenario = st.sidebar.selectbox(
        "Operating case",
        scenario_order,
        index=scenario_order.index(default_active),
        key="operating_case_selector",
        help="Switches every metric and chart below to that case's Excel-computed outputs.",
    )
else:
    selected_scenario = default_active

active = scenarios.get(selected_scenario) or {}

st.sidebar.divider()

scenario_switches = data.get("scenario_switches") or {}
other_switches = {name: meta for name, meta in scenario_switches.items() if name != "Operating case"}
if other_switches:
    for name, meta in other_switches.items():
        options = (meta or {}).get("valid_options") or []
        current = (meta or {}).get("current_value")
        if not options:
            continue
        try:
            idx = options.index(current) if current in options else 0
        except Exception:
            idx = 0
        st.sidebar.selectbox(name, options, index=idx, key=f"scenario_{name}")
    st.sidebar.caption("These reflect the Base case pipeline run")
else:
    st.sidebar.info("No scenario switch data available in outputs.json.")

# ---------------------------------------------------------------------------
# A. Header
# ---------------------------------------------------------------------------
title_col, badge_col = st.columns([5, 1])
with title_col:
    st.title("TVS Motor Company — Equity Research Dashboard")
    st.caption("NSE: TVSMOTOR · Sum-of-the-Parts · DCF · Relative Valuation · Monte Carlo")
with badge_col:
    badge_class = SCENARIO_BADGE_CLASS.get(selected_scenario, "badge-hold")
    st.markdown(
        f'<div class="badge {badge_class}" style="margin-top: 30px;">{selected_scenario.upper()} CASE</div>',
        unsafe_allow_html=True,
    )

last_refreshed = data.get("last_refreshed")
if last_refreshed:
    try:
        ts = datetime.fromisoformat(last_refreshed.replace("Z", "+00:00"))
        age_hours = (datetime.now(timezone.utc) - ts).total_seconds() / 3600
        st.markdown(f"**Last refreshed:** {ts.strftime('%d %b %Y, %H:%M UTC')}")
        if age_hours > 48:
            st.warning(f"⚠️ Data is {age_hours:.0f} hours old — the daily pipeline may not be running.")
    except Exception:
        st.markdown(f"**Last refreshed:** {last_refreshed}")
else:
    st.markdown("**Last refreshed:** unknown")

st.divider()

# ---------------------------------------------------------------------------
# B. Key metrics row (from the selected scenario, except CMP which is shared)
# ---------------------------------------------------------------------------
concluded = active.get("concluded_value_per_share")
upside = active.get("upside_pct")
wacc = (active.get("cost_of_capital") or {}).get("wacc")
recommendation = active.get("recommendation") or "N/A"

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("CMP (₹)", format_inr(cmp_val) if cmp_val is not None else "N/A")
c2.metric("Concluded Value (₹)", format_inr(concluded) if concluded is not None else "N/A")
c3.metric(
    "Upside / (Downside)",
    f"{upside * 100:,.1f}%" if isinstance(upside, (int, float)) else "N/A",
)
c4.metric("WACC", f"{wacc * 100:,.2f}%" if isinstance(wacc, (int, float)) else "N/A")

rec_upper = str(recommendation).strip().upper()
rec_badge_class = "badge-hold"
if rec_upper == "BUY":
    rec_badge_class = "badge-buy"
elif rec_upper == "SELL":
    rec_badge_class = "badge-sell"
with c5:
    st.markdown(f'<div class="badge {rec_badge_class}">{rec_upper}</div>', unsafe_allow_html=True)
    st.caption("Recommendation")

st.divider()

# ---------------------------------------------------------------------------
# D. SOTP waterfall
# ---------------------------------------------------------------------------
st.subheader("Sum-of-the-Parts Bridge (₹ per share)")
sotp_legs = active.get("sotp_legs") or []
if sotp_legs:
    names = [leg.get("name", "") for leg in sotp_legs]
    values = [leg.get("value_per_share") for leg in sotp_legs]
    types = [leg.get("type", "add") for leg in sotp_legs]

    measures = ["relative" for _ in sotp_legs] + ["total"]
    clean_values = [(-abs(v) if t == "subtract" and v is not None and v > 0 else v) for v, t in zip(values, types)]
    total_val = sum(v for v in clean_values if v is not None)

    waterfall_x = names + ["SOTP Value / Share"]
    waterfall_y = clean_values + [total_val]

    fig = go.Figure(
        go.Waterfall(
            x=waterfall_x,
            y=waterfall_y,
            measure=measures,
            text=[format_inr(v) for v in waterfall_y],
            textposition="outside",
            increasing={"marker": {"color": GREEN}},
            decreasing={"marker": {"color": RED}},
            totals={"marker": {"color": BLUE}},
            connector={"line": {"color": MUTED, "width": 1}},
        )
    )
    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor=SURFACE,
        plot_bgcolor=SURFACE,
        showlegend=False,
        margin=dict(t=20, b=20, l=20, r=20),
        height=420,
    )
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("No SOTP leg data available.")

# ---------------------------------------------------------------------------
# E. Football field
# ---------------------------------------------------------------------------
st.subheader("Valuation Football Field (₹ per share)")
football_field = active.get("football_field") or []
if football_field:
    methods = [f.get("methodology", "") for f in football_field]
    lows = [f.get("low") for f in football_field]
    mids = [f.get("mid") for f in football_field]
    highs = [f.get("high") for f in football_field]

    fig = go.Figure()
    for i, (m, lo, mid, hi) in enumerate(zip(methods, lows, mids, highs)):
        color = CATEGORICAL_PALETTE[i % len(CATEGORICAL_PALETTE)]
        if lo is None or hi is None:
            continue
        fig.add_trace(
            go.Scatter(
                x=[lo, hi],
                y=[m, m],
                mode="lines",
                line=dict(color=color, width=10),
                showlegend=False,
                hoverinfo="skip",
            )
        )
        if mid is not None:
            fig.add_trace(
                go.Scatter(
                    x=[mid],
                    y=[m],
                    mode="markers",
                    marker=dict(symbol="diamond", size=13, color=TEXT_PRIMARY, line=dict(color=color, width=2)),
                    name=m,
                    showlegend=False,
                    hovertemplate=f"{m}<br>Low: {format_inr(lo)}<br>Mid: {format_inr(mid)}<br>High: {format_inr(hi)}<extra></extra>",
                )
            )

    if cmp_val is not None:
        fig.add_vline(
            x=cmp_val,
            line_width=2,
            line_dash="dash",
            line_color=RED,
            annotation_text=f"CMP ₹{cmp_val:,.0f}",
            annotation_position="top",
            annotation_font_color=RED,
        )

    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor=SURFACE,
        plot_bgcolor=SURFACE,
        margin=dict(t=20, b=20, l=20, r=20),
        height=420,
        xaxis_title="₹ per share",
        yaxis_title=None,
    )
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("No football field data available.")

# ---------------------------------------------------------------------------
# F. Sensitivity heatmap
# ---------------------------------------------------------------------------
st.subheader("Implied Value per Share — WACC vs Terminal Growth Rate")
sens = active.get("sensitivity_grid") or {}
wacc_values = sens.get("wacc_values") or []
g_values = sens.get("g_values") or []
matrix = sens.get("matrix") or []

sensitivity_valid = (
    bool(wacc_values)
    and bool(g_values)
    and bool(matrix)
    and all(v is not None for v in wacc_values)
    and all(v is not None for v in g_values)
    and all(v is not None for row in matrix for v in row)
)

if sensitivity_valid:
    matrix_arr = np.array(matrix, dtype=float)
    center = cmp_val if cmp_val is not None else float(np.nanmedian(matrix_arr))
    max_dev = max(abs(np.nanmax(matrix_arr) - center), abs(np.nanmin(matrix_arr) - center), 1e-6)

    fig = go.Figure(
        data=go.Heatmap(
            z=matrix_arr,
            x=[f"{g*100:.1f}%" for g in g_values],
            y=[f"{w*100:.2f}%" for w in wacc_values],
            colorscale=[[0, RED], [0.5, "#2A3040"], [1, GREEN]],
            zmid=center,
            zmin=center - max_dev,
            zmax=center + max_dev,
            text=[[f"₹{v:,.0f}" for v in row] for row in matrix_arr],
            texttemplate="%{text}",
            textfont={"size": 11, "color": TEXT_PRIMARY},
            colorbar=dict(title="₹/share"),
            hovertemplate="WACC: %{y}<br>Terminal g: %{x}<br>Value: %{text}<extra></extra>",
        )
    )
    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor=SURFACE,
        plot_bgcolor=SURFACE,
        margin=dict(t=20, b=20, l=20, r=20),
        height=420,
        xaxis_title="Terminal growth rate (g)",
        yaxis_title="WACC",
    )
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("Sensitivity data not available in this refresh")

# ---------------------------------------------------------------------------
# G. Monte Carlo distribution
# ---------------------------------------------------------------------------
st.subheader("Monte Carlo Simulation — Implied Share Price Distribution")
mc = active.get("monte_carlo") or {}
raw_trials = mc.get("raw_trials")
mean = mc.get("mean")
std_dev = mc.get("std_dev")
p10, p25, p50, p75, p90 = mc.get("p10"), mc.get("p25"), mc.get("p50"), mc.get("p75"), mc.get("p90")
prob_above_cmp = mc.get("prob_above_cmp")

sample = None
if raw_trials:
    sample = np.array(raw_trials, dtype=float)
elif mean is not None and std_dev is not None:
    rng = np.random.default_rng(42)
    sample = rng.normal(loc=mean, scale=std_dev, size=5000)

if sample is not None and len(sample) > 0:
    fig = go.Figure()
    fig.add_trace(
        go.Histogram(
            x=sample,
            nbinsx=40,
            marker_color=BLUE,
            opacity=0.85,
            name="Simulated value / share",
        )
    )
    vlines = [("CMP", cmp_val, RED), ("P10", p10, MUTED), ("Median", p50, TEXT_PRIMARY), ("P90", p90, MUTED)]
    for label, x, color in vlines:
        if x is None:
            continue
        fig.add_vline(x=x, line_width=2, line_dash="dash", line_color=color,
                       annotation_text=label, annotation_position="top")
    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor=SURFACE,
        plot_bgcolor=SURFACE,
        margin=dict(t=20, b=20, l=20, r=20),
        height=420,
        xaxis_title="₹ per share",
        yaxis_title="Frequency",
        showlegend=False,
    )
    st.plotly_chart(fig, use_container_width=True)

    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric("P10", format_inr(p10))
    m2.metric("P25", format_inr(p25))
    m3.metric("Median", format_inr(p50))
    m4.metric("P75", format_inr(p75))
    m5.metric("P90", format_inr(p90))
    m6.metric("Prob > CMP", f"{prob_above_cmp * 100:,.1f}%" if isinstance(prob_above_cmp, (int, float)) else "N/A")
else:
    st.info("No Monte Carlo data available.")

# ---------------------------------------------------------------------------
# H. Peer snapshot table (shared across scenarios)
# ---------------------------------------------------------------------------
st.subheader("Peer Snapshot")
peers = market_inputs.get("peers") or {}
rows = []
# TVS Motor's own trading multiples aren't captured in market_inputs; shown blank.
rows.append({"Company": "TVS Motor", "CMP (₹)": cmp_val, "EV/EBITDA": None, "P/E": None})

for key, meta in peers.items():
    rows.append(
        {
            "Company": meta.get("name", key),
            "CMP (₹)": meta.get("price"),
            "EV/EBITDA": meta.get("ev_ebitda"),
            "P/E": meta.get("pe"),
        }
    )

if rows:
    df = pd.DataFrame(rows)

    def highlight_tvs(row):
        if row["Company"] == "TVS Motor":
            return ["background-color: rgba(52,152,219,0.15)"] * len(row)
        return [""] * len(row)

    styled = df.style.apply(highlight_tvs, axis=1).format(
        {"CMP (₹)": "{:,.2f}", "EV/EBITDA": "{:,.2f}", "P/E": "{:,.2f}"}, na_rep="N/A"
    )
    st.dataframe(styled, use_container_width=True, hide_index=True)
else:
    st.info("No peer data available.")

# ---------------------------------------------------------------------------
# I. Footer
# ---------------------------------------------------------------------------
st.divider()
st.markdown(
    '<p class="footer-text">Built by Niranjan Desai · For illustrative purposes only '
    "— not investment advice.</p>",
    unsafe_allow_html=True,
)

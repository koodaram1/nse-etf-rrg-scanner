
import tempfile
import pandas as pd
import streamlit as st

from scanner_engine import run_scanner
from mobile_ui import (
    inject_css, fmt_date, pct, n, rrg_html, safe, metric_rows,
    card, signal_class, etf_class, section, empty_message,
    normal_multiple_from_row
)

st.set_page_config(
    page_title="NSE ETF RRG — Dashboard",
    page_icon="📈",
    layout="centered",
    initial_sidebar_state="collapsed",
)
inject_css()

scan = st.session_state.get("scan")
run_date = scan.get("run_date") if scan else None

st.title("NSE ETF RRG — DASHBOARD")
st.markdown(f'<div class="mobile-date">{fmt_date(run_date)}</div>', unsafe_allow_html=True)

if st.button("RUN SCANNER", type="primary"):
    with st.spinner("Running NSE ETF scanner..."):
        runtime_dir = tempfile.mkdtemp(prefix="nse_etf_rrg_")
        try:
            st.session_state["scan"] = run_scanner(runtime_dir)
            scan = st.session_state["scan"]
            st.success("Scanner completed.")
        except Exception as e:
            st.error(f"Scanner could not complete: {e}")
            st.stop()

if not scan:
    st.caption("Tap RUN SCANNER to load today's dashboard.")
    st.stop()

# ---------------------------
# TOP 3 EQUITY GROUPS
# ---------------------------
section("TOP 3 EQUITY GROUPS")
groups = scan["top_groups"].head(3)
for _, row in groups.iterrows():
    group = safe(row, "Theme", "-")
    q = safe(row, "SectorRRG", "-")
    strength = safe(row, "SectorStrengthScore")
    today = safe(row, "TodaySectorPct")
    html = (
        f'<div class="group-name">{group}</div>'
        f'<div class="muted">{rrg_html(q)}</div>'
        + metric_rows([
            ("Today", pct(today)),
            ("Strength", n(strength, 1)),
        ])
    )
    card(html, "info")

# ---------------------------
# SWING BUY PRIORITY
# ---------------------------
section("SWING BUY PRIORITY")
swing = scan["swing_buys"].head(3)
if swing.empty:
    empty_message("No strict Swing Buy candidate now.")
else:
    for i, (_, row) in enumerate(swing.iterrows(), start=1):
        etf = safe(row, "Symbol", "-")
        theme = safe(row, "Theme", "-")
        signal = safe(row, "BuyPrioritySignal", "-")
        html = (
            f'<div class="{etf_class("buy")}">{"🥇 " if i == 1 else ""}{etf}</div>'
            f'<div class="muted">{theme} • {rrg_html(safe(row, "SectorRRG", "-"))}</div>'
            + metric_rows([
                ("Today", pct(safe(row, "NSE_ChangePct"))),
                ("Volume Pace", n(safe(row, "VolumePaceVs30D"), 2, "x")),
                ("Near Day High", n(safe(row, "RangePositionPct"), 0, "%")),
                ("ETF RRG", rrg_html(safe(row, "Quadrant", "-"))),
                ("Trend", f'{int(safe(row, "TrendPoints", 0))}/4'),
                ("Momentum", str(safe(row, "MomentumPhase", "-"))),
                ("Entry", str(safe(row, "EntryStretch", "-"))),
                ("Buy Quality", n(safe(row, "BuyQualityScore"), 1)),
            ])
            + f'<div class="signal {signal_class(signal)}">{signal}</div>'
        )
        card(html, "buy")

# ---------------------------
# NEAR BUY WATCHLIST
# ---------------------------
section("NEAR BUY WATCHLIST")
near = scan["near_buys"].head(3)
if near.empty:
    empty_message("No near-buy Swing candidate.")
else:
    for _, row in near.iterrows():
        etf = safe(row, "Symbol", "-")
        theme = safe(row, "Theme", "-")
        why = safe(row, "WhyNotQualified", "-")
        html = (
            f'<div class="{etf_class("watch")}">{etf}</div>'
            f'<div class="muted">{theme} • {rrg_html(safe(row, "SectorRRG", "-"))}</div>'
            + metric_rows([
                ("Today", pct(safe(row, "NSE_ChangePct"))),
                ("Near Day High", n(safe(row, "RangePositionPct"), 0, "%")),
                ("Buy Quality", n(safe(row, "BuyQualityScore"), 1)),
            ])
            + f'<div class="signal signal-watch">Why not: {why}</div>'
        )
        card(html, "watch")

# ---------------------------
# NEGATIVE / REBOUND
# ---------------------------
section("NEGATIVE / REBOUND")
rebound = scan["rebound"].head(3)
if rebound.empty:
    empty_message("No liquidity-qualified negative Equity / International ETF.")
else:
    for i, (_, row) in enumerate(rebound.iterrows(), start=1):
        signal = safe(row, "DipSignal", "-")
        qualified = bool(safe(row, "ReboundQualified", False))
        kind = "buy" if qualified else "watch"
        etf = safe(row, "Symbol", "-")
        html = (
            f'<div class="{etf_class(kind)}">{"🥇 " if i == 1 and qualified else ""}{etf}</div>'
            f'<div class="muted">{safe(row, "Theme", "-")} • Group {rrg_html(safe(row, "Theme_Quadrant", "-"))} • ETF {rrg_html(safe(row, "Quadrant", "-"))}</div>'
            + metric_rows([
                ("Today", pct(safe(row, "NSE_ChangePct"))),
                ("Recovery", n(safe(row, "RangePositionPct"), 0, "%")),
                ("Vs Open", pct(safe(row, "ReboundVsOpenPct"))),
                ("Trend", f'{int(safe(row, "TrendPoints", 0))}/4'),
                ("Rebound Score", n(safe(row, "ReboundScore"), 1)),
            ])
            + f'<div class="signal {signal_class(signal)}">{signal}</div>'
        )
        card(html, kind)

# ---------------------------
# DEEP FALL WATCH
# ---------------------------
deep = scan["deep_fall"].head(3)
if not deep.empty:
    section("DEEP FALL WATCH ≥1%")
    for _, row in deep.iterrows():
        etf = safe(row, "Symbol", "-")
        signal = safe(row, "DipSignal", "-")
        html = (
            f'<div class="{etf_class("danger")}">{etf}</div>'
            f'<div class="muted">{safe(row, "Theme", "-")} • {rrg_html(safe(row, "Quadrant", "-"))}</div>'
            + metric_rows([
                ("Today", pct(safe(row, "NSE_ChangePct"))),
                ("Recovery", n(safe(row, "RangePositionPct"), 1, "%")),
                ("Trend", f'{int(safe(row, "TrendPoints", 0))}/4'),
            ])
            + f'<div class="signal {signal_class(signal)}">{signal}</div>'
        )
        card(html, "danger")

# ---------------------------
# UNUSUAL VOLUME
# ---------------------------
section("TOP 3 UNUSUAL VOLUME")
uv = scan["unusual_volume"].head(3)
if uv.empty:
    empty_message("No unusual-volume ETF available.")
else:
    for _, row in uv.iterrows():
        etf = safe(row, "Symbol", "-")
        multiple = normal_multiple_from_row(row)
        flag = safe(row, "VolumeFlag", "-")
        html = (
            f'<div class="{etf_class("info")}">{etf}</div>'
            f'<div class="muted">{safe(row, "Theme", "-")}</div>'
            + metric_rows([
                ("Today", pct(safe(row, "ChgPct"))),
                ("Volume", n(multiple, 2, "x normal")),
                ("Turnover", f'₹{n(safe(row, "TurnoverCr"), 2)} Cr'),
            ])
            + f'<div class="signal signal-info">{flag}</div>'
        )
        card(html, "info")


import streamlit as st

from mobile_ui import (
    inject_css, fmt_date, pct, n, rrg_html, safe, metric_rows,
    card, signal_class, etf_class, section, empty_message
)

st.set_page_config(
    page_title="NSE ETF RRG — Intraday",
    page_icon="⚡",
    layout="centered",
    initial_sidebar_state="collapsed",
)
inject_css()

scan = st.session_state.get("scan")
run_date = scan.get("run_date") if scan else None

st.markdown(
    '<div style="font-size:1.45rem;font-weight:800;line-height:1.35;'
    'padding:0.35rem 0 0.20rem 0;margin:0;overflow:visible;'
    'color:#F8FAFC !important;display:block !important;">'
    'NSE ETF RRG — INTRADAY</div>',
    unsafe_allow_html=True,
)
st.markdown(f'<div class="mobile-date">{fmt_date(run_date)}</div>', unsafe_allow_html=True)

if not scan:
    st.info("Run the scanner from the Dashboard first.")
    st.stop()

section("TOP 3 INTRADAY GROUPS")
groups = scan["intraday_groups"].head(3)
for _, row in groups.iterrows():
    html = (
        f'<div class="group-name">{safe(row, "Theme", "-")}</div>'
        f'<div class="muted">{rrg_html(safe(row, "SectorRRG", "-"))}</div>'
        + metric_rows([
            ("Today", pct(safe(row, "TodaySectorPct"))),
            ("Volume Pace", n(safe(row, "MedianVolumeVs30D"), 2, "x")),
        ])
    )
    card(html, "info")

section("INTRADAY BUY PRIORITY")
buys = scan["intraday_buys"].head(3)
if buys.empty:
    empty_message("NO QUALIFIED INTRADAY BUY NOW")
else:
    for i, (_, row) in enumerate(buys.iterrows(), start=1):
        signal = safe(row, "IntradaySignal", "-")
        etf = safe(row, "Symbol", "-")
        html = (
            f'<div class="{etf_class("buy")}">{"🥇 " if i == 1 else ""}{etf}</div>'
            f'<div class="muted">{safe(row, "Theme", "-")} • {rrg_html(safe(row, "SectorRRG", "-"))}</div>'
            + metric_rows([
                ("Today", pct(safe(row, "NSE_ChangePct"))),
                ("Vs Open", pct(safe(row, "IntradayOpenDrivePct"))),
                ("Volume Pace", n(safe(row, "VolumePaceVs30D"), 2, "x")),
                ("Near Day High", n(safe(row, "RangePositionPct"), 0, "%")),
                ("Tradability", str(safe(row, "IntradayTradability", "-"))),
                ("ETF RRG", rrg_html(safe(row, "Quadrant", "-"))),
                ("Intraday Score", n(safe(row, "IntradayBuyScore"), 1)),
            ])
            + f'<div class="signal {signal_class(signal)}">{signal}</div>'
        )
        card(html, "buy")

section("INTRADAY NEAR BUY")
near = scan["intraday_near"].head(3)
if near.empty:
    empty_message("No near-buy Intraday candidate.")
else:
    for _, row in near.iterrows():
        etf = safe(row, "Symbol", "-")
        why = safe(row, "WhyNotIntradayQualified", "-")
        html = (
            f'<div class="{etf_class("watch")}">{etf}</div>'
            f'<div class="muted">{safe(row, "Theme", "-")}</div>'
            + metric_rows([
                ("Today", pct(safe(row, "NSE_ChangePct"))),
                ("Volume Pace", n(safe(row, "VolumePaceVs30D"), 2, "x")),
                ("Near Day High", n(safe(row, "RangePositionPct"), 0, "%")),
                ("Intraday Score", n(safe(row, "IntradayBuyScore"), 1)),
            ])
            + f'<div class="signal signal-watch">Why not: {why}</div>'
        )
        card(html, "watch")

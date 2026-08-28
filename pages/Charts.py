import math
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from mobile_ui import inject_css, fmt_date, rrg_html, safe, section

st.set_page_config(
    page_title="NSE ETF RRG — Charts",
    page_icon="📊",
    layout="centered",
    initial_sidebar_state="collapsed",
)
inject_css()

scan = st.session_state.get("scan")
run_date = scan.get("run_date") if scan else None

# Keep the same top spacing used on the working Dashboard / Intraday pages.
st.markdown("<div style='height:30px'></div>", unsafe_allow_html=True)
st.subheader("NSE ETF RRG — CHARTS")
st.markdown(f'<div class="mobile-date">{fmt_date(run_date)}</div>', unsafe_allow_html=True)

if not scan:
    st.info("Run the scanner from the Dashboard first.")
    st.stop()

st.write("SCAN KEYS:", sorted(scan.keys()))

st.caption("RRG charts based on the latest completed scan. Benchmark: NIFTY 50")

QUAD_COLORS = {
    "LEADING": "#67C587",
    "IMPROVING": "#5B8FF9",
    "WEAKENING": "#F5B52E",
    "LAGGING": "#FF5C5C",
    "NO DATA": "#9CA3AF",
}


def _finite(v):
    try:
        return math.isfinite(float(v))
    except Exception:
        return False


def _axis_bounds(points):
    vals = [float(v) for v in points if _finite(v)]
    if not vals:
        return 90, 110
    lo = min(vals + [100.0])
    hi = max(vals + [100.0])
    pad = max(2.5, (hi - lo) * 0.18)
    return lo - pad, hi + pad


def _rrg_figure(latest_df, histories, key_col, title):
    latest_df = latest_df.copy()
    latest_df["RS_Ratio"] = pd.to_numeric(latest_df.get("RS_Ratio"), errors="coerce")
    latest_df["RS_Momentum"] = pd.to_numeric(latest_df.get("RS_Momentum"), errors="coerce")
    latest_df = latest_df.dropna(subset=["RS_Ratio", "RS_Momentum"])

    all_x = latest_df["RS_Ratio"].tolist()
    all_y = latest_df["RS_Momentum"].tolist()
    for key in latest_df[key_col].astype(str):
        h = histories.get(key)
        if isinstance(h, pd.DataFrame) and not h.empty:
            all_x += pd.to_numeric(h.get("RS_Ratio"), errors="coerce").dropna().tolist()
            all_y += pd.to_numeric(h.get("RS_Momentum"), errors="coerce").dropna().tolist()

    xmin, xmax = _axis_bounds(all_x)
    ymin, ymax = _axis_bounds(all_y)

    fig = go.Figure()

    # Soft quadrant backgrounds.
    fig.add_shape(type="rect", x0=100, x1=xmax, y0=100, y1=ymax,
                  fillcolor="rgba(44,160,70,0.10)", line_width=0, layer="below")
    fig.add_shape(type="rect", x0=xmin, x1=100, y0=100, y1=ymax,
                  fillcolor="rgba(49,102,235,0.10)", line_width=0, layer="below")
    fig.add_shape(type="rect", x0=100, x1=xmax, y0=ymin, y1=100,
                  fillcolor="rgba(245,181,46,0.10)", line_width=0, layer="below")
    fig.add_shape(type="rect", x0=xmin, x1=100, y0=ymin, y1=100,
                  fillcolor="rgba(255,92,92,0.10)", line_width=0, layer="below")

    fig.add_vline(x=100, line_width=1, line_dash="dot", line_color="#8B949E")
    fig.add_hline(y=100, line_width=1, line_dash="dot", line_color="#8B949E")

    # Trails + current point.
    for _, row in latest_df.iterrows():
        key = str(row[key_col])
        quad = str(row.get("Quadrant", "NO DATA")).upper()
        color = QUAD_COLORS.get(quad, QUAD_COLORS["NO DATA"])
        hist = histories.get(key)

        if isinstance(hist, pd.DataFrame) and not hist.empty:
            hx = pd.to_numeric(hist.get("RS_Ratio"), errors="coerce")
            hy = pd.to_numeric(hist.get("RS_Momentum"), errors="coerce")
            mask = hx.notna() & hy.notna()
            hx, hy = hx[mask], hy[mask]
            if len(hx):
                fig.add_trace(go.Scatter(
                    x=hx, y=hy, mode="lines+markers",
                    line=dict(color=color, width=1.5),
                    marker=dict(size=4, color=color),
                    hoverinfo="skip", showlegend=False,
                ))

        fig.add_trace(go.Scatter(
            x=[row["RS_Ratio"]],
            y=[row["RS_Momentum"]],
            mode="markers+text",
            text=[key],
            textposition="top center",
            textfont=dict(size=11, color=color),
            marker=dict(size=10, color=color, line=dict(width=1, color="#E5E7EB")),
            customdata=[[quad]],
            hovertemplate=(
                f"<b>{key}</b><br>RS Ratio: %{{x:.2f}}<br>"
                "RS Momentum: %{y:.2f}<br>Quadrant: %{customdata[0]}<extra></extra>"
            ),
            showlegend=False,
        ))

    # Quadrant labels.
    fig.add_annotation(x=xmax, y=ymax, text="LEADING", showarrow=False,
                       xanchor="right", yanchor="top", font=dict(color=QUAD_COLORS["LEADING"], size=11))
    fig.add_annotation(x=xmin, y=ymax, text="IMPROVING", showarrow=False,
                       xanchor="left", yanchor="top", font=dict(color=QUAD_COLORS["IMPROVING"], size=11))
    fig.add_annotation(x=xmax, y=ymin, text="WEAKENING", showarrow=False,
                       xanchor="right", yanchor="bottom", font=dict(color=QUAD_COLORS["WEAKENING"], size=11))
    fig.add_annotation(x=xmin, y=ymin, text="LAGGING", showarrow=False,
                       xanchor="left", yanchor="bottom", font=dict(color=QUAD_COLORS["LAGGING"], size=11))

    fig.update_layout(
        title=dict(text=title, font=dict(size=15)),
        height=520,
        margin=dict(l=25, r=20, t=45, b=25),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#E5E7EB", size=11),
        xaxis=dict(title="RS Ratio", range=[xmin, xmax], gridcolor="rgba(148,163,184,0.12)", zeroline=False),
        yaxis=dict(title="RS Momentum", range=[ymin, ymax], gridcolor="rgba(148,163,184,0.12)", zeroline=False),
        hovermode="closest",
    )
    return fig


# ============================================================
# 1) EQUITY THEME RRG
# ============================================================
section("1. EQUITY THEME RRG")
theme_rrg = scan.get("theme_rrg", pd.DataFrame()).copy()
theme_hist = scan.get("theme_rrg_history", {}) or {}

if not theme_rrg.empty:
    if "AssetClass" in theme_rrg.columns:
        ac = theme_rrg["AssetClass"].astype(str).str.upper()
        equity_theme = theme_rrg[ac.isin(["EQUITY", "INTERNATIONAL"])].copy()
    else:
        equity_theme = theme_rrg.copy()

    # Mobile: focus on the most relevant groups, but retain all Leading/Improving
    # plus the highest-strength remaining themes, capped to avoid label clutter.
    wanted = []
    top_groups = scan.get("top_groups", pd.DataFrame())
    if isinstance(top_groups, pd.DataFrame) and "Theme" in top_groups.columns:
        wanted += top_groups["Theme"].astype(str).tolist()
    strong = equity_theme[equity_theme["Quadrant"].astype(str).str.upper().isin(["LEADING", "IMPROVING"])]
    wanted += strong["Theme"].astype(str).tolist()
    wanted = list(dict.fromkeys(wanted))[:12]
    theme_show = equity_theme[equity_theme["Theme"].astype(str).isin(wanted)].copy()
    if theme_show.empty:
        theme_show = equity_theme.head(12).copy()

    st.plotly_chart(
        _rrg_figure(theme_show, theme_hist, "Theme", "Equity themes vs NIFTY 50"),
        use_container_width=True,
        config={"displaylogo": False, "responsive": True},
    )
    st.caption("Trail = latest 5 trading days. Tap a point for RS Ratio / RS Momentum.")

    # Small Top 3 summary beneath chart.
    top3 = scan.get("top_groups", pd.DataFrame()).head(3)
    if isinstance(top3, pd.DataFrame) and not top3.empty:
        st.markdown("**TOP 3 EQUITY GROUPS**")
        for _, r in top3.iterrows():
            nm = str(safe(r, "Theme", "-"))
            q = str(safe(r, "SectorRRG", "-"))
            match = theme_rrg[theme_rrg["Theme"].astype(str).eq(nm)]
            rsr = match.iloc[0]["RS_Ratio"] if not match.empty else None
            rsm = match.iloc[0]["RS_Momentum"] if not match.empty else None
            left, mid, right = st.columns([1.4, 1, 1])
            left.markdown(f"**{nm}**  \n{rrg_html(q)}", unsafe_allow_html=True)
            mid.metric("RS Ratio", f"{float(rsr):.2f}" if _finite(rsr) else "-")
            right.metric("RS Mom", f"{float(rsm):.2f}" if _finite(rsm) else "-")
else:
    st.info("Theme RRG data is not available in this scan.")

st.divider()

# ============================================================
# 2) ETF RRG — IMPORTANT CANDIDATES
# ============================================================
section("2. ETF RRG — IMPORTANT CANDIDATES")
etf_rrg = scan.get("etf_rrg", pd.DataFrame()).copy()
etf_hist = scan.get("etf_rrg_history", {}) or {}

candidate_symbols = []
for key in ["swing_buys", "near_buys"]:
    df = scan.get(key, pd.DataFrame())
    if isinstance(df, pd.DataFrame) and "Symbol" in df.columns:
        candidate_symbols += df["Symbol"].dropna().astype(str).head(5).tolist()

# Add best ETF from each of today's Top 3 groups.
best_by_theme = scan.get("best_by_theme", pd.DataFrame())
top_groups = scan.get("top_groups", pd.DataFrame()).head(3)
if isinstance(best_by_theme, pd.DataFrame) and not best_by_theme.empty and "Theme" in best_by_theme.columns:
    for theme in top_groups.get("Theme", pd.Series(dtype=str)).astype(str):
        m = best_by_theme[best_by_theme["Theme"].astype(str).eq(theme)]
        if not m.empty and "Symbol" in m.columns:
            candidate_symbols.append(str(m.iloc[0]["Symbol"]))

candidate_symbols = list(dict.fromkeys(candidate_symbols))[:10]

if not etf_rrg.empty and candidate_symbols:
    etf_show = etf_rrg[etf_rrg["Symbol"].astype(str).isin(candidate_symbols)].copy()
    if not etf_show.empty:
        st.plotly_chart(
            _rrg_figure(etf_show, etf_hist, "Symbol", "Important ETF candidates vs NIFTY 50"),
            use_container_width=True,
            config={"displaylogo": False, "responsive": True},
        )
        st.caption("Candidates = Swing Buy + Near Buy + best ETF from leading groups. Trail = latest 5 trading days.")

        # Compact current-position table.
        summary = etf_show[["Symbol", "RS_Ratio", "RS_Momentum", "Quadrant"]].copy()
        summary["RS_Ratio"] = pd.to_numeric(summary["RS_Ratio"], errors="coerce").round(2)
        summary["RS_Momentum"] = pd.to_numeric(summary["RS_Momentum"], errors="coerce").round(2)
        summary = summary.rename(columns={"Symbol": "ETF", "RS_Ratio": "RS Ratio", "RS_Momentum": "RS Mom", "Quadrant": "RRG"})
        st.dataframe(summary, hide_index=True, use_container_width=True)
    else:
        st.info("No current Swing / Near Buy ETF has valid RRG data in this scan.")
else:
    st.info("ETF RRG candidate data is not available in this scan.")

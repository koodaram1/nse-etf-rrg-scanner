
import math
from datetime import datetime
from zoneinfo import ZoneInfo
import pandas as pd
import streamlit as st


def inject_css():
    st.markdown(
        """
        <style>
        .block-container {
            max-width: 680px;
            padding-top: 0.75rem;
            padding-left: 0.85rem;
            padding-right: 0.85rem;
            padding-bottom: 2rem;
        }
        h1 {
            font-size: 1.45rem !important;
            line-height: 1.35 !important;
            margin-top: 0 !important;
            margin-bottom: 0.20rem !important;
            padding-top: 0.30rem !important;
            padding-bottom: 0.18rem !important;
            overflow: visible !important;
        }
        h2 { font-size: 1.15rem !important; margin-top: 1.15rem !important; }
        .mobile-date {
            color: #64748b;
            font-size: 0.90rem;
            margin-bottom: 0.8rem;
        }
        .section-title {
            font-size: 0.98rem;
            font-weight: 800;
            letter-spacing: 0.02rem;
            margin-top: 1rem;
            margin-bottom: 0.45rem;
        }
        .card {
            background: white;
            border: 1px solid #dbe3ee;
            border-radius: 14px;
            padding: 0.78rem 0.88rem;
            margin-bottom: 0.65rem;
            box-shadow: 0 1px 3px rgba(15,23,42,.05);
        }
        .card-buy { border-left: 5px solid #16a34a; }
        .card-watch { border-left: 5px solid #f59e0b; }
        .card-danger { border-left: 5px solid #dc2626; }
        .card-info { border-left: 5px solid #2563eb; }
        .etf-buy { color: #15803d; font-size: 1.16rem; font-weight: 900; }
        .etf-watch { color: #b45309; font-size: 1.16rem; font-weight: 900; }
        .etf-danger { color: #dc2626; font-size: 1.16rem; font-weight: 900; }
        .etf-info { color: #1d4ed8; font-size: 1.16rem; font-weight: 900; }
        .group-name { color: #0f172a; font-size: 1.05rem; font-weight: 850; }
        .muted { color: #64748b; font-size: 0.82rem; }
        .kv {
            display: grid;
            grid-template-columns: 1fr auto;
            gap: 0.18rem 0.75rem;
            margin-top: 0.55rem;
            font-size: 0.88rem;
        }
        .kv .label { color: #64748b; }
        .kv .value { color: #0f172a; font-weight: 750; text-align: right; }
        .signal {
            margin-top: 0.62rem;
            padding: 0.42rem 0.55rem;
            border-radius: 8px;
            font-weight: 850;
            font-size: 0.86rem;
        }
        .signal-buy { background:#dcfce7; color:#166534; }
        .signal-watch { background:#fef3c7; color:#92400e; }
        .signal-danger { background:#fee2e2; color:#991b1b; }
        .signal-info { background:#dbeafe; color:#1e40af; }
        .rrg-leading { color:#15803d; font-weight:800; }
        .rrg-improving { color:#1d4ed8; font-weight:800; }
        .rrg-weakening { color:#b45309; font-weight:800; }
        .rrg-lagging { color:#dc2626; font-weight:800; }
        div[data-testid="stButton"] button {
            width: 100%;
            min-height: 2.8rem;
            border-radius: 10px;
            font-weight: 800;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def fmt_date(run_date=None):
    if run_date:
        try:
            return datetime.strptime(str(run_date), "%Y-%m-%d").strftime("%d-%b-%Y")
        except Exception:
            pass
    return datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%d-%b-%Y")


def n(v, digits=2, suffix=""):
    try:
        if pd.isna(v):
            return "-"
        return f"{float(v):.{digits}f}{suffix}"
    except Exception:
        return "-"


def pct(v):
    try:
        if pd.isna(v):
            return "-"
        x = float(v)
        return f"{x:+.2f}%"
    except Exception:
        return "-"


def rrg_html(v):
    s = str(v or "").upper()
    cls = {
        "LEADING":"rrg-leading",
        "IMPROVING":"rrg-improving",
        "WEAKENING":"rrg-weakening",
        "LAGGING":"rrg-lagging",
    }.get(s, "")
    icon = {
        "LEADING":"●",
        "IMPROVING":"●",
        "WEAKENING":"●",
        "LAGGING":"●",
    }.get(s, "")
    return f'<span class="{cls}">{icon} {s or "-"}</span>'


def safe(row, key, default=None):
    try:
        v = row.get(key, default)
        if pd.isna(v):
            return default
        return v
    except Exception:
        return default


def metric_rows(items):
    return '<div class="kv">' + ''.join(
        f'<div class="label">{label}</div><div class="value">{value}</div>'
        for label, value in items
    ) + '</div>'


def card(html, kind="info"):
    st.markdown(f'<div class="card card-{kind}">{html}</div>', unsafe_allow_html=True)


def signal_class(signal):
    s = str(signal or "").upper()
    if "TOP" in s or "STRONG" in s or "GOOD" in s:
        return "signal-buy"
    if "FALLING" in s or "WEAK" in s or "WAIT" in s:
        return "signal-danger"
    if "WATCH" in s:
        return "signal-watch"
    return "signal-info"


def etf_class(kind):
    return {
        "buy":"etf-buy",
        "watch":"etf-watch",
        "danger":"etf-danger",
        "info":"etf-info",
    }.get(kind, "etf-info")


def section(title):
    st.markdown(f'<div class="section-title">{title}</div>', unsafe_allow_html=True)


def empty_message(text):
    st.info(text)


def normal_multiple_from_row(row):
    try:
        day = float(safe(row, "DayVolume"))
        avg = float(safe(row, "Avg30Volume"))
        if avg > 0:
            return day / avg
    except Exception:
        pass
    return None

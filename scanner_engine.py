"""
NSE ETF RRG Scanner Engine
Mobile build V1.1 — intraday scope fix

Derived from the approved V4.8.3 desktop scanner.
Trading/scoring/qualification logic is unchanged.
Fix: notebook variables wrapped inside run_scanner() are function-local, so
legacy existence checks now use locals() instead of globals().
"""

import os
import sys
import warnings
import math
import time
import json
import re
import zipfile
from pathlib import Path
from io import BytesIO, StringIO
from datetime import datetime, timezone
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
import requests
from bs4 import BeautifulSoup
import yfinance as yf
import plotly.graph_objects as go
from plotly.subplots import make_subplots

warnings.filterwarnings("ignore")


def run_scanner(root_dir=None):
    # -----------------------------
    # STEP 1 — Runtime folders
    # -----------------------------
    # Mobile deployment does not depend on Google Drive.
    # The scanner writes temporary/output files inside root_dir.
    ROOT = Path(root_dir) if root_dir else (Path.cwd() / "NSE_ETF_RRG_RUNTIME")

    DATA_DIR = ROOT / "DATA"
    OUTPUT_DIR = ROOT / "OUTPUT"
    CHART_DIR = OUTPUT_DIR / "CHARTS"
    EXCEL_DIR = OUTPUT_DIR / "EXCEL"
    DASH_DIR = OUTPUT_DIR / "DASHBOARD"
    LOG_DIR = ROOT / "LOGS"

    for p in [ROOT, DATA_DIR, OUTPUT_DIR, CHART_DIR, EXCEL_DIR, DASH_DIR, LOG_DIR]:
        p.mkdir(parents=True, exist_ok=True)

    RUN_TS = datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%Y%m%d_%H%M%S")
    RUN_DATE = datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%Y-%m-%d")

    # -----------------------------
    # USER SETTINGS
    # -----------------------------
    BENCHMARK_YF = "^NSEI"          # NIFTY 50
    BENCHMARK_LABEL = "NIFTY 50"
    LOOKBACK_DAYS = 520             # enough for 200DMA + RRG history
    RRG_RS_WINDOW = 60
    RRG_MOM_LAG = 10
    RRG_SMOOTH = 5
    RRG_TRAIL_DAYS = 5
    TOP_THEME_CHART = 24
    TOP_ETF_CHART = 30
    DOWNLOAD_BATCH = 35
    MIN_RRG_OBS = 80

    print("\n" + "="*72)
    print("NSE ETF RRG SCANNER")
    print("="*72)
    print("Run date              :", RUN_DATE)
    print("Benchmark             :", BENCHMARK_LABEL)
    print("Universe filter        :", "Latest NSE CURRENT VOLUME >= 1 lakh (100,000 units); no later liquidity exclusion")
    print("Project folder        :", ROOT)
    print("="*72)

    # -----------------------------
    # Helpers
    # -----------------------------
    def clean_col(x):
        return re.sub(r"[^a-z0-9]", "", str(x).lower())

    def safe_num(x):
        try:
            return float(x)
        except Exception:
            return np.nan

    def retry_get(session, url, tries=3, timeout=30):
        last = None
        for i in range(tries):
            try:
                r = session.get(url, timeout=timeout)
                r.raise_for_status()
                return r
            except Exception as e:
                last = e
                time.sleep(1.5 * (i + 1))
        raise last

    def nse_session():
        s = requests.Session()
        s.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                          "AppleWebKit/537.36 (KHTML, like Gecko) "
                          "Chrome/131.0 Safari/537.36",
            "Accept-Language": "en-US,en;q=0.9",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,"
                      "image/avif,image/webp,*/*;q=0.8",
            "Referer": "https://www.nseindia.com/"
        })
        try:
            s.get("https://www.nseindia.com", timeout=15)
        except Exception:
            pass
        return s

    # -----------------------------
    # STEP 2 — Download official NSE ETF master
    # -----------------------------
    def download_nse_etf_master():
        page = "https://www.nseindia.com/static/market-data/securities-available-for-trading"
        s = nse_session()
        print("\n[1/10] Downloading official NSE ETF master list ...")
        html = retry_get(s, page).text
        soup = BeautifulSoup(html, "html.parser")

        candidates = []
        for a in soup.find_all("a", href=True):
            href = a.get("href", "")
            txt = " ".join(a.stripped_strings)
            combo = (txt + " " + href).lower()
            if ".csv" in combo and "etf" in combo:
                candidates.append(href)

        from urllib.parse import urljoin
        urls = [urljoin(page, h) for h in candidates]

        # Additional likely paths as fallbacks
        fallbacks = [
            "https://nsearchives.nseindia.com/content/equities/eq_etfseclist.csv",
            "https://archives.nseindia.com/content/equities/eq_etfseclist.csv"
        ]
        urls.extend(fallbacks)

        seen = set()
        last_err = None
        for url in urls:
            if url in seen:
                continue
            seen.add(url)
            try:
                r = retry_get(s, url, tries=2, timeout=30)
                raw = r.content
                # try CSV
                for enc in ["utf-8-sig", "utf-8", "latin1"]:
                    try:
                        df = pd.read_csv(BytesIO(raw), encoding=enc)
                        if len(df) > 20 and len(df.columns) >= 4:
                            df.to_csv(DATA_DIR / "NSE_ETF_MASTER_RAW.csv", index=False)
                            print("   Master source:", url)
                            print("   ETFs found   :", len(df))
                            return df, url
                    except Exception:
                        continue
            except Exception as e:
                last_err = e

        raise RuntimeError(
            "Could not download NSE ETF master automatically. "
            f"Last error: {last_err}"
        )

    master_raw, MASTER_SOURCE = download_nse_etf_master()

    # -----------------------------
    # STEP 2A — Download CURRENT NSE ETF market snapshot
    # -----------------------------
    # This is the authoritative source for CURRENT/LATEST-session fields:
    # LTP, previous close, % change, current volume and current traded value.
    # Historical Yahoo data is used only to build the past time series.
    def _pick(d, keys, default=np.nan):
        for k in keys:
            if k in d and d[k] not in (None, "", "-"):
                return d[k]
        return default

    def _to_float(x):
        if isinstance(x, (int, float, np.number)):
            return float(x)
        if x is None:
            return np.nan
        s = str(x).replace(",", "").replace("₹", "").strip()
        if s in ("", "-", "NA", "N/A", "None"):
            return np.nan
        try:
            return float(s)
        except Exception:
            return np.nan

    def download_nse_etf_live_snapshot():
        print("\n[1A/10] Downloading CURRENT NSE ETF market snapshot ...")
        s = nse_session()

        # NSE's ETF market page is the official live market-watch source.
        page_url = "https://www.nseindia.com/market-data/exchange-traded-funds-etf"
        try:
            s.get(page_url, timeout=20)
        except Exception:
            pass

        # NSE has historically served the dynamic table through its JSON API.
        # Try the official endpoint and a small set of compatible variants.
        api_urls = [
            "https://www.nseindia.com/api/etf",
            "https://www.nseindia.com/api/etf?type=all",
            "https://www.nseindia.com/api/etf?category=all"
        ]

        last_error = None
        payload = None
        source_url = None

        for url in api_urls:
            try:
                r = s.get(
                    url,
                    headers={
                        "Accept": "application/json,text/plain,*/*",
                        "Referer": page_url,
                        "User-Agent": s.headers.get("User-Agent", "Mozilla/5.0")
                    },
                    timeout=25
                )
                if r.status_code == 200:
                    j = r.json()
                    if j:
                        payload = j
                        source_url = url
                        break
            except Exception as e:
                last_error = e

        if payload is None:
            raise RuntimeError(
                "NSE CURRENT ETF snapshot could not be downloaded. "
                "Scanner stopped intentionally so it will NOT silently use stale Yahoo data "
                f"as the latest market snapshot. Last error: {last_error}"
            )

        # Locate the row list in common NSE response shapes.
        rows = None
        if isinstance(payload, list):
            rows = payload
        elif isinstance(payload, dict):
            for key in ["data", "marketData", "etf", "records"]:
                val = payload.get(key)
                if isinstance(val, list):
                    rows = val
                    break
                if isinstance(val, dict):
                    for subkey in ["data", "rows", "records"]:
                        if isinstance(val.get(subkey), list):
                            rows = val[subkey]
                            break
                if rows is not None:
                    break

        if not rows:
            raise RuntimeError(
                "NSE ETF snapshot downloaded, but no ETF rows were found in the response. "
                "The NSE response format may have changed."
            )

        out = []
        for d in rows:
            if not isinstance(d, dict):
                continue

            symbol = str(_pick(d, ["symbol", "Symbol", "symbolName"], "")).strip().upper()
            if not symbol:
                continue

            ltp = _to_float(_pick(d, ["ltP", "ltp", "lastPrice", "last", "LTP"]))
            prev = _to_float(_pick(d, ["prevClose", "previousClose", "previousclose", "prevclose"]))
            day_open = _to_float(_pick(d, ["open", "openPrice", "dayOpen"]))
            day_high = _to_float(_pick(d, ["high", "dayHigh", "highPrice"]))
            day_low = _to_float(_pick(d, ["low", "dayLow", "lowPrice"]))
            chg = _to_float(_pick(d, ["change", "netChange", "chg"]))
            pchg = _to_float(_pick(d, ["pChange", "percentChange", "perChange", "changePercent", "pchange"]))
            vol = _to_float(_pick(d, ["qty", "volume", "totalTradedVolume", "tradedVolume", "totalTradedQty"]))
            val_cr = _to_float(_pick(d, ["value", "tradedValue", "totalTradedValue", "turnover", "totalTradedValueCr"]))

            # Reconstruct fields if NSE omits one but provides the others.
            if pd.isna(chg) and pd.notna(ltp) and pd.notna(prev):
                chg = ltp - prev
            if pd.isna(pchg) and pd.notna(ltp) and pd.notna(prev) and prev != 0:
                pchg = (ltp / prev - 1.0) * 100.0

            # NSE page displays VALUE in ₹ crores. Some JSON variants may return rupees.
            # If value is obviously rupee-scale, convert to crores.
            if pd.notna(val_cr) and val_cr > 1_000_000:
                val_cr = val_cr / 1e7

            # If value is unavailable, calculate current traded value from NSE LTP x NSE volume.
            if pd.isna(val_cr) and pd.notna(ltp) and pd.notna(vol):
                val_cr = (ltp * vol) / 1e7

            out.append({
                "Symbol": symbol,
                "NSE_LTP": ltp,
                "NSE_PrevClose": prev,
                "NSE_Open": day_open,
                "NSE_High": day_high,
                "NSE_Low": day_low,
                "NSE_Change": chg,
                "NSE_ChangePct": pchg,
                "NSE_CurrentVolume": vol,
                "NSE_CurrentValueCr": val_cr
            })

        live = pd.DataFrame(out).drop_duplicates("Symbol")
        if live.empty:
            raise RuntimeError("NSE live ETF snapshot contained no usable symbol rows.")

        # Timestamp every run in Indian Standard Time.
        ts_ist = datetime.now(ZoneInfo("Asia/Kolkata"))
        timestamp_text = ts_ist.strftime("%d-%b-%Y %I:%M:%S %p IST")
        live["NSE_DownloadTimestamp"] = timestamp_text

        # Preserve any NSE response-level timestamp if supplied.
        nse_asof = ""
        if isinstance(payload, dict):
            for k in ["timestamp", "timeStamp", "lastUpdateTime", "dateTime"]:
                if payload.get(k):
                    nse_asof = str(payload.get(k))
                    break
        live["NSE_AsOf"] = nse_asof

        live.to_csv(DATA_DIR / "NSE_ETF_LIVE_SNAPSHOT.csv", index=False)

        print("   NSE live source       :", source_url)
        print("   NSE ETF rows          :", len(live))
        print("   Download timestamp    :", timestamp_text)
        if nse_asof:
            print("   NSE reported as-of    :", nse_asof)

        return live, timestamp_text, source_url

    nse_live, NSE_DOWNLOAD_TIMESTAMP, NSE_LIVE_SOURCE = download_nse_etf_live_snapshot()

    # Identify columns flexibly
    norm_map = {clean_col(c): c for c in master_raw.columns}

    def find_col(options):
        for opt in options:
            key = clean_col(opt)
            if key in norm_map:
                return norm_map[key]
        # contains fallback
        for c in master_raw.columns:
            cc = clean_col(c)
            for opt in options:
                if clean_col(opt) in cc or cc in clean_col(opt):
                    return c
        return None

    COL_SYMBOL = find_col(["Symbol"])
    COL_UNDERLYING = find_col(["Underlying"])
    COL_SECURITY = find_col(["Security Name", "SecurityName"])
    COL_LISTING = find_col(["Date of Listing", "DateofListing"])
    COL_ISIN = find_col(["ISIN Number", "ISINNumber", "ISIN"])

    if not COL_SYMBOL:
        raise ValueError(f"Symbol column not found. NSE columns: {list(master_raw.columns)}")

    master = pd.DataFrame()
    master["Symbol"] = master_raw[COL_SYMBOL].astype(str).str.strip().str.upper()
    master["Underlying"] = master_raw[COL_UNDERLYING].astype(str).str.strip() if COL_UNDERLYING else ""
    master["SecurityName"] = master_raw[COL_SECURITY].astype(str).str.strip() if COL_SECURITY else ""
    master["DateOfListing"] = master_raw[COL_LISTING] if COL_LISTING else ""
    master["ISIN"] = master_raw[COL_ISIN].astype(str).str.strip() if COL_ISIN else ""
    master = master[master["Symbol"].ne("") & master["Symbol"].ne("NAN")].drop_duplicates("Symbol").reset_index(drop=True)

    # -----------------------------
    # STEP 3 — Automatic classification
    # -----------------------------
    def asset_class(symbol, underlying, security):
        # Use every available identity field so Gold/Silver/Debt/Liquid ETFs
        # cannot accidentally enter the Equity trade universe.
        t = f"{symbol} {underlying} {security}".upper()

        if re.search(r"SILVER|SILV", t):
            return "SILVER"
        if re.search(r"GOLD", t):
            return "GOLD"
        if re.search(r"LIQUID|OVERNIGHT|1D RATE|LIQUID RATE|CASH", t):
            return "LIQUID"
        if re.search(r"G[- ]?SEC|GSEC|GILT|SDL|BHARAT BOND|BOND|DEBT|GOVERNMENT SECUR", t):
            return "DEBT"
        if re.search(r"NASDAQ|HANG SENG|NYSE|FANG|S&P 500|S&P500|GLOBAL|OVERSEAS", t):
            return "INTERNATIONAL"
        if "MSCI" in t and "INDIA" not in t:
            return "INTERNATIONAL"
        return "EQUITY"

    THEME_RULES = [
        ("PSU BANK", r"PSU\s*BANK"),
        ("PRIVATE BANK", r"PRIVATE\s*BANK|PVT\s*BANK"),
        ("BANK", r"\bBANK\b|BANKING"),
        ("FINANCIAL SERVICES", r"FINANCIAL\s*SERV|FINANCE"),
        ("IT", r"\bIT\b|INFORMATION TECHNOLOGY"),
        ("HEALTHCARE", r"HEALTH|PHARMA|HOSPITAL"),
        ("AUTO / EV", r"\bAUTO\b|AUTOMOTIVE|\bEV\b|ELECTRIC VEH"),
        ("FMCG", r"\bFMCG\b"),
        ("CONSUMPTION", r"CONSUM"),
        ("METAL", r"\bMETAL"),
        ("OIL & GAS / ENERGY", r"OIL|GAS|\bENERGY\b"),
        ("INFRASTRUCTURE", r"INFRA|LOGISTIC"),
        ("REALTY", r"REALTY|REAL ESTATE"),
        ("DEFENCE", r"DEFEN[CS]E"),
        ("RAILWAYS", r"RAIL"),
        ("CAPITAL MARKET", r"CAPITAL MARKET|CAPITAL MARKETS"),
        ("CHEMICAL", r"CHEM"),
        ("MANUFACTURING", r"MANUFACTUR"),
        ("DIGITAL / INTERNET", r"DIGITAL|INTERNET"),
        ("TOURISM", r"TOURISM|TRAVEL"),
        ("SERVICES", r"SERVICES SECTOR"),
        ("COMMODITIES EQUITY", r"NIFTY COMMODIT"),
        ("ALPHA", r"\bALPHA\b"),
        ("MOMENTUM", r"MOMENTUM"),
        ("QUALITY", r"QUALITY"),
        ("VALUE", r"\bVALUE\b"),
        ("LOW VOLATILITY", r"LOW VOL|LOWVOL|VOLATILITY"),
        ("DIVIDEND", r"DIVIDEND"),
        ("EQUAL WEIGHT", r"EQUAL WEIGHT|EQUALWEIGHT"),
        ("ESG", r"\bESG\b"),
        ("MNC", r"\bMNC\b"),
        ("IPO", r"\bIPO\b"),
        ("SMALLCAP", r"SMALL\s*CAP|SMALLCAP"),
        ("MIDCAP", r"MID\s*CAP|MIDCAP"),
        ("NEXT 50", r"NEXT\s*50|JUNIOR"),
        ("NIFTY 500", r"NIFTY\s*500"),
        ("NIFTY 200", r"NIFTY\s*200"),
        ("NIFTY 100", r"NIFTY\s*100"),
        ("NIFTY 50", r"NIFTY\s*50|NIFTY50"),
        ("TOTAL / MULTICAP", r"TOTAL MARKET|MULTICAP|LARGE.?MID"),
        ("SENSEX / BSE BROAD", r"SENSEX|BSE\s*500|BHARAT\s*22"),
    ]

    def theme_name(underlying, security, aclass):
        t = f"{underlying} {security}".upper()

        if aclass in ["GOLD", "SILVER", "LIQUID", "DEBT"]:
            return aclass

        # International equity stays inside GROUP 1 (Equity) but keeps its
        # own underlying index/theme rather than being collapsed into one bucket.
        if aclass == "INTERNATIONAL":
            x = str(underlying).upper()
            x = re.sub(r"TOTAL RETURN INDEX|TRI|INDEX|\(TRI\)", "", x)
            x = re.sub(r"\s+", " ", x).strip()
            return ("INTL - " + x[:38]) if x else "INTL EQUITY"

        for name, pat in THEME_RULES:
            if re.search(pat, t):
                return name

        # Clean NSE underlying as fallback
        x = str(underlying).upper()
        x = re.sub(r"TOTAL RETURN INDEX|TRI|INDEX|\(TRI\)", "", x)
        x = re.sub(r"\s+", " ", x).strip()
        return x[:45] if x else "OTHER EQUITY"

    master["AssetClass"] = [
        asset_class(sym, u, s)
        for sym, u, s in zip(master["Symbol"], master["Underlying"], master["SecurityName"])
    ]
    master["Theme"] = [theme_name(u, s, a) for u, s, a in zip(master["Underlying"], master["SecurityName"], master["AssetClass"])]
    master["YFTicker"] = master["Symbol"] + ".NS"

    # Merge CURRENT NSE snapshot onto the ETF master.
    master = master.merge(nse_live, on="Symbol", how="left")

    # ------------------------------------------------------------
    # PRIMARY NSE VOLUME UNIVERSE FILTER
    # ------------------------------------------------------------
    # User-selected rule:
    #   Keep an ETF for ALL further analysis only when the latest NSE
    #   snapshot shows CURRENT VOLUME >= 1 lakh (100,000 units/shares).
    #
    # IMPORTANT:
    # - This is a QUANTITY/VOLUME filter, NOT Rs 1 lakh turnover and NOT Rs 1 crore.
    # - It is applied ONCE at the beginning of the run.
    # - After an ETF passes this filter, no second liquidity threshold removes it.
    # - Historical/current volume is then used only for ranking/analysis.
    NSE_MIN_CURRENT_VOLUME = 100_000

    master_all = master.copy()

    master_excluded_low_volume = master_all[
        master_all["NSE_CurrentVolume"].fillna(0) < NSE_MIN_CURRENT_VOLUME
    ].copy()
    master_excluded_low_volume["ExclusionReason"] = "Latest NSE Volume < 1 lakh"

    master = master_all[
        master_all["NSE_CurrentVolume"].fillna(0) >= NSE_MIN_CURRENT_VOLUME
    ].copy().reset_index(drop=True)

    # Save both selected universe and exclusions for audit.
    master.to_csv(DATA_DIR / "ETF_MASTER_CURRENT.csv", index=False)
    master_excluded_low_volume.to_csv(
        DATA_DIR / "ETF_EXCLUDED_NSE_VOLUME_BELOW_1LAKH.csv", index=False
    )

    print("\n   NSE CURRENT-VOLUME UNIVERSE FILTER")
    print(f"   Full NSE ETF master      : {len(master_all)}")
    print(f"   Volume threshold         : {NSE_MIN_CURRENT_VOLUME:,} units (1 lakh)")
    print(f"   Selected for analysis    : {len(master)}")
    print(f"   Ignored below threshold  : {len(master_excluded_low_volume)}")

    print("\n   Selected asset-class counts")
    print(master["AssetClass"].value_counts().to_string())

    if master.empty:
        raise RuntimeError(
            "No ETF passed the latest NSE current-volume >= 1 lakh filter."
        )

    # -----------------------------
    # STEP 4 — Download historical price and volume
    # -----------------------------
    print("\n[2/10] Downloading ETF market history for NSE-volume-selected universe ...")
    all_tickers = master["YFTicker"].tolist()
    all_tickers_with_bench = list(dict.fromkeys(all_tickers + [BENCHMARK_YF]))

    def yf_batch_download(tickers, period_days=LOOKBACK_DAYS):
        end = pd.Timestamp.today().normalize() + pd.Timedelta(days=1)
        start = end - pd.Timedelta(days=int(period_days * 1.65))  # calendar cushion
        price_parts = []
        vol_parts = []
        high_parts = []
        low_parts = []
        failed = []

        for i in range(0, len(tickers), DOWNLOAD_BATCH):
            batch = tickers[i:i+DOWNLOAD_BATCH]
            print(f"   Batch {i//DOWNLOAD_BATCH + 1:02d} | {len(batch):02d} tickers")
            ok = False
            data = None
            for attempt in range(2):
                try:
                    data = yf.download(
                        tickers=batch,
                        start=start.strftime("%Y-%m-%d"),
                        end=end.strftime("%Y-%m-%d"),
                        auto_adjust=False,
                        progress=False,
                        group_by="column",
                        threads=True,
                        timeout=25
                    )
                    if data is not None and len(data) > 0:
                        ok = True
                        break
                except Exception:
                    time.sleep(2)

            if not ok:
                failed.extend(batch)
                continue

            try:
                if isinstance(data.columns, pd.MultiIndex):
                    if "Close" in data.columns.get_level_values(0):
                        c0 = data["Close"].copy()
                    elif "Adj Close" in data.columns.get_level_values(0):
                        c0 = data["Adj Close"].copy()
                    else:
                        c0 = pd.DataFrame(index=data.index)

                    v0 = data["Volume"].copy() if "Volume" in data.columns.get_level_values(0) else pd.DataFrame(index=data.index)
                    h0 = data["High"].copy() if "High" in data.columns.get_level_values(0) else pd.DataFrame(index=data.index)
                    l0 = data["Low"].copy() if "Low" in data.columns.get_level_values(0) else pd.DataFrame(index=data.index)
                else:
                    tk = batch[0]
                    close_col = "Close" if "Close" in data.columns else "Adj Close"
                    c0 = data[[close_col]].rename(columns={close_col: tk})
                    v0 = data[["Volume"]].rename(columns={"Volume": tk}) if "Volume" in data.columns else pd.DataFrame(index=data.index, columns=[tk])
                    h0 = data[["High"]].rename(columns={"High": tk}) if "High" in data.columns else pd.DataFrame(index=data.index, columns=[tk])
                    l0 = data[["Low"]].rename(columns={"Low": tk}) if "Low" in data.columns else pd.DataFrame(index=data.index, columns=[tk])

                price_parts.append(c0)
                vol_parts.append(v0)
                high_parts.append(h0)
                low_parts.append(l0)
            except Exception:
                failed.extend(batch)

        prices = pd.concat(price_parts, axis=1) if price_parts else pd.DataFrame()
        volumes = pd.concat(vol_parts, axis=1) if vol_parts else pd.DataFrame()
        highs = pd.concat(high_parts, axis=1) if high_parts else pd.DataFrame()
        lows = pd.concat(low_parts, axis=1) if low_parts else pd.DataFrame()

        for df0 in [prices, volumes, highs, lows]:
            if not df0.empty:
                df0.drop(columns=df0.columns[df0.columns.duplicated()], inplace=True)

        return (
            prices.sort_index(),
            volumes.sort_index(),
            highs.sort_index(),
            lows.sort_index(),
            sorted(set(failed))
        )

    prices, volumes, highs, lows, failed_batches = yf_batch_download(all_tickers_with_bench)

    if BENCHMARK_YF not in prices.columns or prices[BENCHMARK_YF].dropna().empty:
        print("   Benchmark missing from batch. Retrying NIFTY 50 separately ...")
        b = yf.download(BENCHMARK_YF, period="3y", auto_adjust=False, progress=False)
        if isinstance(b.columns, pd.MultiIndex):
            bclose = b["Close"].iloc[:,0]
        else:
            bclose = b["Close"]
        prices[BENCHMARK_YF] = bclose

    # Keep enough observations
    prices = prices.sort_index().tail(LOOKBACK_DAYS)
    volumes = volumes.reindex(prices.index).sort_index().tail(LOOKBACK_DAYS)
    highs = highs.reindex(prices.index).sort_index().tail(LOOKBACK_DAYS)
    lows = lows.reindex(prices.index).sort_index().tail(LOOKBACK_DAYS)

    # ------------------------------------------------------------
    # CURRENT NSE PRICE / HIGH / LOW OVERRIDE
    # ------------------------------------------------------------
    # Current NSE data is injected into the newest observation so the positive
    # swing ranking reflects the market at the exact time the scanner is run.
    current_price_date = pd.Timestamp.now(tz="Asia/Kolkata").tz_localize(None).normalize()

    for _, mr in master.iterrows():
        tk = mr["YFTicker"]
        nse_ltp = mr.get("NSE_LTP", np.nan)
        nse_high = mr.get("NSE_High", np.nan)
        nse_low = mr.get("NSE_Low", np.nan)

        if pd.notna(nse_ltp) and nse_ltp > 0:
            if tk not in prices.columns:
                prices[tk] = np.nan
            prices.loc[current_price_date, tk] = float(nse_ltp)

        if pd.notna(nse_high) and nse_high > 0:
            if tk not in highs.columns:
                highs[tk] = np.nan
            highs.loc[current_price_date, tk] = float(nse_high)

        if pd.notna(nse_low) and nse_low > 0:
            if tk not in lows.columns:
                lows[tk] = np.nan
            lows.loc[current_price_date, tk] = float(nse_low)

    prices = prices.sort_index()
    highs = highs.sort_index()
    lows = lows.sort_index()

    prices.to_csv(DATA_DIR / "PRICE_HISTORY_WITH_NSE_CURRENT.csv")
    volumes.to_csv(DATA_DIR / "VOLUME_HISTORY_COMPLETED_DAYS.csv")

    # -----------------------------
    # STEP 5 — Calculate ETF analytics
    # -----------------------------
    print("\n[3/10] Calculating liquidity, returns, moving averages and technical data ...")

    def last_valid(s):
        s = s.dropna()
        return s.iloc[-1] if len(s) else np.nan

    def ret_n(s, n):
        s = s.dropna()
        if len(s) <= n:
            return np.nan
        return (s.iloc[-1] / s.iloc[-1-n] - 1.0) * 100

    def dma(s, n):
        s = s.dropna()
        return s.tail(n).mean() if len(s) >= max(5, n//2) else np.nan

    def atr_from_history(close_s, high_s, low_s, n=20):
        df0 = pd.concat(
            [
                close_s.rename("Close"),
                high_s.rename("High"),
                low_s.rename("Low")
            ],
            axis=1
        ).dropna(subset=["Close"])
        if len(df0) < max(10, n):
            return np.nan
        prev_close = df0["Close"].shift(1)
        tr = pd.concat(
            [
                (df0["High"] - df0["Low"]).abs(),
                (df0["High"] - prev_close).abs(),
                (df0["Low"] - prev_close).abs()
            ],
            axis=1
        ).max(axis=1, skipna=True)
        tr = tr.replace([np.inf, -np.inf], np.nan).dropna()
        return tr.tail(n).mean() if len(tr) >= max(10, n//2) else np.nan

    rows = []
    for _, r in master.iterrows():
        tk = r["YFTicker"]
        if tk not in prices.columns:
            continue

        p = prices[tk].dropna()
        if p.empty:
            continue

        v = volumes[tk].reindex(p.index).fillna(0) if tk in volumes.columns else pd.Series(0, index=p.index)
        h = highs[tk].reindex(p.index) if tk in highs.columns else pd.Series(np.nan, index=p.index)
        l = lows[tk].reindex(p.index) if tk in lows.columns else pd.Series(np.nan, index=p.index)

        # CURRENT fields come from NSE snapshot, not Yahoo.
        nse_row = r
        ltp = nse_row.get("NSE_LTP", np.nan)
        if pd.isna(ltp):
            ltp = p.iloc[-1]

        nse_prev_close = nse_row.get("NSE_PrevClose", np.nan)
        nse_open = nse_row.get("NSE_Open", np.nan)
        nse_high = nse_row.get("NSE_High", np.nan)
        nse_low = nse_row.get("NSE_Low", np.nan)
        nse_change = nse_row.get("NSE_Change", np.nan)
        nse_change_pct = nse_row.get("NSE_ChangePct", np.nan)
        day_vol = nse_row.get("NSE_CurrentVolume", np.nan)
        nse_current_value_cr = nse_row.get("NSE_CurrentValueCr", np.nan)

        # Turnover calculations in crores
        # USER RULE:
        #   Today Turnover       = Current LTP x Today's Volume / 1 crore
        #   30D Liquidity Value  = Current LTP x 30-day Average Volume / 1 crore
        avg5_volume = v.tail(5).mean() if len(v) else np.nan
        avg10_volume = v.tail(10).mean() if len(v) else np.nan
        avg20_volume = v.tail(20).mean() if len(v) else np.nan
        avg30_volume = v.tail(30).mean() if len(v) else np.nan

        # Positive-side volume pace uses only PRIOR COMPLETED sessions.
        if tk in volumes.columns:
            v_completed = volumes[tk].dropna()
            v_completed = v_completed[v_completed.index < current_price_date]
            prior30_avg_volume = v_completed.tail(30).mean() if len(v_completed) else np.nan
        else:
            prior30_avg_volume = np.nan

        # Current-LTP liquidity estimates requested by user
        ltp_x_avg5vol_cr = (ltp * avg5_volume) / 1e7 if pd.notna(avg5_volume) else np.nan
        ltp_x_avg10vol_cr = (ltp * avg10_volume) / 1e7 if pd.notna(avg10_volume) else np.nan
        ltp_x_avg20vol_cr = (ltp * avg20_volume) / 1e7 if pd.notna(avg20_volume) else np.nan
        avg30_turnover = (ltp * avg30_volume) / 1e7 if pd.notna(avg30_volume) else np.nan

        # True historical ADTV (average of daily Close x Volume)
        daily_turnover_cr = (p * v) / 1e7
        adtv5 = daily_turnover_cr.tail(5).mean() if len(daily_turnover_cr) else np.nan
        adtv10 = daily_turnover_cr.tail(10).mean() if len(daily_turnover_cr) else np.nan
        adtv20 = daily_turnover_cr.tail(20).mean() if len(daily_turnover_cr) else np.nan
        adtv30 = daily_turnover_cr.tail(30).mean() if len(daily_turnover_cr) else np.nan

        today_turnover = nse_current_value_cr if pd.notna(nse_current_value_cr) else ((ltp * day_vol) / 1e7 if pd.notna(day_vol) else np.nan)
        vol_multiple = day_vol / avg30_volume if pd.notna(avg30_volume) and avg30_volume > 0 else np.nan

        high52 = p.tail(252).max() if len(p) else np.nan
        low52 = p.tail(252).min() if len(p) else np.nan
        pct_from_high = (ltp / high52 - 1) * 100 if pd.notna(high52) and high52 else np.nan
        pct_above_low = (ltp / low52 - 1) * 100 if pd.notna(low52) and low52 else np.nan

        daily_ret = nse_change_pct if pd.notna(nse_change_pct) else ret_n(p, 1)
        annualized_vol = p.pct_change().tail(60).std() * np.sqrt(252) * 100

        dma20_val = dma(p, 20)
        dma50_val = dma(p, 50)
        dma100_val = dma(p, 100)
        dma200_val = dma(p, 200)
        # ATR uses completed historical sessions only so an early intraday
        # partial high-low range cannot make the ETF look artificially unstretched.
        completed_mask = p.index < current_price_date
        atr20 = atr_from_history(
            p.loc[completed_mask],
            h.loc[completed_mask],
            l.loc[completed_mask],
            20
        )
        entry_stretch_atr = (
            (ltp - dma20_val) / atr20
            if pd.notna(ltp) and pd.notna(dma20_val) and pd.notna(atr20) and atr20 > 0
            else np.nan
        )

        rows.append({
            "Symbol": r["Symbol"],
            "Underlying": r["Underlying"],
            "SecurityName": r["SecurityName"],
            "DateOfListing": r["DateOfListing"],
            "ISIN": r["ISIN"],
            "AssetClass": r["AssetClass"],
            "Theme": r["Theme"],
            "YFTicker": tk,
            "NSE_DownloadTimestamp": nse_row.get("NSE_DownloadTimestamp", NSE_DOWNLOAD_TIMESTAMP),
            "NSE_AsOf": nse_row.get("NSE_AsOf", ""),
            "NSE_PrevClose": nse_prev_close,
            "NSE_Open": nse_open,
            "NSE_High": nse_high,
            "NSE_Low": nse_low,
            "NSE_Change": nse_change,
            "NSE_ChangePct": nse_change_pct,
            "LTP": ltp,
            "DayVolume": day_vol,
            "TodayTurnoverCr": today_turnover,
            "Avg5Volume": avg5_volume,
            "Avg10Volume": avg10_volume,
            "Avg20Volume": avg20_volume,
            "Avg30Volume": avg30_volume,
            "Prior30AvgVolume": prior30_avg_volume,
            "LTPxAvg5VolCr": ltp_x_avg5vol_cr,
            "LTPxAvg10VolCr": ltp_x_avg10vol_cr,
            "LTPxAvg20VolCr": ltp_x_avg20vol_cr,
            "Avg30TurnoverCr": avg30_turnover,
            "ADTV5Cr": adtv5,
            "ADTV10Cr": adtv10,
            "ADTV20Cr": adtv20,
            "ADTV30Cr": adtv30,
            "VolumeMultiple": vol_multiple,
            "DayReturnPct": daily_ret,
            "Week1ReturnPct": ret_n(p, 5),
            "Week2ReturnPct": ret_n(p, 10),
            "Month1ReturnPct": ret_n(p, 21),
            "Month3ReturnPct": ret_n(p, 63),
            "Month6ReturnPct": ret_n(p, 126),
            "Year1ReturnPct": ret_n(p, 252),
            "High52W": high52,
            "Low52W": low52,
            "PctFrom52WHigh": pct_from_high,
            "PctAbove52WLow": pct_above_low,
            "DMA20": dma20_val,
            "DMA50": dma50_val,
            "DMA100": dma100_val,
            "DMA200": dma200_val,
            "ATR20": atr20,
            "EntryStretchATR": entry_stretch_atr,
            "AnnVol60DPct": annualized_vol,
            "HistoryDays": int(p.notna().sum())
        })

    etf = pd.DataFrame(rows)

    if etf.empty:
        raise RuntimeError("No ETF price data could be downloaded.")

    # ------------------------------------------------------------
    # LIQUIDITY / UNIVERSE STATUS
    # ------------------------------------------------------------
    # No secondary liquidity gate is applied here.
    # Every ETF in `etf` already passed the ONLY eligibility condition:
    # latest NSE current volume >= 1 lakh units/shares.
    #
    # Historical ADTV, turnover and volume multiples remain available as
    # analytical/ranking fields only. They DO NOT remove an ETF.
    etf["LiquidityStatus"] = "NSE VOLUME >= 1 LAKH"
    etf["LiquidityTier"] = "SELECTED NSE VOLUME UNIVERSE"
    etf["LiquidEligible"] = True

    eligible = etf.copy()

    # Liquidity acceleration
    etf["LiquidityAcceleration"] = np.where(
        etf["ADTV20Cr"].fillna(0) > 0,
        etf["ADTV5Cr"] / etf["ADTV20Cr"],
        np.nan
    )
    eligible["LiquidityAcceleration"] = etf.loc[eligible.index, "LiquidityAcceleration"]

    print("   ETFs with price data    :", len(etf))
    print("   ETFs in analysis universe:", len(eligible))
    print("   Universe rule           : Latest NSE Current Volume >= 1 lakh")
    print("   Failed batch tickers    :", len(failed_batches))

    # -----------------------------
    # STEP 6 — RRG engine
    # -----------------------------
    print("\n[4/10] Building RRG calculations ...")

    def rrg_series(asset_price, bench_price):
        df = pd.concat([asset_price.rename("asset"), bench_price.rename("bench")], axis=1).dropna()
        if len(df) < MIN_RRG_OBS:
            return pd.DataFrame(index=df.index, columns=["RS_Ratio", "RS_Momentum"])

        rs = df["asset"] / df["bench"]

        # Relative-strength ratio centered near 100
        rs_base = rs.rolling(RRG_RS_WINDOW).mean()
        ratio = 100.0 * rs / rs_base

        # Smooth to reduce noise
        ratio = ratio.ewm(span=RRG_SMOOTH, adjust=False).mean()

        # Momentum of the RS-Ratio, centered near 100
        momentum = 100.0 * ratio / ratio.shift(RRG_MOM_LAG)
        momentum = momentum.ewm(span=RRG_SMOOTH, adjust=False).mean()

        return pd.DataFrame({"RS_Ratio": ratio, "RS_Momentum": momentum}).dropna()

    def quadrant(rr, rm):
        if pd.isna(rr) or pd.isna(rm):
            return "NO DATA"
        if rr >= 100 and rm >= 100:
            return "LEADING"
        if rr < 100 and rm >= 100:
            return "IMPROVING"
        if rr >= 100 and rm < 100:
            return "WEAKENING"
        return "LAGGING"

    bench = prices[BENCHMARK_YF].dropna()

    # ETF-level RRG
    rrg_etf_rows = []
    rrg_history = {}
    for _, r in eligible.iterrows():
        tk = r["YFTicker"]
        if tk not in prices.columns:
            continue
        rr = rrg_series(prices[tk], bench)
        if rr.empty:
            continue
        rrg_history[r["Symbol"]] = rr

        latest = rr.iloc[-1]
        prev = rr.iloc[-min(RRG_TRAIL_DAYS+1, len(rr))]
        rrg_etf_rows.append({
            "Symbol": r["Symbol"],
            "Theme": r["Theme"],
            "AssetClass": r["AssetClass"],
            "RS_Ratio": latest["RS_Ratio"],
            "RS_Momentum": latest["RS_Momentum"],
            "Quadrant": quadrant(latest["RS_Ratio"], latest["RS_Momentum"]),
            "RS_Ratio_5D_Change": latest["RS_Ratio"] - prev["RS_Ratio"],
            "RS_Momentum_5D_Change": latest["RS_Momentum"] - prev["RS_Momentum"],
        })

    rrg_etf = pd.DataFrame(rrg_etf_rows)

    # Theme / underlying-family proxy:
    # Use highest-liquidity eligible ETF inside each theme as the representative proxy.
    theme_reps = (
        eligible.sort_values(["Theme", "Avg30TurnoverCr"], ascending=[True, False])
                .groupby("Theme", as_index=False)
                .first()
    )

    theme_rows = []
    theme_rrg_history = {}
    for _, tr in theme_reps.iterrows():
        tk = tr["YFTicker"]
        theme = tr["Theme"]
        if tk not in prices.columns:
            continue
        rr = rrg_series(prices[tk], bench)
        if rr.empty:
            continue
        theme_rrg_history[theme] = rr
        latest = rr.iloc[-1]
        prev = rr.iloc[-min(RRG_TRAIL_DAYS+1, len(rr))]
        theme_rows.append({
            "Theme": theme,
            "AssetClass": tr["AssetClass"],
            "ProxyETF": tr["Symbol"],
            "ProxyAvg30TurnoverCr": tr["Avg30TurnoverCr"],
            "RS_Ratio": latest["RS_Ratio"],
            "RS_Momentum": latest["RS_Momentum"],
            "Quadrant": quadrant(latest["RS_Ratio"], latest["RS_Momentum"]),
            "RS_Ratio_5D_Change": latest["RS_Ratio"] - prev["RS_Ratio"],
            "RS_Momentum_5D_Change": latest["RS_Momentum"] - prev["RS_Momentum"],
        })

    theme_rrg = pd.DataFrame(theme_rows)

    # -----------------------------
    # STEP 7 — Data-driven ETF Trade Score
    # -----------------------------
    print("\n[5/10] Ranking ETFs inside the strongest themes ...")

    analysis = eligible.merge(
        rrg_etf[["Symbol","RS_Ratio","RS_Momentum","Quadrant","RS_Ratio_5D_Change","RS_Momentum_5D_Change"]],
        on="Symbol", how="left"
    )

    analysis = analysis.merge(
        theme_rrg[["Theme","ProxyETF","RS_Ratio","RS_Momentum","Quadrant","RS_Ratio_5D_Change","RS_Momentum_5D_Change"]]
            .rename(columns={
                "RS_Ratio":"Theme_RS_Ratio",
                "RS_Momentum":"Theme_RS_Momentum",
                "Quadrant":"Theme_Quadrant",
                "RS_Ratio_5D_Change":"Theme_RS_Ratio_5D_Change",
                "RS_Momentum_5D_Change":"Theme_RS_Momentum_5D_Change"
            }),
        on="Theme", how="left"
    )

    def pct_rank(s):
        return s.rank(pct=True).fillna(0.5) * 100

    # Scores: 0-100 components
    analysis["LiquidityScore"] = pct_rank(np.log1p(analysis["Avg30TurnoverCr"].clip(lower=0)))
    analysis["VolumeImpulseScore"] = pct_rank(analysis["VolumeMultiple"].clip(lower=0, upper=5))
    analysis["MomentumScore"] = (
        0.05*pct_rank(analysis["DayReturnPct"]) +   # deliberately low weight: today's losers can still be good dip setups
        0.25*pct_rank(analysis["Week1ReturnPct"]) +
        0.30*pct_rank(analysis["Month1ReturnPct"]) +
        0.40*pct_rank(analysis["Month3ReturnPct"])
    )

    analysis["TrendPoints"] = (
        (analysis["LTP"] > analysis["DMA20"]).astype(int) +
        (analysis["LTP"] > analysis["DMA50"]).astype(int) +
        (analysis["LTP"] > analysis["DMA100"]).astype(int) +
        (analysis["LTP"] > analysis["DMA200"]).astype(int)
    )
    analysis["TrendScore"] = analysis["TrendPoints"] / 4 * 100

    analysis["ETF_RRG_Score"] = np.select(
        [
            analysis["Quadrant"].eq("LEADING"),
            analysis["Quadrant"].eq("IMPROVING"),
            analysis["Quadrant"].eq("WEAKENING"),
            analysis["Quadrant"].eq("LAGGING"),
        ],
        [90, 75, 50, 25],
        default=40
    ) + analysis["RS_Momentum_5D_Change"].fillna(0).clip(-2,2)*2.5
    analysis["ETF_RRG_Score"] = analysis["ETF_RRG_Score"].clip(0,100)

    analysis["ThemeRRGScore"] = np.select(
        [
            analysis["Theme_Quadrant"].eq("LEADING"),
            analysis["Theme_Quadrant"].eq("IMPROVING"),
            analysis["Theme_Quadrant"].eq("WEAKENING"),
            analysis["Theme_Quadrant"].eq("LAGGING"),
        ],
        [95, 80, 45, 20],
        default=40
    ) + analysis["Theme_RS_Momentum_5D_Change"].fillna(0).clip(-2,2)*2.5
    analysis["ThemeRRGScore"] = analysis["ThemeRRGScore"].clip(0,100)

    # Near-high score: strongest when close to 52W high but not wildly extended above it
    analysis["NearHighScore"] = (100 + analysis["PctFrom52WHigh"].fillna(-50)*2).clip(0,100)

    # Volatility penalty only at extreme percentile
    analysis["VolPenalty"] = pct_rank(analysis["AnnVol60DPct"]).clip(0,100)

    analysis["TradeScore"] = (
        0.27*analysis["ThemeRRGScore"] +
        0.18*analysis["ETF_RRG_Score"] +
        0.16*analysis["TrendScore"] +
        0.14*analysis["MomentumScore"] +
        0.13*analysis["LiquidityScore"] +
        0.07*analysis["VolumeImpulseScore"] +
        0.05*analysis["NearHighScore"]
        - 0.04*analysis["VolPenalty"]
    ).clip(0,100)

    def signal_label(row):
        # Equity trading candidates only when theme is rotating well.
        if row["AssetClass"] in ["EQUITY","INTERNATIONAL"]:
            if row["Theme_Quadrant"] == "LEADING" and row["Quadrant"] in ["LEADING","IMPROVING"] and row["TradeScore"] >= 72:
                return "A+ TRADE CANDIDATE"
            if row["Theme_Quadrant"] in ["LEADING","IMPROVING"] and row["TradeScore"] >= 64:
                return "A CANDIDATE"
            if row["TradeScore"] >= 55:
                return "WATCH"
            return "AVOID / WAIT"
        else:
            if row["Quadrant"] in ["LEADING","IMPROVING"] and row["TradeScore"] >= 65:
                return "STRONG ASSET"
            if row["TradeScore"] >= 55:
                return "WATCH"
            return "WEAK / WAIT"

    analysis["Signal"] = analysis.apply(signal_label, axis=1)
    analysis["RankOverall"] = analysis["TradeScore"].rank(method="dense", ascending=False).astype(int)
    analysis["RankInTheme"] = analysis.groupby("Theme")["TradeScore"].rank(method="dense", ascending=False).astype(int)

    # ------------------------------------------------------------
    # LATEST NSE SESSION PRIORITY ENGINE
    # ------------------------------------------------------------
    # Latest NSE data is primary. RRG/trend is supporting context.

    analysis["CurrentVolumeMultiple30D"] = np.where(
        analysis["Avg30Volume"].fillna(0) > 0,
        analysis["DayVolume"] / analysis["Avg30Volume"],
        np.nan
    )
    analysis["CurrentVsAvg30VolumePct"] = (analysis["CurrentVolumeMultiple30D"] - 1.0) * 100.0

    analysis["RangePositionPct"] = np.where(
        (analysis["NSE_High"] - analysis["NSE_Low"]).fillna(0) > 0,
        (analysis["LTP"] - analysis["NSE_Low"]) /
        (analysis["NSE_High"] - analysis["NSE_Low"]) * 100.0,
        np.nan
    )
    analysis["RangePositionPct"] = analysis["RangePositionPct"].clip(0, 100)

    def volume_surge_flag(x):
        if pd.isna(x): return "NO DATA"
        if x >= 200: return "EXTREME VOLUME"
        if x >= 100: return "VERY HIGH VOLUME"
        if x >= 50: return "HIGH VOLUME"
        if x >= 25: return "ABOVE AVG"
        return "NORMAL"

    analysis["VolumeSurgeFlag"] = analysis["CurrentVsAvg30VolumePct"].apply(volume_surge_flag)

    try:
        _now_ist = datetime.now(ZoneInfo("Asia/Kolkata"))
        _mins = _now_ist.hour * 60 + _now_ist.minute
        _open_m, _close_m = 9*60+15, 15*60+30
        ANALYSIS_TIME_IST = _now_ist.strftime("%d-%b-%Y %I:%M:%S %p IST")
        CURRENT_SESSION_MODE = (
            "INTRADAY"
            if _now_ist.weekday() < 5 and _open_m <= _mins < _close_m
            else "COMPLETED SESSION"
        )
    except Exception:
        ANALYSIS_TIME_IST = datetime.now().strftime("%d-%b-%Y %I:%M:%S %p")
        CURRENT_SESSION_MODE = "COMPLETED SESSION"

    analysis["LatestDataMode"] = CURRENT_SESSION_MODE

    # Intraday volume must be compared with the amount of volume normally expected
    # by this time of day, not with a full day's average volume.
    if CURRENT_SESSION_MODE == "INTRADAY":
        MARKET_PROGRESS_FRACTION = max(
            0.08,
            min(1.0, (_mins - _open_m) / float(_close_m - _open_m))
        )
    else:
        MARKET_PROGRESS_FRACTION = 1.0

    analysis["VolumePaceVs30D"] = np.where(
        analysis["Prior30AvgVolume"].fillna(0) > 0,
        analysis["DayVolume"] /
        (analysis["Prior30AvgVolume"] * MARKET_PROGRESS_FRACTION),
        np.nan
    )
    analysis["MarketProgressPct"] = MARKET_PROGRESS_FRACTION * 100.0

    analysis["LatestThemeContext"] = np.select(
        [
            analysis["Theme_Quadrant"].eq("LEADING"),
            analysis["Theme_Quadrant"].eq("IMPROVING"),
            analysis["Theme_Quadrant"].eq("WEAKENING"),
            analysis["Theme_Quadrant"].eq("LAGGING"),
        ],
        [100, 82, 45, 18], default=40
    )
    analysis["LatestETFContext"] = np.select(
        [
            analysis["Quadrant"].eq("LEADING"),
            analysis["Quadrant"].eq("IMPROVING"),
            analysis["Quadrant"].eq("WEAKENING"),
            analysis["Quadrant"].eq("LAGGING"),
        ],
        [100, 85, 50, 20], default=40
    )

    analysis["LatestVolumeScore"] = (
        20 + analysis["CurrentVolumeMultiple30D"].fillna(0).clip(0, 3) * (80/3)
    ).clip(0, 100)

    # ---------- HEAVY BUYING / ACCUMULATION ----------
    analysis["BuyingPriceScore"] = analysis["NSE_ChangePct"].fillna(0).clip(0, 4) / 4 * 100
    analysis["BuyingRangeScore"] = analysis["RangePositionPct"].fillna(50)

    # 70% latest-session data, 30% structural context.
    analysis["BuyingPressureScore"] = (
        0.30*analysis["BuyingPriceScore"] +
        0.25*analysis["BuyingRangeScore"] +
        0.15*analysis["LatestVolumeScore"] +
        0.15*analysis["LatestThemeContext"] +
        0.10*analysis["LatestETFContext"] +
        0.05*analysis["TrendScore"]
    ).clip(0, 100)

    def buying_signal(row):
        if row["AssetClass"] not in ["EQUITY","INTERNATIONAL"]:
            return "N/A"
        if pd.isna(row["NSE_ChangePct"]) or row["NSE_ChangePct"] <= 0:
            return "NO BUYING SIGNAL"

        volx = row.get("CurrentVolumeMultiple30D", np.nan)
        rp = row.get("RangePositionPct", np.nan)
        score = row.get("BuyingPressureScore", np.nan)

        if (
            pd.notna(volx) and volx >= 1.75
            and pd.notna(rp) and rp >= 80
            and pd.notna(score) and score >= 75
        ):
            return "A+ HEAVY BUYING"
        if (
            pd.notna(volx) and volx >= 1.25
            and pd.notna(rp) and rp >= 70
            and pd.notna(score) and score >= 65
        ):
            return "STRONG BUYING"
        if (
            pd.notna(volx) and volx >= 1.00
            and pd.notna(rp) and rp >= 60
            and pd.notna(score) and score >= 55
        ):
            return "BUYING PRESSURE"
        if pd.notna(rp) and rp < 35:
            return "WEAK POSITIVE / FADING"
        return "POSITIVE BUY WATCH"

    analysis["BuyingSignal"] = analysis.apply(buying_signal, axis=1)

    # ---------- DIP / SELLOFF ----------
    analysis["DipMagnitudeScore"] = np.select(
        [
            (analysis["NSE_ChangePct"] <= -1.0) & (analysis["NSE_ChangePct"] > -1.75),
            (analysis["NSE_ChangePct"] <= -1.75) & (analysis["NSE_ChangePct"] > -2.5),
            (analysis["NSE_ChangePct"] <= -2.5) & (analysis["NSE_ChangePct"] > -3.5),
            analysis["NSE_ChangePct"] <= -3.5,
        ],
        [90, 82, 65, 40], default=0
    )

    def dip_volume_character(row):
        volx = row.get("CurrentVolumeMultiple30D", np.nan)
        rp = row.get("RangePositionPct", np.nan)
        if pd.isna(volx): return "VOLUME DATA N/A"
        if volx >= 2.0 and pd.notna(rp) and rp < 40: return "EXTREME SELLOFF"
        if volx >= 1.3 and pd.notna(rp) and rp < 45: return "HIGH VOLUME SELLOFF"
        if volx >= 1.3 and pd.notna(rp) and rp >= 65: return "HIGH VOLUME ABSORPTION"
        if volx <= 0.8: return "LIGHT VOLUME DIP"
        return "NORMAL VOLUME DIP"

    analysis["DipVolumeCharacter"] = analysis.apply(dip_volume_character, axis=1)

    def dip_volume_score(row):
        # For this swing/mean-reversion dip strategy, heavier volume during a
        # meaningful down move is POSITIVE opportunity evidence, not an automatic rejection.
        volx = row.get("CurrentVolumeMultiple30D", np.nan)
        if pd.isna(volx):
            return 35
        if volx >= 4.0: return 100
        if volx >= 3.0: return 95
        if volx >= 2.0: return 88
        if volx >= 1.3: return 78
        if volx >= 0.8: return 62
        return 40

    analysis["DipVolumeQualityScore"] = analysis.apply(dip_volume_score, axis=1)

    # Current turnover matters as an execution/participation confirmation.
    analysis["DipTurnoverScore"] = (
        analysis["TodayTurnoverCr"].fillna(0).clip(0, 2.0) / 2.0 * 100
    ).clip(0,100)

    # Aggressive dip-buy score:
    # 80% latest-session evidence, 20% structural context.
    # Largest weights go to FALL SIZE and CURRENT VOLUME, per the chosen strategy.
    analysis["DipBuyScore"] = (
        0.40*analysis["DipMagnitudeScore"] +
        0.30*analysis["DipVolumeQualityScore"] +
        0.10*analysis["DipTurnoverScore"] +
        0.10*analysis["LatestThemeContext"] +
        0.05*analysis["LatestETFContext"] +
        0.05*analysis["TrendScore"]
    ).clip(0,100)

    def dip_signal(row):
        if row["AssetClass"] not in ["EQUITY","INTERNATIONAL"]:
            return "N/A"
        if pd.isna(row["NSE_ChangePct"]) or row["NSE_ChangePct"] > -1.0:
            return "NO DIP"

        score = row.get("DipBuyScore", np.nan)
        volx = row.get("CurrentVolumeMultiple30D", np.nan)
        chg = row.get("NSE_ChangePct", np.nan)

        # Extreme selloff is deliberately treated as a higher-priority DEEP DIP
        # when volume confirms strong participation. Risk remains visible in DipType.
        if pd.notna(score) and score >= 80 and pd.notna(volx) and volx >= 2.0:
            return "A+ DEEP DIP BUY"
        if pd.notna(score) and score >= 68:
            return "STRONG DIP BUY"
        if pd.notna(score) and score >= 55:
            return "DIP BUY WATCH"
        return "LOW PRIORITY DIP"

    analysis["DipSignal"] = analysis.apply(dip_signal, axis=1)
    analysis["DipBucket"] = pd.cut(
        analysis["NSE_ChangePct"],
        bins=[-np.inf,-3.0,-2.0,-1.0,np.inf],
        labels=["Below -3%","-3% to -2%","-2% to -1%","Above -1%"],
        right=True
    )

    # "Best ETF" is highest-ranked liquid ETF inside each theme
    best_by_theme = (
        analysis.sort_values(["Theme","TradeScore","Avg30TurnoverCr"], ascending=[True,False,False])
                .groupby("Theme", as_index=False)
                .first()
    )

    # Final safety gate: commodity ETFs must never enter Equity candidates.
    id_text = (
        analysis["Symbol"].astype(str) + " " +
        analysis["Underlying"].astype(str) + " " +
        analysis["SecurityName"].astype(str)
    ).str.upper()

    analysis.loc[id_text.str.contains("SILVER|SILV", regex=True, na=False), "AssetClass"] = "SILVER"
    analysis.loc[
        id_text.str.contains("GOLD", regex=True, na=False) &
        ~id_text.str.contains("SILVER|SILV", regex=True, na=False),
        "AssetClass"
    ] = "GOLD"

    equity = analysis[analysis["AssetClass"].isin(["EQUITY","INTERNATIONAL"])].copy()
    non_equity = analysis[analysis["AssetClass"].isin(["GOLD","SILVER","DEBT","LIQUID"])].copy()

    # ------------------------------------------------------------------
    # ALL latest-session losers must come from the COMPLETE ETF universe.
    # `analysis` is already the NSE-volume-selected universe.
    # ------------------------------------------------------------------
    latest_all = etf.copy()

    latest_all["CurrentVolumeMultiple30D"] = np.where(
        latest_all["Avg30Volume"].fillna(0) > 0,
        latest_all["DayVolume"] / latest_all["Avg30Volume"],
        np.nan
    )
    latest_all["CurrentVsAvg30VolumePct"] = (
        latest_all["CurrentVolumeMultiple30D"] - 1.0
    ) * 100.0

    latest_all["RangePositionPct"] = np.where(
        (latest_all["NSE_High"] - latest_all["NSE_Low"]).fillna(0) > 0,
        (latest_all["LTP"] - latest_all["NSE_Low"]) /
        (latest_all["NSE_High"] - latest_all["NSE_Low"]) * 100.0,
        np.nan
    )
    latest_all["RangePositionPct"] = latest_all["RangePositionPct"].clip(0, 100)

    latest_all["DipVolumeCharacter"] = latest_all.apply(dip_volume_character, axis=1)

    # Full-universe aggressive dip score: latest fall + latest volume + turnover.
    latest_all["FullDipMagnitudeScore"] = (
        latest_all["NSE_ChangePct"].fillna(0).abs().clip(0, 4.0) / 4.0 * 100
    )
    latest_all.loc[latest_all["NSE_ChangePct"] >= 0, "FullDipMagnitudeScore"] = 0

    latest_all["FullDipVolumeScore"] = latest_all["CurrentVolumeMultiple30D"].apply(
        lambda x: 35 if pd.isna(x) else (
            100 if x >= 4.0 else
            95 if x >= 3.0 else
            88 if x >= 2.0 else
            78 if x >= 1.3 else
            62 if x >= 0.8 else
            40
        )
    )
    latest_all["FullDipTurnoverScore"] = (
        latest_all["TodayTurnoverCr"].fillna(0).clip(0,2.0) / 2.0 * 100
    ).clip(0,100)

    latest_all["AggressiveDipScore"] = (
        0.45*latest_all["FullDipMagnitudeScore"] +
        0.45*latest_all["FullDipVolumeScore"] +
        0.10*latest_all["FullDipTurnoverScore"]
    ).clip(0,100)

    def aggressive_dip_signal(row):
        if pd.isna(row.get("NSE_ChangePct")) or row.get("NSE_ChangePct") >= 0:
            return "NO DIP"

        score = row.get("AggressiveDipScore", np.nan)
        volx = row.get("CurrentVolumeMultiple30D", np.nan)
        fall = abs(row.get("NSE_ChangePct", 0))

        # Mean-reversion interpretation:
        # stronger current fall + exceptional volume = higher dip-buy priority.
        # Very high relative volume can upgrade the signal even when the fall is < 1%.
        if (
            pd.notna(volx) and volx >= 5.0
            and fall >= 0.50
        ):
            return "A+ EXTREME VOLUME DIP"

        if (
            pd.notna(volx) and volx >= 3.0
            and fall >= 0.40
        ):
            return "STRONG VOLUME DIP BUY"

        if (
            pd.notna(score) and score >= 65
        ):
            return "STRONG DIP BUY"

        if (
            pd.notna(volx) and volx >= 2.0
            and fall >= 0.25
        ):
            return "VOLUME DIP BUY"

        if pd.notna(score) and score >= 45:
            return "DIP BUY WATCH"

        return "LOW PRIORITY DIP"

    latest_all["AggressiveDipSignal"] = latest_all.apply(aggressive_dip_signal, axis=1)

    def liquidity_rejection_reason(row):
        # No later liquidity rejection: ETF already passed NSE >= 1 lakh current volume.
        return "ELIGIBLE"

    latest_all["DipEligibilityResult"] = latest_all.apply(liquidity_rejection_reason, axis=1)

    # Bring qualified analysis scores/signals into the full-universe view.
    enrich_cols = [
        c for c in [
            "Symbol","Theme_Quadrant","Quadrant",
            "ThemeRRGScore","ETF_RRG_Score","TrendScore","TrendPoints","MomentumScore",
            "DipBuyScore","DipSignal","BuyingPressureScore","BuyingSignal"
        ] if c in analysis.columns
    ]
    latest_all = latest_all.merge(
        analysis[enrich_cols].drop_duplicates("Symbol"),
        on="Symbol", how="left"
    )

    # Keep liquidity as a separate execution classification, but do not use it
    # to suppress or demote a deep/high-volume dip opportunity.
    latest_all["DipSignal"] = latest_all["AggressiveDipSignal"]

    # ------------------------------------------------------------
    # TOP 10 DAILY LOSERS — selected NSE Volume >= 1 lakh universe
    # ------------------------------------------------------------
    # No -1% threshold. Take every negative ETF in the selected universe.
    # Larger absolute fall + larger latest NSE volume receive higher rank.
    losers_pool = latest_all[
        latest_all["AssetClass"].isin(["EQUITY","INTERNATIONAL"]) &
        (latest_all["NSE_ChangePct"] < 0)
    ].copy()

    if not losers_pool.empty:
        losers_pool["AbsPctFall"] = losers_pool["NSE_ChangePct"].abs()
        losers_pool["PctMoveRankScore"] = (
            losers_pool["AbsPctFall"].rank(pct=True, method="average") * 100
        )
        losers_pool["CurrentVolumeRankScore"] = (
            losers_pool["DayVolume"].rank(pct=True, method="average") * 100
        )
        losers_pool["DailyLoserRankScore"] = (
            0.60 * losers_pool["PctMoveRankScore"] +
            0.40 * losers_pool["CurrentVolumeRankScore"]
        )
        all_latest_losers = (
            losers_pool
            .sort_values(
                ["DailyLoserRankScore","AbsPctFall","DayVolume"],
                ascending=[False,False,False]
            )
            .head(10)
            .reset_index(drop=True)
        )
        all_latest_losers.insert(0, "DipRank", np.arange(1, len(all_latest_losers)+1))
    else:
        all_latest_losers = losers_pool.copy()
        all_latest_losers.insert(0, "DipRank", pd.Series(dtype=int))

    # Qualified dips still come only from the NSE-volume-selected analysis universe.
    dip_candidates = all_latest_losers.copy()

    # ------------------------------------------------------------------
    # FULL latest-session buying analysis.
    # Positive ranking is built only after the primary NSE-volume universe filter.
    # Every positive equity/international ETF in the selected universe is eligible for Top-10 ranking.
    # ------------------------------------------------------------------
    latest_all["BuyingPriceScore"] = (
        latest_all["NSE_ChangePct"].fillna(0).clip(lower=0, upper=4) / 4 * 100
    )
    latest_all["BuyingRangeScore"] = latest_all["RangePositionPct"].fillna(50)

    # Latest-data-only buying score for the COMPLETE ETF universe.
    # Price move 40%, session-range strength 35%, volume participation 25%.
    latest_all["LatestBuyingScore"] = (
        0.40 * latest_all["BuyingPriceScore"] +
        0.35 * latest_all["BuyingRangeScore"] +
        0.25 * (
            20 + latest_all["CurrentVolumeMultiple30D"].fillna(0).clip(0, 3) * (80/3)
        ).clip(0,100)
    ).clip(0,100)

    def full_buying_signal(row):
        if row["AssetClass"] not in ["EQUITY","INTERNATIONAL"]:
            return "N/A"
        if pd.isna(row["NSE_ChangePct"]) or row["NSE_ChangePct"] <= 0:
            return "NO BUYING SIGNAL"

        volx = row.get("CurrentVolumeMultiple30D", np.nan)
        rp = row.get("RangePositionPct", np.nan)
        score = row.get("LatestBuyingScore", np.nan)
        chg = row.get("NSE_ChangePct", np.nan)

        # Every ETF here already passed the NSE >= 1 lakh current-volume filter.
        # Signal strength now applies to ALL positive Top-10 candidates.
        if (
            pd.notna(volx) and volx >= 1.75
            and pd.notna(rp) and rp >= 80
            and pd.notna(score) and score >= 75
        ):
            return "A+ HEAVY BUYING"

        if (
            pd.notna(volx) and volx >= 1.25
            and pd.notna(rp) and rp >= 70
            and pd.notna(score) and score >= 65
        ):
            return "STRONG BUYING"

        if (
            pd.notna(volx) and volx >= 1.00
            and pd.notna(rp) and rp >= 60
            and pd.notna(score) and score >= 55
        ):
            return "BUYING PRESSURE"

        if pd.notna(rp) and rp < 35:
            return "WEAK POSITIVE / FADING"

        # Positive ETF with >=1 lakh current NSE volume but without strong
        # confirmation from range/relative volume.
        return "POSITIVE BUY WATCH"

    latest_all["BuyingSignalFull"] = latest_all.apply(full_buying_signal, axis=1)

    def buying_eligibility_result(row):
        # No later liquidity rejection: ETF already passed NSE >= 1 lakh current volume.
        return "ELIGIBLE"

    latest_all["BuyingEligibilityResult"] = latest_all.apply(buying_eligibility_result, axis=1)

    # ------------------------------------------------------------
    # GROUP 1 — EQUITY SWING BUY PRIORITY
    # Indian Equity + International Equity
    # ------------------------------------------------------------
    # The positive side is built in two stages:
    #   A) Rank up to the Top 5 LEADING / IMPROVING sectors/themes.
    #   B) Rank only high-quality ETFs inside those Top 5 sectors/themes.
    #
    # Rank 1 is not forced. If no ETF passes the strict swing-entry quality
    # conditions, the dashboard shows NO QUALIFIED BUY NOW.

    group1_equity = analysis[
        analysis["AssetClass"].isin(["EQUITY","INTERNATIONAL"])
    ].copy()

    def _pct_rank_100(s):
        s = pd.to_numeric(s, errors="coerce")
        if s.notna().sum() <= 1:
            return pd.Series(50.0, index=s.index)
        return s.rank(pct=True, method="average").fillna(0.5) * 100.0

    # ---------- A) SECTOR / THEME STRENGTH ----------
    sector_base = group1_equity.copy()

    if not sector_base.empty:
        sector_stats = (
            sector_base.groupby("Theme", dropna=False)
            .agg(
                SectorRRG=("Theme_Quadrant","first"),
                SectorRRGBase=("ThemeRRGScore","first"),
                SectorRRGLevel=("Theme_RS_Momentum","first"),
                SectorRRGChange5D=("Theme_RS_Momentum_5D_Change","first"),
                TodaySectorPct=("NSE_ChangePct","median"),
                ETFCount=("Symbol","count"),
                PositiveETFCount=("NSE_ChangePct", lambda s: int((pd.to_numeric(s, errors="coerce") > 0).sum())),
                MedianVolumeVs30D=("VolumePaceVs30D","median"),
                SectorTurnoverCr=("TodayTurnoverCr","sum")
            )
            .reset_index()
        )
        sector_stats["BreadthPct"] = np.where(
            sector_stats["ETFCount"] > 0,
            sector_stats["PositiveETFCount"] / sector_stats["ETFCount"] * 100.0,
            0.0
        )

        # Only strong rotation quadrants can supply Buy Priority ETFs.
        qualified_sectors = sector_stats[
            sector_stats["SectorRRG"].isin(["LEADING","IMPROVING"])
        ].copy()

        if not qualified_sectors.empty:
            qualified_sectors["RRGStrengthScore"] = (
                pd.to_numeric(qualified_sectors["SectorRRGBase"], errors="coerce")
                .fillna(
                    pd.Series(
                        np.where(qualified_sectors["SectorRRG"].eq("LEADING"), 95.0, 80.0),
                        index=qualified_sectors.index
                    )
                )
                .clip(0,100)
            )
            qualified_sectors["RRGMomentumScore"] = (
                0.60 * _pct_rank_100(qualified_sectors["SectorRRGLevel"]) +
                0.40 * _pct_rank_100(qualified_sectors["SectorRRGChange5D"])
            ).clip(0,100)
            qualified_sectors["TodayPerformanceScore"] = _pct_rank_100(
                qualified_sectors["TodaySectorPct"]
            ).clip(0,100)
            qualified_sectors["BreadthScore"] = qualified_sectors["BreadthPct"].clip(0,100)
            qualified_sectors["ParticipationScore"] = (
                0.60 * _pct_rank_100(qualified_sectors["MedianVolumeVs30D"]) +
                0.40 * _pct_rank_100(qualified_sectors["SectorTurnoverCr"])
            ).clip(0,100)

            qualified_sectors["SectorStrengthScore"] = (
                0.35 * qualified_sectors["RRGStrengthScore"] +
                0.20 * qualified_sectors["RRGMomentumScore"] +
                0.20 * qualified_sectors["TodayPerformanceScore"] +
                0.15 * qualified_sectors["BreadthScore"] +
                0.10 * qualified_sectors["ParticipationScore"]
            ).clip(0,100)

            top5_sectors = (
                qualified_sectors
                .sort_values(
                    ["SectorStrengthScore","TodayPerformanceScore","SectorTurnoverCr"],
                    ascending=[False,False,False]
                )
                .head(5)
                .reset_index(drop=True)
            )
            top5_sectors.insert(0, "SectorRank", np.arange(1, len(top5_sectors)+1))
        else:
            top5_sectors = qualified_sectors.copy()
            top5_sectors.insert(0, "SectorRank", pd.Series(dtype=int))
    else:
        sector_stats = pd.DataFrame()
        top5_sectors = pd.DataFrame(columns=[
            "SectorRank","Theme","SectorRRG","SectorStrengthScore","TodaySectorPct",
            "BreadthPct","MedianVolumeVs30D","SectorTurnoverCr"
        ])

    # ---------- B) ETF MOMENTUM PHASE ----------
    def momentum_phase_and_score(row):
        # Latest NSE % change is authoritative for the current session.
        d = row.get("NSE_ChangePct", row.get("DayReturnPct", np.nan))
        w = row.get("Week1ReturnPct", np.nan)
        m1 = row.get("Month1ReturnPct", np.nan)
        m3 = row.get("Month3ReturnPct", np.nan)

        if any(pd.isna(x) for x in [d,w,m1,m3]):
            return pd.Series(["NO DATA", 45.0])

        # 1) Strong turnaround:
        # Longer-term return is still negative, but the damage is healing,
        # the week is no longer negative, and today is positive.
        if m3 < 0 and m1 > m3 and w >= 0 and d > 0:
            return pd.Series(["STRONG TURNAROUND", 100.0])

        # 2) Fresh momentum:
        # Longer-term performance is still modest, while 1M / 1W / today
        # have turned positive. This is preferred to chasing an extended move.
        if m3 <= 5 and m1 > 0 and w >= 0 and d > 0:
            return pd.Series(["FRESH MOMENTUM", 90.0])

        # 3) Healthy trend:
        # 3M and 1M are positive. A small weekly pause up to -0.50% is allowed
        # when the ETF has turned positive again today.
        if m3 > 0 and m1 > 0 and w >= -0.50 and d > 0:
            return pd.Series(["HEALTHY TREND", 75.0])

        # 4) Early improvement:
        # Long-term return is negative, but 1M is less negative / improving
        # and the ETF is positive today.
        if m3 < 0 and m1 > m3 and d > 0:
            return pd.Series(["EARLY IMPROVEMENT", 65.0])

        # 5) Fading:
        # Longer-term trend is positive, but recent 1M or 1W deterioration
        # is meaningful despite today's positive move.
        if m3 > 0 and (m1 < 0 or w < -0.50):
            return pd.Series(["FADING", 25.0])

        return pd.Series(["STILL WEAK", 10.0])

    # ---------- C) ENTRY STRETCH ----------
    def entry_stretch_label(x):
        if pd.isna(x):
            return "NO DATA"
        if x < 0:
            return "BELOW TREND"
        if x < 1.0:
            return "FRESH"
        if x < 2.0:
            return "HEALTHY"
        if x < 3.0:
            return "EXTENDED"
        return "OVEREXTENDED"

    def entry_stretch_score(label):
        return {
            "FRESH": 100.0,
            "HEALTHY": 85.0,
            "EXTENDED": 45.0,
            "OVEREXTENDED": 0.0,
            "BELOW TREND": 25.0,
            "NO DATA": 40.0,
        }.get(label, 40.0)

    # ---------- D) BUILD ETF CANDIDATE POOL ----------
    if not top5_sectors.empty:
        sector_map = top5_sectors[
            ["Theme","SectorRank","SectorRRG","SectorStrengthScore"]
        ].copy()

        buy_pool = (
            group1_equity[
                (group1_equity["NSE_ChangePct"] > 0) &
                (group1_equity["Quadrant"].isin(["LEADING","IMPROVING"]))
            ]
            .merge(sector_map, on="Theme", how="inner")
            .copy()
        )
    else:
        buy_pool = group1_equity.iloc[0:0].copy()

    if not buy_pool.empty:
        buy_pool[["MomentumPhase","MomentumImprovementScore"]] = buy_pool.apply(
            momentum_phase_and_score, axis=1
        )

        buy_pool["EntryStretch"] = buy_pool["EntryStretchATR"].apply(entry_stretch_label)
        buy_pool["EntryStretchScore"] = buy_pool["EntryStretch"].apply(entry_stretch_score)

        # Today's price strength is relative to the qualified ETF pool.
        buy_pool["TodayPriceStrengthScore"] = _pct_rank_100(
            buy_pool["NSE_ChangePct"]
        ).clip(0,100)

        # Tradability = actual current volume + money traded + participation
        # compared with the ETF's own normal volume.
        buy_pool["CurrentNSEVolumeScore"] = _pct_rank_100(
            buy_pool["DayVolume"]
        ).clip(0,100)
        buy_pool["TurnoverScore"] = _pct_rank_100(
            buy_pool["TodayTurnoverCr"]
        ).clip(0,100)

        buy_pool["VolumeVs30DScore"] = np.select(
            [
                buy_pool["VolumePaceVs30D"].fillna(0) >= 3.0,
                buy_pool["VolumePaceVs30D"].fillna(0) >= 2.0,
                buy_pool["VolumePaceVs30D"].fillna(0) >= 1.5,
                buy_pool["VolumePaceVs30D"].fillna(0) >= 1.0,
                buy_pool["VolumePaceVs30D"].fillna(0) >= 0.75,
                buy_pool["VolumePaceVs30D"].fillna(0) >= 0.50,
            ],
            [100, 92, 84, 72, 58, 42],
            default=20
        ).astype(float)

        buy_pool["TradabilityScore"] = (
            0.40 * buy_pool["CurrentNSEVolumeScore"] +
            0.35 * buy_pool["TurnoverScore"] +
            0.25 * buy_pool["VolumeVs30DScore"]
        ).clip(0,100)

        buy_pool["Tradability"] = pd.cut(
            buy_pool["TradabilityScore"],
            bins=[-np.inf,35,50,65,80,np.inf],
            labels=["WEAK","FAIR","GOOD","STRONG","EXCELLENT"],
            right=False
        ).astype(str)

        # Price holding score: raw position in today's NSE high-low range.
        buy_pool["NearDayHighScore"] = buy_pool["RangePositionPct"].fillna(50).clip(0,100)

        # Trend / RRG component scores.
        buy_pool["TrendStrengthScore"] = np.select(
            [
                buy_pool["TrendPoints"] >= 4,
                buy_pool["TrendPoints"] == 3,
                buy_pool["TrendPoints"] == 2,
                buy_pool["TrendPoints"] == 1,
            ],
            [100,80,55,30],
            default=10
        ).astype(float)

        buy_pool["ETFRRGStrengthScore"] = (
            pd.to_numeric(buy_pool["ETF_RRG_Score"], errors="coerce")
            .fillna(
                pd.Series(
                    np.where(buy_pool["Quadrant"].eq("LEADING"), 92.0, 78.0),
                    index=buy_pool.index
                )
            )
            .clip(0,100)
        )

        buy_pool["BuyQualityScore"] = (
            0.20 * buy_pool["SectorStrengthScore"] +
            0.20 * buy_pool["MomentumImprovementScore"] +
            0.15 * buy_pool["TradabilityScore"] +
            0.15 * buy_pool["NearDayHighScore"] +
            0.10 * buy_pool["TodayPriceStrengthScore"] +
            0.10 * buy_pool["ETFRRGStrengthScore"] +
            0.05 * buy_pool["TrendStrengthScore"] +
            0.05 * buy_pool["EntryStretchScore"]
        ).clip(0,100)

        # Strict swing-entry gate. This prevents weak-volume, fading,
        # overextended or structurally weak ETFs from becoming Buy Rank 1.
        buy_pool["StrictBuyQualified"] = (
            (buy_pool["VolumePaceVs30D"].fillna(0) >= 0.50) &
            (buy_pool["TradabilityScore"].fillna(0) >= 50) &
            (buy_pool["RangePositionPct"].fillna(0) >= 65) &
            (buy_pool["TrendPoints"].fillna(0) >= 3) &
            (buy_pool["MomentumImprovementScore"].fillna(0) >= 60) &
            (~buy_pool["EntryStretch"].eq("OVEREXTENDED")) &
            (buy_pool["BuyQualityScore"].fillna(0) >= 65)
        )

        def buy_priority_signal(row):
            score = row.get("BuyQualityScore", np.nan)
            if not bool(row.get("StrictBuyQualified", False)):
                return "WATCH ONLY"
            if pd.notna(score) and score >= 85:
                return "TOP BUY SETUP"
            if pd.notna(score) and score >= 75:
                return "STRONG BUY SETUP"
            return "GOOD BUY SETUP"

        buy_pool["BuyPrioritySignal"] = buy_pool.apply(buy_priority_signal, axis=1)

        def why_not_qualified(row):
            reasons = []

            volpace = row.get("VolumePaceVs30D", np.nan)
            trad = row.get("TradabilityScore", np.nan)
            near_high = row.get("RangePositionPct", np.nan)
            trend = row.get("TrendPoints", np.nan)
            momentum = row.get("MomentumImprovementScore", np.nan)
            stretch = str(row.get("EntryStretch", ""))
            quality = row.get("BuyQualityScore", np.nan)

            if pd.isna(volpace) or volpace < 0.50:
                reasons.append(
                    f"Volume vs normal {0 if pd.isna(volpace) else volpace:.2f}x < 0.50x"
                )
            if pd.isna(trad) or trad < 50:
                reasons.append(
                    f"Tradability {0 if pd.isna(trad) else trad:.1f} < 50"
                )
            if pd.isna(near_high) or near_high < 65:
                reasons.append(
                    f"Near day high {0 if pd.isna(near_high) else near_high:.1f}% < 65%"
                )
            if pd.isna(trend) or trend < 3:
                reasons.append(
                    f"Trend {0 if pd.isna(trend) else int(trend)}/4 < 3/4"
                )
            if pd.isna(momentum) or momentum < 60:
                phase = str(row.get("MomentumPhase", "NO DATA"))
                reasons.append(
                    f"Momentum {phase} ({0 if pd.isna(momentum) else momentum:.0f}) < 60"
                )
            if stretch == "OVEREXTENDED":
                reasons.append("Entry stretch OVEREXTENDED")
            if pd.isna(quality) or quality < 65:
                reasons.append(
                    f"Buy quality {0 if pd.isna(quality) else quality:.1f} < 65"
                )

            return "; ".join(reasons) if reasons else "Qualified"

        buy_pool["WhyNotQualified"] = buy_pool.apply(why_not_qualified, axis=1)
        buy_pool["FailedRuleCount"] = buy_pool["WhyNotQualified"].apply(
            lambda x: 0 if x == "Qualified" else len(str(x).split("; "))
        )

        # Rebuild qualified table after diagnostics so detailed output includes
        # the same transparent rule fields.
        all_latest_gainers = (
            buy_pool[buy_pool["StrictBuyQualified"]]
            .sort_values(
                ["BuyQualityScore","SectorRank","TradabilityScore","RangePositionPct"],
                ascending=[False,True,False,False]
            )
            .head(10)
            .reset_index(drop=True)
        )
        if not all_latest_gainers.empty:
            all_latest_gainers.insert(
                0, "BuyRank", np.arange(1, len(all_latest_gainers)+1)
            )
        else:
            all_latest_gainers.insert(0, "BuyRank", pd.Series(dtype=int))

        positive_watchlist = (
            buy_pool[~buy_pool["StrictBuyQualified"]]
            .sort_values(
                ["FailedRuleCount","BuyQualityScore","SectorRank","TradabilityScore"],
                ascending=[True,False,True,False]
            )
            .reset_index(drop=True)
        )
        if not positive_watchlist.empty:
            positive_watchlist.insert(
                0, "NearBuyRank", np.arange(1, len(positive_watchlist)+1)
            )
    else:
        all_latest_gainers = buy_pool.copy()
        all_latest_gainers.insert(0, "BuyRank", pd.Series(dtype=int))
        positive_watchlist = buy_pool.copy()
        positive_watchlist["NearBuyRank"] = pd.Series(dtype=int)
        positive_watchlist["WhyNotQualified"] = pd.Series(dtype=str)
        positive_watchlist["FailedRuleCount"] = pd.Series(dtype=int)

    # Buy Priority contains only strict, sector-confirmed swing candidates.
    buying_opportunities = all_latest_gainers.copy()

    # ------------------------------------------------------------
    # GROUP 1 — EQUITY INTRADAY BUY PRIORITY
    # Indian Equity + International Equity
    # ------------------------------------------------------------
    # V4.5.2 — INTRADAY-ONLY UNIVERSE CHANGE
    #
    # Swing and negative/dip scanners are NOT changed:
    #   - they continue to use the hard NSE current-volume >= 1 lakh universe.
    #
    # Intraday no longer requires 1 lakh units.
    # It builds a dynamic TOP 50 ACTIVE EQUITY universe at the moment of the scan:
    #
    #   Intraday Activity Score =
    #       50% Current NSE Turnover
    #       30% Current NSE Volume
    #       20% Time-adjusted Volume Pace vs prior 30-day normal
    #
    # Gold / Silver / Debt / Liquid are excluded.
    # Only Equity + International Equity can enter the intraday Top 50.
    #
    # A broad provisional Top 120 is first selected using current NSE turnover +
    # current volume. Historical volume is downloaded only for missing names in
    # that broad set; the FINAL Top 50 uses the full 50/30/20 Activity Score.
    #
    # Detailed activity calculations stay in the background. The visible INTRADAY
    # sheet remains compact.

    try:
        if CURRENT_SESSION_MODE == "INTRADAY":
            if _mins < 9*60+40:
                INTRADAY_SCAN_WINDOW = "OPENING NOISE — WAIT"
            elif _mins <= 10*60+30:
                INTRADAY_SCAN_WINDOW = "PRIMARY INTRADAY WINDOW"
            elif _mins <= 12*60+30:
                INTRADAY_SCAN_WINDOW = "CONTINUATION WINDOW"
            elif _mins <= 13*60+30:
                INTRADAY_SCAN_WINDOW = "MIDDAY — SELECTIVE"
            elif _mins <= 14*60+30:
                INTRADAY_SCAN_WINDOW = "SECONDARY WINDOW"
            elif _mins <= 15*60+10:
                INTRADAY_SCAN_WINDOW = "LATE MOMENTUM WINDOW"
            else:
                INTRADAY_SCAN_WINDOW = "LATE ENTRY CAUTION"
        else:
            INTRADAY_SCAN_WINDOW = "MARKET CLOSED — REFERENCE ONLY"
    except Exception:
        INTRADAY_SCAN_WINDOW = "MARKET CLOSED — REFERENCE ONLY"

    try:
        INTRADAY_ACTIONABLE_WINDOW = bool(
            CURRENT_SESSION_MODE == "INTRADAY" and
            (9*60+40) <= _mins <= (15*60+10)
        )
    except Exception:
        INTRADAY_ACTIONABLE_WINDOW = False

    INTRADAY_ACTIVITY_TOP_N = 50
    INTRADAY_ACTIVITY_SEED_N = 120

    def _intraday_rank100(s):
        s = pd.to_numeric(s, errors="coerce")
        if s.notna().sum() <= 1:
            return pd.Series(50.0, index=s.index)
        return s.rank(pct=True, method="average").fillna(0.5) * 100.0

    # ---------- A) BUILD TOP-50 ACTIVE EQUITY UNIVERSE ----------
    intraday_master_all = master_all[
        master_all["AssetClass"].isin(["EQUITY", "INTERNATIONAL"])
    ].copy()

    intraday_master_all["ActivityCurrentVolume"] = pd.to_numeric(
        intraday_master_all["NSE_CurrentVolume"], errors="coerce"
    ).fillna(0)

    intraday_master_all["ActivityTurnoverCr"] = pd.to_numeric(
        intraday_master_all["NSE_CurrentValueCr"], errors="coerce"
    )

    intraday_master_all["ActivityTurnoverCr"] = intraday_master_all["ActivityTurnoverCr"].fillna(
        (
            pd.to_numeric(intraday_master_all["NSE_LTP"], errors="coerce").fillna(0) *
            intraday_master_all["ActivityCurrentVolume"]
        ) / 1e7
    )

    intraday_master_all = intraday_master_all[
        (intraday_master_all["ActivityCurrentVolume"] > 0) &
        (pd.to_numeric(intraday_master_all["NSE_LTP"], errors="coerce").fillna(0) > 0)
    ].copy()

    if not intraday_master_all.empty:
        intraday_master_all["SeedTurnoverRank"] = _intraday_rank100(
            intraday_master_all["ActivityTurnoverCr"]
        )
        intraday_master_all["SeedVolumeRank"] = _intraday_rank100(
            intraday_master_all["ActivityCurrentVolume"]
        )
        intraday_master_all["SeedActivityScore"] = (
            0.60 * intraday_master_all["SeedTurnoverRank"] +
            0.40 * intraday_master_all["SeedVolumeRank"]
        )

        if CURRENT_SESSION_MODE == "INTRADAY":
            intraday_seed = (
                intraday_master_all
                .sort_values(
                    ["SeedActivityScore", "ActivityTurnoverCr", "ActivityCurrentVolume"],
                    ascending=[False, False, False]
                )
                .head(INTRADAY_ACTIVITY_SEED_N)
                .copy()
            )
        else:
            core_syms = set(analysis["Symbol"].astype(str))
            intraday_seed = intraday_master_all[
                intraday_master_all["Symbol"].astype(str).isin(core_syms)
            ].copy()
    else:
        intraday_seed = intraday_master_all.copy()

    extra_prices = pd.DataFrame()
    extra_volumes = pd.DataFrame()
    extra_highs = pd.DataFrame()
    extra_lows = pd.DataFrame()
    extra_failed = []

    if CURRENT_SESSION_MODE == "INTRADAY" and not intraday_seed.empty:
        existing_tickers = set(prices.columns.astype(str))
        extra_tickers = [
            t for t in intraday_seed["YFTicker"].dropna().astype(str).tolist()
            if t not in existing_tickers
        ]

        if extra_tickers:
            print(
                f"   Intraday activity universe: broad Top {len(intraday_seed)} "
                f"Equity / International ETFs"
            )
            print(f"   Extra histories needed      : {len(extra_tickers)}")
            extra_prices, extra_volumes, extra_highs, extra_lows, extra_failed = (
                yf_batch_download(extra_tickers)
            )

    def _intraday_history_series(ticker, base_df, extra_df):
        if ticker in base_df.columns:
            return base_df[ticker].dropna().copy()
        if ticker in extra_df.columns:
            return extra_df[ticker].dropna().copy()
        return pd.Series(dtype=float)

    analysis_by_symbol = {
        str(r["Symbol"]): r.to_dict()
        for _, r in analysis.iterrows()
    }

    activity_rows = []

    for _, mr in intraday_seed.iterrows():
        sym = str(mr["Symbol"])
        tk = str(mr["YFTicker"])

        if sym in analysis_by_symbol:
            d = analysis_by_symbol[sym].copy()
            d["DayVolume"] = pd.to_numeric(
                pd.Series([mr.get("NSE_CurrentVolume", np.nan)]), errors="coerce"
            ).iloc[0]
            d["TodayTurnoverCr"] = pd.to_numeric(
                pd.Series([mr.get("NSE_CurrentValueCr", np.nan)]), errors="coerce"
            ).iloc[0]
            if pd.isna(d["TodayTurnoverCr"]):
                _ltp0 = pd.to_numeric(
                    pd.Series([mr.get("NSE_LTP", np.nan)]), errors="coerce"
                ).iloc[0]
                d["TodayTurnoverCr"] = (
                    (_ltp0 * d["DayVolume"]) / 1e7
                    if pd.notna(_ltp0) and pd.notna(d["DayVolume"])
                    else np.nan
                )
            d["IntradaySupplemental"] = False
            activity_rows.append(d)
            continue

        # Supplemental intraday-only ETF: may be below 1 lakh current volume.
        p = _intraday_history_series(tk, prices, extra_prices)
        v = _intraday_history_series(tk, volumes, extra_volumes)
        h = _intraday_history_series(tk, highs, extra_highs)
        l = _intraday_history_series(tk, lows, extra_lows)

        if len(p) < MIN_RRG_OBS:
            continue

        try:
            p.index = pd.to_datetime(p.index).tz_localize(None)
            v.index = pd.to_datetime(v.index).tz_localize(None)
            h.index = pd.to_datetime(h.index).tz_localize(None)
            l.index = pd.to_datetime(l.index).tz_localize(None)
        except Exception:
            p.index = pd.to_datetime(p.index)
            v.index = pd.to_datetime(v.index)
            h.index = pd.to_datetime(h.index)
            l.index = pd.to_datetime(l.index)

        ltp = pd.to_numeric(pd.Series([mr.get("NSE_LTP", np.nan)]), errors="coerce").iloc[0]
        prev = pd.to_numeric(pd.Series([mr.get("NSE_PrevClose", np.nan)]), errors="coerce").iloc[0]
        opn = pd.to_numeric(pd.Series([mr.get("NSE_Open", np.nan)]), errors="coerce").iloc[0]
        hi = pd.to_numeric(pd.Series([mr.get("NSE_High", np.nan)]), errors="coerce").iloc[0]
        lo = pd.to_numeric(pd.Series([mr.get("NSE_Low", np.nan)]), errors="coerce").iloc[0]
        chg = pd.to_numeric(pd.Series([mr.get("NSE_ChangePct", np.nan)]), errors="coerce").iloc[0]
        dayvol = pd.to_numeric(pd.Series([mr.get("NSE_CurrentVolume", np.nan)]), errors="coerce").iloc[0]
        turnover = pd.to_numeric(pd.Series([mr.get("NSE_CurrentValueCr", np.nan)]), errors="coerce").iloc[0]
        if pd.isna(turnover) and pd.notna(ltp) and pd.notna(dayvol):
            turnover = ltp * dayvol / 1e7

        v_completed = v[v.index < current_price_date]
        prior30 = v_completed.tail(30).mean() if len(v_completed) else np.nan

        range_pos = (
            (ltp - lo) / (hi - lo) * 100.0
            if pd.notna(ltp) and pd.notna(hi) and pd.notna(lo) and hi > lo
            else np.nan
        )
        if pd.notna(range_pos):
            range_pos = float(np.clip(range_pos, 0, 100))

        completed_mask = p.index < current_price_date
        atr20 = atr_from_history(
            p.loc[completed_mask],
            h.reindex(p.index).loc[completed_mask],
            l.reindex(p.index).loc[completed_mask],
            20
        )

        p_rrg = p.copy()
        if pd.notna(ltp):
            p_rrg.loc[current_price_date] = float(ltp)
            p_rrg = p_rrg.sort_index()

        rr = rrg_series(p_rrg, bench)
        if rr.empty:
            etf_quad = "NO DATA"
            etf_rrg_score = 40.0
        else:
            etf_quad = quadrant(
                rr.iloc[-1]["RS_Ratio"], rr.iloc[-1]["RS_Momentum"]
            )
            etf_rrg_score = (
                92.0 if etf_quad == "LEADING"
                else 78.0 if etf_quad == "IMPROVING"
                else 45.0 if etf_quad == "WEAKENING"
                else 20.0
            )

        activity_rows.append({
            "Symbol": mr["Symbol"],
            "Underlying": mr["Underlying"],
            "SecurityName": mr["SecurityName"],
            "AssetClass": mr["AssetClass"],
            "Theme": mr["Theme"],
            "YFTicker": tk,
            "NSE_DownloadTimestamp": mr.get("NSE_DownloadTimestamp", NSE_DOWNLOAD_TIMESTAMP),
            "NSE_AsOf": mr.get("NSE_AsOf", ""),
            "NSE_PrevClose": prev,
            "NSE_Open": opn,
            "NSE_High": hi,
            "NSE_Low": lo,
            "NSE_ChangePct": chg,
            "LTP": ltp,
            "DayVolume": dayvol,
            "TodayTurnoverCr": turnover,
            "Prior30AvgVolume": prior30,
            "RangePositionPct": range_pos,
            "ATR20": atr20,
            "Quadrant": etf_quad,
            "ETF_RRG_Score": etf_rrg_score,
            "LatestDataMode": CURRENT_SESSION_MODE,
            "IntradaySupplemental": True,
        })

    intraday_activity = pd.DataFrame(activity_rows)

    if not intraday_activity.empty:
        intraday_activity["DayVolume"] = pd.to_numeric(
            intraday_activity["DayVolume"], errors="coerce"
        ).fillna(0)
        intraday_activity["TodayTurnoverCr"] = pd.to_numeric(
            intraday_activity["TodayTurnoverCr"], errors="coerce"
        ).fillna(0)
        intraday_activity["Prior30AvgVolume"] = pd.to_numeric(
            intraday_activity["Prior30AvgVolume"], errors="coerce"
        )

        intraday_activity["VolumePaceVs30D"] = np.where(
            intraday_activity["Prior30AvgVolume"].fillna(0) > 0,
            intraday_activity["DayVolume"] /
            (intraday_activity["Prior30AvgVolume"] * MARKET_PROGRESS_FRACTION),
            np.nan
        )

        intraday_activity["ActivityTurnoverScore"] = _intraday_rank100(
            intraday_activity["TodayTurnoverCr"]
        ).clip(0, 100)
        intraday_activity["ActivityVolumeScore"] = _intraday_rank100(
            intraday_activity["DayVolume"]
        ).clip(0, 100)
        intraday_activity["ActivityVolumePaceScore"] = _intraday_rank100(
            intraday_activity["VolumePaceVs30D"]
        ).clip(0, 100)

        intraday_activity["IntradayActivityScore"] = (
            0.50 * intraday_activity["ActivityTurnoverScore"] +
            0.30 * intraday_activity["ActivityVolumeScore"] +
            0.20 * intraday_activity["ActivityVolumePaceScore"]
        ).clip(0, 100)

        intraday_activity_universe = (
            intraday_activity
            .sort_values(
                ["IntradayActivityScore", "TodayTurnoverCr", "DayVolume"],
                ascending=[False, False, False]
            )
            .head(INTRADAY_ACTIVITY_TOP_N)
            .reset_index(drop=True)
        )
        intraday_activity_universe.insert(
            0, "IntradayActivityRank",
            np.arange(1, len(intraday_activity_universe) + 1)
        )
    else:
        intraday_activity_universe = pd.DataFrame()

    INTRADAY_ACTIVITY_COUNT = len(intraday_activity_universe)

    # ---------- B) INTRADAY EQUITY GROUP STRENGTH ----------
    if not intraday_activity_universe.empty:
        intraday_group_stats = (
            intraday_activity_universe.groupby("Theme", dropna=False)
            .agg(
                TodaySectorPct=("NSE_ChangePct", "median"),
                ETFCount=("Symbol", "count"),
                PositiveETFCount=(
                    "NSE_ChangePct",
                    lambda s: int(
                        (pd.to_numeric(s, errors="coerce") > 0).sum()
                    )
                ),
                MedianVolumeVs30D=("VolumePaceVs30D", "median"),
                SectorTurnoverCr=("TodayTurnoverCr", "sum")
            )
            .reset_index()
        )

        intraday_group_stats["BreadthPct"] = np.where(
            intraday_group_stats["ETFCount"] > 0,
            intraday_group_stats["PositiveETFCount"] /
            intraday_group_stats["ETFCount"] * 100.0,
            0.0
        )

        if 'qualified_sectors' in locals() and not qualified_sectors.empty:
            rrg_meta_cols = [
                c for c in [
                    "Theme", "SectorRRG", "RRGStrengthScore",
                    "RRGMomentumScore"
                ] if c in qualified_sectors.columns
            ]
            intraday_group_stats = intraday_group_stats.merge(
                qualified_sectors[rrg_meta_cols].drop_duplicates("Theme"),
                on="Theme", how="left"
            )
        else:
            intraday_group_stats["SectorRRG"] = np.nan
            intraday_group_stats["RRGStrengthScore"] = np.nan
            intraday_group_stats["RRGMomentumScore"] = np.nan

        intraday_group_pool = intraday_group_stats[
            intraday_group_stats["SectorRRG"].isin(["LEADING", "IMPROVING"]) &
            (
                pd.to_numeric(
                    intraday_group_stats["TodaySectorPct"], errors="coerce"
                ).fillna(0) > 0
            )
        ].copy()

        if not intraday_group_pool.empty:
            intraday_group_pool["TodayPerformanceScore"] = _intraday_rank100(
                intraday_group_pool["TodaySectorPct"]
            ).clip(0, 100)
            intraday_group_pool["BreadthScore"] = (
                pd.to_numeric(
                    intraday_group_pool["BreadthPct"], errors="coerce"
                ).fillna(0).clip(0, 100)
            )
            intraday_group_pool["ParticipationScore"] = (
                0.60 * _intraday_rank100(
                    intraday_group_pool["MedianVolumeVs30D"]
                ) +
                0.40 * _intraday_rank100(
                    intraday_group_pool["SectorTurnoverCr"]
                )
            ).clip(0, 100)

            intraday_group_pool["IntradayGroupStrengthScore"] = (
                0.25 * pd.to_numeric(
                    intraday_group_pool["RRGStrengthScore"], errors="coerce"
                ).fillna(
                    pd.Series(
                        np.where(
                            intraday_group_pool["SectorRRG"].eq("LEADING"),
                            92.0, 78.0
                        ),
                        index=intraday_group_pool.index
                    )
                ) +
                0.15 * pd.to_numeric(
                    intraday_group_pool["RRGMomentumScore"], errors="coerce"
                ).fillna(50.0) +
                0.25 * intraday_group_pool["TodayPerformanceScore"] +
                0.20 * intraday_group_pool["BreadthScore"] +
                0.15 * intraday_group_pool["ParticipationScore"]
            ).clip(0, 100)

            top5_intraday_groups = (
                intraday_group_pool
                .sort_values(
                    [
                        "IntradayGroupStrengthScore",
                        "TodayPerformanceScore",
                        "SectorTurnoverCr"
                    ],
                    ascending=[False, False, False]
                )
                .head(5)
                .reset_index(drop=True)
            )
            top5_intraday_groups.insert(
                0, "IntradayGroupRank",
                np.arange(1, len(top5_intraday_groups) + 1)
            )
        else:
            top5_intraday_groups = pd.DataFrame(columns=[
                "IntradayGroupRank", "Theme", "SectorRRG",
                "IntradayGroupStrengthScore", "TodaySectorPct",
                "BreadthPct", "MedianVolumeVs30D", "SectorTurnoverCr"
            ])
    else:
        top5_intraday_groups = pd.DataFrame(columns=[
            "IntradayGroupRank", "Theme", "SectorRRG",
            "IntradayGroupStrengthScore", "TodaySectorPct",
            "BreadthPct", "MedianVolumeVs30D", "SectorTurnoverCr"
        ])

    # ---------- C) INTRADAY ETF CANDIDATE POOL ----------
    if not top5_intraday_groups.empty and not intraday_activity_universe.empty:
        intraday_group_map = top5_intraday_groups[[
            "Theme", "IntradayGroupRank",
            "SectorRRG", "IntradayGroupStrengthScore"
        ]].copy()

        intraday_pool = (
            intraday_activity_universe[
                (
                    pd.to_numeric(
                        intraday_activity_universe["NSE_ChangePct"],
                        errors="coerce"
                    ).fillna(0) > 0
                ) &
                (
                    intraday_activity_universe["Quadrant"]
                    .isin(["LEADING", "IMPROVING"])
                )
            ]
            .merge(intraday_group_map, on="Theme", how="inner")
            .copy()
        )
    else:
        intraday_pool = pd.DataFrame()

    if not intraday_pool.empty:
        intraday_pool["IntradayOpenDrivePct"] = np.where(
            pd.to_numeric(
                intraday_pool["NSE_Open"], errors="coerce"
            ).fillna(0) > 0,
            (
                pd.to_numeric(intraday_pool["LTP"], errors="coerce") /
                pd.to_numeric(intraday_pool["NSE_Open"], errors="coerce") - 1.0
            ) * 100.0,
            np.nan
        )

        intraday_pool["IntradayMoveATR"] = np.where(
            pd.to_numeric(
                intraday_pool["ATR20"], errors="coerce"
            ).fillna(0) > 0,
            (
                pd.to_numeric(intraday_pool["LTP"], errors="coerce") -
                pd.to_numeric(
                    intraday_pool["NSE_PrevClose"], errors="coerce"
                )
            ) / pd.to_numeric(intraday_pool["ATR20"], errors="coerce"),
            np.nan
        )

        intraday_pool["IntradayTodayStrengthScore"] = _intraday_rank100(
            intraday_pool["NSE_ChangePct"]
        ).clip(0, 100)

        intraday_pool["IntradayOpenDriveScore"] = _intraday_rank100(
            intraday_pool["IntradayOpenDrivePct"]
        ).clip(0, 100)

        intraday_pool["IntradayNearHighScore"] = (
            pd.to_numeric(
                intraday_pool["RangePositionPct"], errors="coerce"
            ).fillna(50).clip(0, 100)
        )

        intraday_pool["IntradayVolumePaceScore"] = np.select(
            [
                intraday_pool["VolumePaceVs30D"].fillna(0) >= 3.0,
                intraday_pool["VolumePaceVs30D"].fillna(0) >= 2.0,
                intraday_pool["VolumePaceVs30D"].fillna(0) >= 1.5,
                intraday_pool["VolumePaceVs30D"].fillna(0) >= 1.0,
                intraday_pool["VolumePaceVs30D"].fillna(0) >= 0.75,
                intraday_pool["VolumePaceVs30D"].fillna(0) >= 0.50,
            ],
            [100, 94, 88, 78, 65, 45],
            default=20
        ).astype(float)

        intraday_pool["IntradayCurrentVolumeScore"] = _intraday_rank100(
            intraday_pool["DayVolume"]
        ).clip(0, 100)
        intraday_pool["IntradayTurnoverScore"] = _intraday_rank100(
            intraday_pool["TodayTurnoverCr"]
        ).clip(0, 100)

        intraday_pool["IntradayTradabilityScore"] = (
            0.40 * intraday_pool["IntradayCurrentVolumeScore"] +
            0.35 * intraday_pool["IntradayTurnoverScore"] +
            0.25 * intraday_pool["IntradayVolumePaceScore"]
        ).clip(0, 100)

        intraday_pool["IntradayTradability"] = pd.cut(
            intraday_pool["IntradayTradabilityScore"],
            bins=[-np.inf, 35, 50, 65, 80, np.inf],
            labels=["WEAK", "FAIR", "GOOD", "STRONG", "EXCELLENT"],
            right=False
        ).astype(str)

        if "ETF_RRG_Score" not in intraday_pool.columns:
            intraday_pool["ETF_RRG_Score"] = np.nan

        intraday_pool["IntradayETFRRGScore"] = (
            pd.to_numeric(
                intraday_pool["ETF_RRG_Score"], errors="coerce"
            )
            .fillna(
                pd.Series(
                    np.where(
                        intraday_pool["Quadrant"].eq("LEADING"),
                        92.0, 78.0
                    ),
                    index=intraday_pool.index
                )
            )
            .clip(0, 100)
        )

        intraday_pool["IntradayExtensionScore"] = np.select(
            [
                intraday_pool["IntradayMoveATR"].fillna(99) <= 0.25,
                intraday_pool["IntradayMoveATR"].fillna(99) <= 0.75,
                intraday_pool["IntradayMoveATR"].fillna(99) <= 1.10,
                intraday_pool["IntradayMoveATR"].fillna(99) <= 1.35,
            ],
            [70, 100, 85, 55],
            default=20
        ).astype(float)

        intraday_pool["IntradayBuyScore"] = (
            0.20 * intraday_pool["IntradayTodayStrengthScore"] +
            0.20 * intraday_pool["IntradayVolumePaceScore"] +
            0.15 * intraday_pool["IntradayNearHighScore"] +
            0.15 * intraday_pool["IntradayTradabilityScore"] +
            0.10 * intraday_pool["IntradayOpenDriveScore"] +
            0.10 * intraday_pool["IntradayGroupStrengthScore"] +
            0.05 * intraday_pool["IntradayETFRRGScore"] +
            0.05 * intraday_pool["IntradayExtensionScore"]
        ).clip(0, 100)

        # Hard entry rules unchanged; only the intraday universe changed.
        intraday_pool["StrictIntradayQualified"] = (
            INTRADAY_ACTIONABLE_WINDOW &
            (
                pd.to_numeric(
                    intraday_pool["NSE_ChangePct"], errors="coerce"
                ).fillna(0) >= 0.25
            ) &
            (
                pd.to_numeric(
                    intraday_pool["IntradayOpenDrivePct"], errors="coerce"
                ).fillna(-99) >= 0.00
            ) &
            (
                pd.to_numeric(
                    intraday_pool["RangePositionPct"], errors="coerce"
                ).fillna(0) >= 70
            ) &
            (
                pd.to_numeric(
                    intraday_pool["VolumePaceVs30D"], errors="coerce"
                ).fillna(0) >= 0.75
            ) &
            (
                pd.to_numeric(
                    intraday_pool["IntradayTradabilityScore"], errors="coerce"
                ).fillna(0) >= 50
            ) &
            (
                pd.to_numeric(
                    intraday_pool["IntradayMoveATR"], errors="coerce"
                ).fillna(99) <= 1.35
            ) &
            (
                pd.to_numeric(
                    intraday_pool["IntradayBuyScore"], errors="coerce"
                ).fillna(0) >= 70
            )
        )

        def intraday_signal(row):
            score = row.get("IntradayBuyScore", np.nan)
            if not bool(row.get("StrictIntradayQualified", False)):
                return "WATCH ONLY"
            if pd.notna(score) and score >= 85:
                return "TOP INTRADAY SETUP"
            if pd.notna(score) and score >= 77:
                return "STRONG INTRADAY SETUP"
            return "GOOD INTRADAY SETUP"

        intraday_pool["IntradaySignal"] = intraday_pool.apply(
            intraday_signal, axis=1
        )

        def why_not_intraday(row):
            reasons = []
            day = row.get("NSE_ChangePct", np.nan)
            opendrive = row.get("IntradayOpenDrivePct", np.nan)
            near = row.get("RangePositionPct", np.nan)
            volpace = row.get("VolumePaceVs30D", np.nan)
            trad = row.get("IntradayTradabilityScore", np.nan)
            moveatr = row.get("IntradayMoveATR", np.nan)
            score = row.get("IntradayBuyScore", np.nan)

            if not INTRADAY_ACTIONABLE_WINDOW:
                reasons.append(
                    f"Scan timing not actionable ({INTRADAY_SCAN_WINDOW})"
                )
            if pd.isna(day) or day < 0.25:
                reasons.append(
                    f"Today {0 if pd.isna(day) else day:.2f}% < +0.25%"
                )
            if pd.isna(opendrive) or opendrive < 0:
                reasons.append(
                    f"Below open ({0 if pd.isna(opendrive) else opendrive:.2f}%)"
                )
            if pd.isna(near) or near < 70:
                reasons.append(
                    f"Near day high {0 if pd.isna(near) else near:.1f}% < 70%"
                )
            if pd.isna(volpace) or volpace < 0.75:
                reasons.append(
                    f"Volume pace {0 if pd.isna(volpace) else volpace:.2f}x < 0.75x"
                )
            if pd.isna(trad) or trad < 50:
                reasons.append(
                    f"Tradability {0 if pd.isna(trad) else trad:.1f} < 50"
                )
            if pd.isna(moveatr) or moveatr > 1.35:
                reasons.append(
                    f"Move {0 if pd.isna(moveatr) else moveatr:.2f} ATR > 1.35 ATR"
                )
            if pd.isna(score) or score < 70:
                reasons.append(
                    f"Intraday score {0 if pd.isna(score) else score:.1f} < 70"
                )
            return "; ".join(reasons) if reasons else "Qualified"

        intraday_pool["WhyNotIntradayQualified"] = intraday_pool.apply(
            why_not_intraday, axis=1
        )
        intraday_pool["IntradayFailedRuleCount"] = (
            intraday_pool["WhyNotIntradayQualified"].apply(
                lambda x: 0 if x == "Qualified"
                else len(str(x).split("; "))
            )
        )

        intraday_buys = (
            intraday_pool[intraday_pool["StrictIntradayQualified"]]
            .sort_values(
                [
                    "IntradayBuyScore", "IntradayGroupRank",
                    "IntradayTradabilityScore", "RangePositionPct"
                ],
                ascending=[False, True, False, False]
            )
            .head(10)
            .reset_index(drop=True)
        )
        if not intraday_buys.empty:
            intraday_buys.insert(
                0, "IntradayRank",
                np.arange(1, len(intraday_buys) + 1)
            )
        else:
            intraday_buys.insert(
                0, "IntradayRank", pd.Series(dtype=int)
            )

        intraday_watchlist = (
            intraday_pool[
                ~intraday_pool["StrictIntradayQualified"]
            ]
            .sort_values(
                [
                    "IntradayFailedRuleCount", "IntradayBuyScore",
                    "IntradayGroupRank", "IntradayTradabilityScore"
                ],
                ascending=[True, False, True, False]
            )
            .reset_index(drop=True)
        )
        if not intraday_watchlist.empty:
            intraday_watchlist.insert(
                0, "IntradayWatchRank",
                np.arange(1, len(intraday_watchlist) + 1)
            )
    else:
        intraday_buys = intraday_pool.copy()
        intraday_buys["IntradayRank"] = pd.Series(dtype=int)
        intraday_watchlist = intraday_pool.copy()
        intraday_watchlist["IntradayWatchRank"] = pd.Series(dtype=int)
        intraday_watchlist["WhyNotIntradayQualified"] = pd.Series(dtype=str)
        intraday_watchlist["IntradayFailedRuleCount"] = pd.Series(dtype=int)


    # ------------------------------------------------------------
    # V4.6.0 — NEGATIVE / DIP ETF DYNAMIC LIQUIDITY UNIVERSE
    # ------------------------------------------------------------
    # IMPORTANT:
    #   POSITIVE SWING = UNCHANGED
    #   POSITIVE INTRADAY = UNCHANGED
    #
    # This block replaces ONLY the negative/dip universe used later for
    # all_latest_losers / dip_candidates.
    #
    # MARKET HOURS:
    #   Top 75 active negative Equity + International Equity ETFs
    #   Negative Activity Score =
    #       40% Current NSE Turnover
    #       25% Current NSE Volume
    #       25% Time-adjusted Volume Pace vs prior 30D normal
    #       10% 30D Average Turnover
    #
    # COMPLETED SESSION:
    #   No Top-75 cap.
    #   Liquidity gate =
    #       30D Average Turnover >= ₹1.00 Cr
    #       Current-session Turnover >= ₹0.50 Cr
    #
    # Gold / Silver / Debt / Liquid are excluded.
    # The existing dip opportunity score/signal philosophy is retained for now.
    # We are changing the NEGATIVE LIQUIDITY UNIVERSE first, then we can separately
    # review the rebound/dip ranking logic after this live test.

    NEG_INTRADAY_TOP_N = 75
    NEG_INTRADAY_SEED_N = 140

    # Negative intraday candidates are intended for a possible rebound / next-day
    # swing entry, so they must be normally liquid in rupee terms as well.
    NEG_INTRADAY_MIN_AVG30_TURNOVER_CR = 1.00
    NEG_INTRADAY_MIN_CURRENT_TURNOVER_FLOOR_CR = 0.25
    NEG_INTRADAY_TARGET_FULLDAY_TURNOVER_CR = 1.00

    # Time-adjusted current-turnover guardrail:
    # e.g. early session -> at least ₹0.25 Cr;
    # later in the day -> rises gradually toward roughly ₹1 Cr.
    NEG_INTRADAY_MIN_TODAY_TURNOVER_CR = (
        max(
            NEG_INTRADAY_MIN_CURRENT_TURNOVER_FLOOR_CR,
            NEG_INTRADAY_TARGET_FULLDAY_TURNOVER_CR * MARKET_PROGRESS_FRACTION
        )
        if CURRENT_SESSION_MODE == "INTRADAY"
        else NEG_INTRADAY_MIN_CURRENT_TURNOVER_FLOOR_CR
    )

    NEG_EOD_MIN_AVG30_TURNOVER_CR = 1.00
    NEG_EOD_MIN_TODAY_TURNOVER_CR = 0.50

    def _neg_rank100(s):
        s = pd.to_numeric(s, errors="coerce")
        if s.notna().sum() <= 1:
            return pd.Series(50.0, index=s.index)
        return s.rank(pct=True, method="average").fillna(0.5) * 100.0

    def _neg_hist(ticker, base_df, intraday_extra_df, negative_extra_df):
        if ticker in base_df.columns:
            return base_df[ticker].dropna().copy()
        if ticker in intraday_extra_df.columns:
            return intraday_extra_df[ticker].dropna().copy()
        if ticker in negative_extra_df.columns:
            return negative_extra_df[ticker].dropna().copy()
        return pd.Series(dtype=float)

    # All currently negative Equity / International Equity ETFs from the NSE master.
    negative_master = master_all[
        master_all["AssetClass"].isin(["EQUITY", "INTERNATIONAL"]) &
        (pd.to_numeric(master_all["NSE_ChangePct"], errors="coerce").fillna(0) < 0)
    ].copy()

    negative_master["NegCurrentVolume"] = pd.to_numeric(
        negative_master["NSE_CurrentVolume"], errors="coerce"
    ).fillna(0)

    negative_master["NegTodayTurnoverCr"] = pd.to_numeric(
        negative_master["NSE_CurrentValueCr"], errors="coerce"
    )
    negative_master["NegTodayTurnoverCr"] = negative_master["NegTodayTurnoverCr"].fillna(
        (
            pd.to_numeric(negative_master["NSE_LTP"], errors="coerce").fillna(0) *
            negative_master["NegCurrentVolume"]
        ) / 1e7
    )

    negative_master = negative_master[
        (negative_master["NegCurrentVolume"] > 0) &
        (pd.to_numeric(negative_master["NSE_LTP"], errors="coerce").fillna(0) > 0)
    ].copy()

    # Broad prefilter only to limit extra Yahoo downloads.
    # The final market-hours universe is re-ranked using full 40/25/25/10 scoring.
    if not negative_master.empty:
        negative_master["NegSeedTurnoverScore"] = _neg_rank100(
            negative_master["NegTodayTurnoverCr"]
        )
        negative_master["NegSeedVolumeScore"] = _neg_rank100(
            negative_master["NegCurrentVolume"]
        )
        negative_master["NegSeedScore"] = (
            0.60 * negative_master["NegSeedTurnoverScore"] +
            0.40 * negative_master["NegSeedVolumeScore"]
        )

        if CURRENT_SESSION_MODE == "INTRADAY":
            negative_seed = (
                negative_master
                .sort_values(
                    ["NegSeedScore", "NegTodayTurnoverCr", "NegCurrentVolume"],
                    ascending=[False, False, False]
                )
                .head(NEG_INTRADAY_SEED_N)
                .copy()
            )
            NEGATIVE_UNIVERSE_MODE = (
                "INTRADAY TOP 75 ACTIVE NEGATIVE EQUITY / INTERNATIONAL "
                f"| Avg30 Turnover >= ₹{NEG_INTRADAY_MIN_AVG30_TURNOVER_CR:.2f} Cr "
                f"| Current Turnover >= ₹{NEG_INTRADAY_MIN_TODAY_TURNOVER_CR:.2f} Cr"
            )
        else:
            # For completed-session mode, today turnover >= ₹0.50 Cr is already a
            # required gate, so there is no reason to download histories below it.
            negative_seed = negative_master[
                negative_master["NegTodayTurnoverCr"] >= NEG_EOD_MIN_TODAY_TURNOVER_CR
            ].copy()
            NEGATIVE_UNIVERSE_MODE = (
                "COMPLETED SESSION — 30D AVG TURNOVER >= ₹1.00 Cr"
            )
    else:
        negative_seed = negative_master.copy()
        NEGATIVE_UNIVERSE_MODE = "NO NEGATIVE EQUITY ETF"

    # Download only missing histories for the negative-only supplemental universe.
    neg_extra_prices = pd.DataFrame()
    neg_extra_volumes = pd.DataFrame()
    neg_extra_highs = pd.DataFrame()
    neg_extra_lows = pd.DataFrame()
    neg_extra_failed = []

    if not negative_seed.empty:
        _base_existing = set(prices.columns.astype(str))
        _intraday_existing = set(extra_prices.columns.astype(str)) if 'extra_prices' in locals() else set()

        neg_extra_tickers = [
            t for t in negative_seed["YFTicker"].dropna().astype(str).tolist()
            if t not in _base_existing and t not in _intraday_existing
        ]

        if neg_extra_tickers:
            print(
                f"   Negative liquidity universe: extra histories needed = "
                f"{len(neg_extra_tickers)}"
            )
            (
                neg_extra_prices,
                neg_extra_volumes,
                neg_extra_highs,
                neg_extra_lows,
                neg_extra_failed
            ) = yf_batch_download(neg_extra_tickers)

    # Existing 1-lakh analysis rows are reused without changing them.
    _neg_analysis_by_symbol = {
        str(r["Symbol"]): r.to_dict()
        for _, r in analysis.iterrows()
    }

    # Theme structural RRG lookup.
    _theme_rrg_map = {}
    if not theme_rrg.empty and "Theme" in theme_rrg.columns:
        for _, _tr in theme_rrg.iterrows():
            _theme_rrg_map[str(_tr["Theme"])] = {
                "Theme_Quadrant": _tr.get("Quadrant", "NO DATA"),
                "Theme_RS_Ratio": _tr.get("RS_Ratio", np.nan),
                "Theme_RS_Momentum": _tr.get("RS_Momentum", np.nan),
            }

    negative_rows = []

    for _, mr in negative_seed.iterrows():
        sym = str(mr["Symbol"])
        tk = str(mr["YFTicker"])

        # ---------------- Existing >=1 lakh analytics ----------------
        if sym in _neg_analysis_by_symbol:
            d = _neg_analysis_by_symbol[sym].copy()

            # Use the CURRENT NSE snapshot values.
            d["DayVolume"] = pd.to_numeric(
                pd.Series([mr.get("NSE_CurrentVolume", np.nan)]), errors="coerce"
            ).iloc[0]
            d["TodayTurnoverCr"] = pd.to_numeric(
                pd.Series([mr.get("NSE_CurrentValueCr", np.nan)]), errors="coerce"
            ).iloc[0]
            if pd.isna(d["TodayTurnoverCr"]):
                _ltp = pd.to_numeric(
                    pd.Series([mr.get("NSE_LTP", np.nan)]), errors="coerce"
                ).iloc[0]
                if pd.notna(_ltp) and pd.notna(d["DayVolume"]):
                    d["TodayTurnoverCr"] = _ltp * d["DayVolume"] / 1e7

            # Prior completed sessions are preferred for normal-volume comparison.
            _prior30 = pd.to_numeric(
                pd.Series([d.get("Prior30AvgVolume", np.nan)]), errors="coerce"
            ).iloc[0]
            if pd.isna(_prior30):
                _prior30 = pd.to_numeric(
                    pd.Series([d.get("Avg30Volume", np.nan)]), errors="coerce"
                ).iloc[0]
            d["Prior30AvgVolume"] = _prior30

            _ltp = pd.to_numeric(
                pd.Series([d.get("LTP", np.nan)]), errors="coerce"
            ).iloc[0]
            d["Avg30TurnoverCr"] = (
                _ltp * _prior30 / 1e7
                if pd.notna(_ltp) and pd.notna(_prior30)
                else d.get("Avg30TurnoverCr", np.nan)
            )

            negative_rows.append(d)
            continue

        # ---------------- Supplemental negative-only analytics ----------------
        p = _neg_hist(
            tk,
            prices,
            extra_prices if 'extra_prices' in locals() else pd.DataFrame(),
            neg_extra_prices
        )
        v = _neg_hist(
            tk,
            volumes,
            extra_volumes if 'extra_volumes' in locals() else pd.DataFrame(),
            neg_extra_volumes
        )
        h = _neg_hist(
            tk,
            highs,
            extra_highs if 'extra_highs' in locals() else pd.DataFrame(),
            neg_extra_highs
        )
        l = _neg_hist(
            tk,
            lows,
            extra_lows if 'extra_lows' in locals() else pd.DataFrame(),
            neg_extra_lows
        )

        if len(p) < MIN_RRG_OBS:
            continue

        try:
            p.index = pd.to_datetime(p.index).tz_localize(None)
            v.index = pd.to_datetime(v.index).tz_localize(None)
            h.index = pd.to_datetime(h.index).tz_localize(None)
            l.index = pd.to_datetime(l.index).tz_localize(None)
        except Exception:
            p.index = pd.to_datetime(p.index)
            v.index = pd.to_datetime(v.index)
            h.index = pd.to_datetime(h.index)
            l.index = pd.to_datetime(l.index)

        ltp = pd.to_numeric(
            pd.Series([mr.get("NSE_LTP", np.nan)]), errors="coerce"
        ).iloc[0]
        prev = pd.to_numeric(
            pd.Series([mr.get("NSE_PrevClose", np.nan)]), errors="coerce"
        ).iloc[0]
        opn = pd.to_numeric(
            pd.Series([mr.get("NSE_Open", np.nan)]), errors="coerce"
        ).iloc[0]
        hi = pd.to_numeric(
            pd.Series([mr.get("NSE_High", np.nan)]), errors="coerce"
        ).iloc[0]
        lo = pd.to_numeric(
            pd.Series([mr.get("NSE_Low", np.nan)]), errors="coerce"
        ).iloc[0]
        chg = pd.to_numeric(
            pd.Series([mr.get("NSE_ChangePct", np.nan)]), errors="coerce"
        ).iloc[0]
        dayvol = pd.to_numeric(
            pd.Series([mr.get("NSE_CurrentVolume", np.nan)]), errors="coerce"
        ).iloc[0]
        turnover = pd.to_numeric(
            pd.Series([mr.get("NSE_CurrentValueCr", np.nan)]), errors="coerce"
        ).iloc[0]
        if pd.isna(turnover) and pd.notna(ltp) and pd.notna(dayvol):
            turnover = ltp * dayvol / 1e7

        # Prior 30 COMPLETED sessions.
        v_completed = v[v.index < current_price_date]
        prior30 = v_completed.tail(30).mean() if len(v_completed) else np.nan
        avg30_turnover = (
            ltp * prior30 / 1e7
            if pd.notna(ltp) and pd.notna(prior30)
            else np.nan
        )

        range_pos = (
            (ltp - lo) / (hi - lo) * 100.0
            if pd.notna(ltp) and pd.notna(hi) and pd.notna(lo) and hi > lo
            else np.nan
        )
        if pd.notna(range_pos):
            range_pos = float(np.clip(range_pos, 0, 100))

        # Historical returns / trend context for display.
        p_completed = p[p.index < current_price_date]
        dma20_val = dma(p_completed, 20)
        dma50_val = dma(p_completed, 50)
        dma100_val = dma(p_completed, 100)
        dma200_val = dma(p_completed, 200)

        trend_points = int(
            (pd.notna(ltp) and pd.notna(dma20_val) and ltp > dma20_val) +
            (pd.notna(ltp) and pd.notna(dma50_val) and ltp > dma50_val) +
            (pd.notna(ltp) and pd.notna(dma100_val) and ltp > dma100_val) +
            (pd.notna(ltp) and pd.notna(dma200_val) and ltp > dma200_val)
        )

        atr20_val = atr_from_history(
            p_completed,
            h.reindex(p.index).loc[p.index < current_price_date],
            l.reindex(p.index).loc[p.index < current_price_date],
            20
        )

        # ETF structural RRG.
        p_rrg = p.copy()
        if pd.notna(ltp):
            p_rrg.loc[current_price_date] = float(ltp)
            p_rrg = p_rrg.sort_index()

        rr = rrg_series(p_rrg, bench)
        if rr.empty:
            etf_quad = "NO DATA"
        else:
            etf_quad = quadrant(
                rr.iloc[-1]["RS_Ratio"],
                rr.iloc[-1]["RS_Momentum"]
            )

        theme_meta = _theme_rrg_map.get(str(mr["Theme"]), {})

        negative_rows.append({
            "Symbol": mr["Symbol"],
            "Underlying": mr["Underlying"],
            "SecurityName": mr["SecurityName"],
            "AssetClass": mr["AssetClass"],
            "Theme": mr["Theme"],
            "YFTicker": tk,
            "NSE_DownloadTimestamp": mr.get(
                "NSE_DownloadTimestamp", NSE_DOWNLOAD_TIMESTAMP
            ),
            "NSE_AsOf": mr.get("NSE_AsOf", ""),
            "NSE_PrevClose": prev,
            "NSE_Open": opn,
            "NSE_High": hi,
            "NSE_Low": lo,
            "NSE_ChangePct": chg,
            "LTP": ltp,
            "DayVolume": dayvol,
            "TodayTurnoverCr": turnover,
            "Avg30Volume": prior30,
            "Prior30AvgVolume": prior30,
            "Avg30TurnoverCr": avg30_turnover,
            "DayReturnPct": chg,
            "Week1ReturnPct": ret_n(p, 5),
            "Month1ReturnPct": ret_n(p, 21),
            "Month3ReturnPct": ret_n(p, 63),
            "DMA20": dma20_val,
            "DMA50": dma50_val,
            "DMA100": dma100_val,
            "DMA200": dma200_val,
            "TrendPoints": trend_points,
            "TrendScore": trend_points / 4.0 * 100.0,
            "ATR20": atr20_val,
            "RangePositionPct": range_pos,
            "Quadrant": etf_quad,
            "Theme_Quadrant": theme_meta.get("Theme_Quadrant", "NO DATA"),
            "Theme_RS_Ratio": theme_meta.get("Theme_RS_Ratio", np.nan),
            "Theme_RS_Momentum": theme_meta.get("Theme_RS_Momentum", np.nan),
            "LatestDataMode": CURRENT_SESSION_MODE,
            "NegativeSupplemental": True,
        })

    negative_base = pd.DataFrame(negative_rows)

    if not negative_base.empty:
        for c in [
            "NSE_ChangePct", "DayVolume", "TodayTurnoverCr",
            "Prior30AvgVolume", "Avg30TurnoverCr", "RangePositionPct"
        ]:
            if c not in negative_base.columns:
                negative_base[c] = np.nan
            negative_base[c] = pd.to_numeric(negative_base[c], errors="coerce")

        # Time-adjusted pace for MARKET HOURS; full-day multiple after close.
        negative_base["NegativeVolumePace"] = np.where(
            negative_base["Prior30AvgVolume"].fillna(0) > 0,
            negative_base["DayVolume"] /
            (
                negative_base["Prior30AvgVolume"] *
                (MARKET_PROGRESS_FRACTION if CURRENT_SESSION_MODE == "INTRADAY" else 1.0)
            ),
            np.nan
        )

        # Keep legacy-compatible names for downstream dip display/chart code.
        negative_base["CurrentVolumeMultiple30D"] = negative_base["NegativeVolumePace"]
        negative_base["VolumeMultiple"] = negative_base["NegativeVolumePace"]
        negative_base["CurrentVsAvg30VolumePct"] = (
            negative_base["NegativeVolumePace"] - 1.0
        ) * 100.0

        # ---------------- MARKET HOURS: TOP 75 ACTIVE NEGATIVE ETFs ----------------
        if CURRENT_SESSION_MODE == "INTRADAY":
            # Hard liquidity guardrails first. This prevents a tiny/illiquid ETF
            # from appearing as an "A+ dip" merely because its volume is many times
            # its own very small normal volume.
            negative_base = negative_base[
                (negative_base["Avg30TurnoverCr"] >= NEG_INTRADAY_MIN_AVG30_TURNOVER_CR) &
                (negative_base["TodayTurnoverCr"] >= NEG_INTRADAY_MIN_TODAY_TURNOVER_CR)
            ].copy()

            negative_base["NegTurnoverScore"] = _neg_rank100(
                negative_base["TodayTurnoverCr"]
            )
            negative_base["NegCurrentVolumeScore"] = _neg_rank100(
                negative_base["DayVolume"]
            )
            negative_base["NegVolumePaceScore"] = _neg_rank100(
                negative_base["NegativeVolumePace"]
            )
            negative_base["NegNormalLiquidityScore"] = _neg_rank100(
                negative_base["Avg30TurnoverCr"]
            )

            negative_base["NegativeActivityScore"] = (
                0.40 * negative_base["NegTurnoverScore"] +
                0.25 * negative_base["NegCurrentVolumeScore"] +
                0.25 * negative_base["NegVolumePaceScore"] +
                0.10 * negative_base["NegNormalLiquidityScore"]
            ).clip(0, 100)

            negative_universe = (
                negative_base
                .sort_values(
                    [
                        "NegativeActivityScore",
                        "TodayTurnoverCr",
                        "DayVolume"
                    ],
                    ascending=[False, False, False]
                )
                .head(NEG_INTRADAY_TOP_N)
                .copy()
            )

            negative_universe["NegativeLiquidityScore"] = (
                negative_universe["NegativeActivityScore"]
            )
            negative_universe["DipEligibilityResult"] = "INTRADAY TOP 75 ACTIVE"

        # ---------------- COMPLETED SESSION: HISTORICAL LIQUIDITY GATE ----------------
        else:
            negative_universe = negative_base[
                (negative_base["Avg30TurnoverCr"] >= NEG_EOD_MIN_AVG30_TURNOVER_CR) &
                (negative_base["TodayTurnoverCr"] >= NEG_EOD_MIN_TODAY_TURNOVER_CR)
            ].copy()

            if not negative_universe.empty:
                negative_universe["NegAvg30TurnoverScore"] = _neg_rank100(
                    negative_universe["Avg30TurnoverCr"]
                )
                negative_universe["NegTodayTurnoverScore"] = _neg_rank100(
                    negative_universe["TodayTurnoverCr"]
                )
                negative_universe["NegVolumePaceScore"] = _neg_rank100(
                    negative_universe["NegativeVolumePace"]
                )

                negative_universe["NegativeLiquidityScore"] = (
                    0.55 * negative_universe["NegAvg30TurnoverScore"] +
                    0.25 * negative_universe["NegTodayTurnoverScore"] +
                    0.20 * negative_universe["NegVolumePaceScore"]
                ).clip(0, 100)

            negative_universe["NegativeActivityScore"] = (
                negative_universe["NegativeLiquidityScore"]
                if "NegativeLiquidityScore" in negative_universe.columns
                else np.nan
            )
            negative_universe["DipEligibilityResult"] = "EOD LIQUIDITY QUALIFIED"

        if not negative_universe.empty:
            # ============================================================
            # NEGATIVE REBOUND / NEXT-DAY OPPORTUNITY ENGINE
            # ============================================================
            # Objective:
            #   Rank a liquid ETF that is DOWN today but is showing evidence
            #   that the sell-off may be stabilising / reversing, so Rank 1 is
            #   the best current rebound candidate rather than the biggest loser.
            #
            # Rebound Quality Score = 100%
            #   20% Recovery from today's low
            #   15% Price recovery versus today's open
            #   10% Theme / Group RRG context
            #   10% ETF RRG context
            #   10% Trend strength (20/50/100/200 DMA)
            #   10% Recent momentum context
            #   10% Sell-off volume character / absorption
            #    5% Liquidity quality
            #    5% Dip magnitude quality
            #    5% ATR pullback quality
            #
            # The liquidity universe was already filtered before this point.
            # This engine deliberately penalises ETFs still sitting near the
            # session low or structurally deteriorating.

            # ---------- participation / sell-off character ----------
            def _negative_volume_character(row):
                volx = row.get("NegativeVolumePace", np.nan)
                rp = row.get("RangePositionPct", np.nan)

                if pd.isna(volx):
                    return "VOLUME DATA N/A"

                # High relative volume + strong recovery from low = possible
                # absorption rather than uncontrolled selling.
                if volx >= 1.30 and pd.notna(rp) and rp >= 65:
                    return "HIGH VOLUME ABSORPTION"

                # Very high relative volume while still trapped near the low is a
                # warning, not an automatic buy.
                if volx >= 2.00 and pd.notna(rp) and rp < 25:
                    return "EXTREME SELLOFF"

                if volx >= 1.30 and pd.notna(rp) and rp < 45:
                    return "HIGH VOLUME SELLOFF"

                if volx <= 0.80:
                    return "LIGHT VOLUME DIP"

                return "NORMAL VOLUME DIP"

            negative_universe["DipVolumeCharacter"] = negative_universe.apply(
                _negative_volume_character, axis=1
            )

            negative_universe["AbsPctFall"] = (
                pd.to_numeric(
                    negative_universe["NSE_ChangePct"], errors="coerce"
                ).abs()
            )

            # ---------- 1) Recovery from today's low ----------
            def _recovery_score(x):
                if pd.isna(x):
                    return 15.0
                if x >= 75:
                    return 100.0
                if x >= 60:
                    return 90.0
                if x >= 45:
                    return 80.0
                if x >= 30:
                    return 62.0
                if x >= 15:
                    return 38.0
                return 12.0

            negative_universe["RecoveryFromLowScore"] = (
                pd.to_numeric(
                    negative_universe["RangePositionPct"], errors="coerce"
                ).apply(_recovery_score)
            )

            # ---------- 2) Price recovery versus today's open ----------
            negative_universe["ReboundVsOpenPct"] = np.where(
                pd.to_numeric(
                    negative_universe["NSE_Open"], errors="coerce"
                ).fillna(0) > 0,
                (
                    pd.to_numeric(
                        negative_universe["LTP"], errors="coerce"
                    ) /
                    pd.to_numeric(
                        negative_universe["NSE_Open"], errors="coerce"
                    ) - 1.0
                ) * 100.0,
                np.nan
            )

            def _vs_open_score(x):
                if pd.isna(x):
                    return 20.0
                if x >= 0.50:
                    return 100.0
                if x >= 0.00:
                    return 90.0
                if x >= -0.25:
                    return 75.0
                if x >= -0.60:
                    return 55.0
                if x >= -1.00:
                    return 35.0
                return 15.0

            negative_universe["ReboundVsOpenScore"] = (
                negative_universe["ReboundVsOpenPct"].apply(_vs_open_score)
            )

            # ---------- 3 & 4) Structural RRG context ----------
            def _rrg_rebound_score(q):
                q = str(q).upper()
                if q == "LEADING":
                    return 100.0
                if q == "IMPROVING":
                    return 85.0
                if q == "WEAKENING":
                    return 55.0
                if q == "LAGGING":
                    return 20.0
                return 40.0

            if "Theme_Quadrant" not in negative_universe.columns:
                negative_universe["Theme_Quadrant"] = "NO DATA"
            if "Quadrant" not in negative_universe.columns:
                negative_universe["Quadrant"] = "NO DATA"

            negative_universe["ReboundThemeRRGScore"] = (
                negative_universe["Theme_Quadrant"].apply(_rrg_rebound_score)
            )
            negative_universe["ReboundETFRRGScore"] = (
                negative_universe["Quadrant"].apply(_rrg_rebound_score)
            )

            # ---------- 5) Trend strength ----------
            if "TrendPoints" not in negative_universe.columns:
                negative_universe["TrendPoints"] = np.nan

            def _trend_rebound_score(x):
                if pd.isna(x):
                    return 35.0
                x = int(x)
                return {
                    4: 100.0,
                    3: 85.0,
                    2: 60.0,
                    1: 30.0,
                    0: 10.0,
                }.get(x, 35.0)

            negative_universe["ReboundTrendScore"] = (
                pd.to_numeric(
                    negative_universe["TrendPoints"], errors="coerce"
                ).apply(_trend_rebound_score)
            )

            # ---------- 6) Recent momentum context ----------
            for _c in ["Week1ReturnPct", "Month1ReturnPct", "Month3ReturnPct"]:
                if _c not in negative_universe.columns:
                    negative_universe[_c] = np.nan

            def _recent_momentum_score(row):
                w = row.get("Week1ReturnPct", np.nan)
                m1 = row.get("Month1ReturnPct", np.nan)
                m3 = row.get("Month3ReturnPct", np.nan)

                if pd.isna(m1) and pd.isna(m3):
                    return 40.0

                # Best case: today's red session is a pullback inside a still
                # constructive 1M/3M trend.
                if pd.notna(m1) and pd.notna(m3) and m1 > 0 and m3 > 0:
                    if pd.notna(w) and w >= 0:
                        return 100.0
                    if pd.notna(w) and w >= -1.50:
                        return 90.0
                    return 80.0

                if pd.notna(m1) and m1 > 0:
                    return 75.0

                if (
                    pd.notna(m3) and m3 > 0 and
                    pd.notna(m1) and m1 >= -2.0
                ):
                    return 65.0

                if pd.notna(m3) and m3 > 0:
                    return 50.0

                return 25.0

            negative_universe["ReboundMomentumScore"] = negative_universe.apply(
                _recent_momentum_score, axis=1
            )

            # ---------- 7) Volume / absorption character ----------
            _volume_character_score = {
                "HIGH VOLUME ABSORPTION": 100.0,
                "NORMAL VOLUME DIP": 80.0,
                "LIGHT VOLUME DIP": 68.0,
                "HIGH VOLUME SELLOFF": 42.0,
                "EXTREME SELLOFF": 18.0,
                "VOLUME DATA N/A": 35.0,
            }
            negative_universe["ReboundVolumeCharacterScore"] = (
                negative_universe["DipVolumeCharacter"]
                .map(_volume_character_score)
                .fillna(40.0)
            )

            # ---------- 8) Liquidity quality ----------
            negative_universe["ReboundLiquidityScore"] = (
                pd.to_numeric(
                    negative_universe["NegativeLiquidityScore"], errors="coerce"
                ).fillna(50.0).clip(0, 100)
            )

            # ---------- 9) Dip magnitude quality ----------
            # We want a meaningful discount, but do not reward an uncontrolled
            # collapse simply because the percentage loss is larger.
            def _dip_magnitude_quality(x):
                if pd.isna(x):
                    return 25.0
                if x < 0.25:
                    return 25.0
                if x < 0.50:
                    return 50.0
                if x < 1.00:
                    return 85.0
                if x <= 2.00:
                    return 100.0
                if x <= 3.00:
                    return 80.0
                if x <= 3.50:
                    return 55.0
                return 25.0

            negative_universe["ReboundDipMagnitudeScore"] = (
                negative_universe["AbsPctFall"].apply(_dip_magnitude_quality)
            )

            # ---------- 10) ATR pullback quality ----------
            if "ATR20" not in negative_universe.columns:
                negative_universe["ATR20"] = np.nan

            negative_universe["DayFallATR"] = np.where(
                pd.to_numeric(
                    negative_universe["ATR20"], errors="coerce"
                ).fillna(0) > 0,
                (
                    pd.to_numeric(
                        negative_universe["LTP"], errors="coerce"
                    ) -
                    pd.to_numeric(
                        negative_universe["NSE_PrevClose"], errors="coerce"
                    )
                ).abs() /
                pd.to_numeric(
                    negative_universe["ATR20"], errors="coerce"
                ),
                np.nan
            )

            def _atr_pullback_score(x):
                if pd.isna(x):
                    return 40.0
                if x < 0.25:
                    return 45.0
                if x <= 0.75:
                    return 90.0
                if x <= 1.25:
                    return 100.0
                if x <= 1.75:
                    return 75.0
                if x <= 2.25:
                    return 45.0
                return 18.0

            negative_universe["ReboundATRScore"] = (
                negative_universe["DayFallATR"].apply(_atr_pullback_score)
            )

            # ---------- Legacy sell-off intensity kept only as background ----------
            negative_universe["PctMoveRankScore"] = _neg_rank100(
                negative_universe["AbsPctFall"]
            )
            negative_universe["CurrentVolumeRankScore"] = (
                negative_universe["NegativeLiquidityScore"]
            )
            negative_universe["DailyLoserRankScore"] = (
                0.60 * negative_universe["PctMoveRankScore"] +
                0.40 * negative_universe["NegativeLiquidityScore"]
            ).clip(0, 100)

            def _neg_volume_quality(x):
                if pd.isna(x):
                    return 35.0
                if x >= 4.0:
                    return 100.0
                if x >= 3.0:
                    return 95.0
                if x >= 2.0:
                    return 88.0
                if x >= 1.3:
                    return 78.0
                if x >= 0.8:
                    return 62.0
                return 40.0

            negative_universe["FullDipMagnitudeScore"] = (
                negative_universe["AbsPctFall"].clip(0, 4.0) / 4.0 * 100.0
            )
            negative_universe["FullDipVolumeScore"] = (
                negative_universe["NegativeVolumePace"].apply(_neg_volume_quality)
            )
            negative_universe["FullDipTurnoverScore"] = (
                negative_universe["TodayTurnoverCr"]
                .fillna(0).clip(0, 2.0) / 2.0 * 100.0
            ).clip(0, 100)
            negative_universe["AggressiveDipScore"] = (
                0.45 * negative_universe["FullDipMagnitudeScore"] +
                0.45 * negative_universe["FullDipVolumeScore"] +
                0.10 * negative_universe["FullDipTurnoverScore"]
            ).clip(0, 100)

            # ---------- FINAL REBOUND QUALITY SCORE ----------
            negative_universe["ReboundScore"] = (
                0.20 * negative_universe["RecoveryFromLowScore"] +
                0.15 * negative_universe["ReboundVsOpenScore"] +
                0.10 * negative_universe["ReboundThemeRRGScore"] +
                0.10 * negative_universe["ReboundETFRRGScore"] +
                0.10 * negative_universe["ReboundTrendScore"] +
                0.10 * negative_universe["ReboundMomentumScore"] +
                0.10 * negative_universe["ReboundVolumeCharacterScore"] +
                0.05 * negative_universe["ReboundLiquidityScore"] +
                0.05 * negative_universe["ReboundDipMagnitudeScore"] +
                0.05 * negative_universe["ReboundATRScore"]
            ).clip(0, 100)

            # ---------- STRICT REBOUND QUALIFICATION ----------
            # Do not force a Buy Priority if the sell-off has not stabilised.
            negative_universe["ReboundQualified"] = (
                (negative_universe["AbsPctFall"] >= 0.30) &
                (negative_universe["AbsPctFall"] <= 3.50) &
                (
                    pd.to_numeric(
                        negative_universe["RangePositionPct"], errors="coerce"
                    ).fillna(-1) >= 45
                ) &
                (
                    pd.to_numeric(
                        negative_universe["ReboundVsOpenPct"], errors="coerce"
                    ).fillna(-99) >= -0.25
                ) &
                (
                    pd.to_numeric(
                        negative_universe["TrendPoints"], errors="coerce"
                    ).fillna(-1) >= 2
                ) &
                (
                    negative_universe["Theme_Quadrant"]
                    .astype(str).isin(["LEADING", "IMPROVING", "WEAKENING"])
                ) &
                (
                    negative_universe["Quadrant"]
                    .astype(str).isin(["LEADING", "IMPROVING", "WEAKENING"])
                ) &
                (
                    pd.to_numeric(
                        negative_universe["DayFallATR"], errors="coerce"
                    ).fillna(99) <= 1.75
                ) &
                (
                    pd.to_numeric(
                        negative_universe["ReboundScore"], errors="coerce"
                    ).fillna(0) >= 70
                )
            )

            def _why_not_rebound(row):
                reasons = []

                fall = row.get("AbsPctFall", np.nan)
                rp = row.get("RangePositionPct", np.nan)
                vo = row.get("ReboundVsOpenPct", np.nan)
                tr = row.get("TrendPoints", np.nan)
                tq = str(row.get("Theme_Quadrant", "NO DATA"))
                eq = str(row.get("Quadrant", "NO DATA"))
                atrx = row.get("DayFallATR", np.nan)
                score = row.get("ReboundScore", np.nan)

                if pd.isna(fall) or fall < 0.30:
                    reasons.append(
                        f"Dip {0 if pd.isna(fall) else fall:.2f}% < 0.30%"
                    )
                elif fall > 3.50:
                    reasons.append(f"Dip {fall:.2f}% > 3.50% breakdown guard")

                if pd.isna(rp) or rp < 45:
                    reasons.append(
                        f"Recovery from low {0 if pd.isna(rp) else rp:.1f}% < 45%"
                    )

                if pd.isna(vo) or vo < -0.25:
                    reasons.append(
                        f"Vs open {0 if pd.isna(vo) else vo:.2f}% < -0.25%"
                    )

                if pd.isna(tr) or tr < 2:
                    reasons.append(
                        f"Trend {0 if pd.isna(tr) else int(tr)}/4 < 2/4"
                    )

                if tq not in ["LEADING", "IMPROVING", "WEAKENING"]:
                    reasons.append(f"Group RRG {tq}")

                if eq not in ["LEADING", "IMPROVING", "WEAKENING"]:
                    reasons.append(f"ETF RRG {eq}")

                if pd.isna(atrx):
                    reasons.append("ATR pullback unavailable")
                elif atrx > 1.75:
                    reasons.append(f"Fall {atrx:.2f} ATR > 1.75 ATR")

                if pd.isna(score) or score < 70:
                    reasons.append(
                        f"Rebound score {0 if pd.isna(score) else score:.1f} < 70"
                    )

                return "; ".join(reasons) if reasons else "Qualified"

            negative_universe["WhyNotReboundQualified"] = negative_universe.apply(
                _why_not_rebound, axis=1
            )
            negative_universe["ReboundFailedRuleCount"] = (
                negative_universe["WhyNotReboundQualified"].apply(
                    lambda x: 0 if x == "Qualified"
                    else len(str(x).split("; "))
                )
            )

            def _rebound_signal(row):
                score = row.get("ReboundScore", np.nan)
                qualified = bool(row.get("ReboundQualified", False))
                rp = row.get("RangePositionPct", np.nan)
                volchar = str(row.get("DipVolumeCharacter", ""))

                if qualified:
                    if pd.notna(score) and score >= 85:
                        return "TOP REBOUND BUY SETUP"
                    if pd.notna(score) and score >= 77:
                        return "STRONG REBOUND BUY SETUP"
                    return "GOOD REBOUND BUY SETUP"

                if (
                    pd.notna(rp) and rp < 15 and
                    volchar in ["EXTREME SELLOFF", "HIGH VOLUME SELLOFF"]
                ):
                    return "FALLING KNIFE — WAIT"

                if pd.notna(score) and score >= 60:
                    return "REBOUND WATCH"

                return "WAIT / WEAK REBOUND"

            negative_universe["ReboundSignal"] = negative_universe.apply(
                _rebound_signal, axis=1
            )

            # Compatibility fields used by the existing workbook/chart code.
            negative_universe["DipBuyScore"] = negative_universe["ReboundScore"]
            negative_universe["DipSignal"] = negative_universe["ReboundSignal"]
            negative_universe["LiquidityStatus"] = negative_universe[
                "DipEligibilityResult"
            ]
            negative_universe["LiquidityTier"] = NEGATIVE_UNIVERSE_MODE
            negative_universe["LiquidEligible"] = True

            # ---------- STRICT BUY PRIORITY ----------
            strict_rebound = (
                negative_universe[negative_universe["ReboundQualified"]]
                .sort_values(
                    [
                        "ReboundScore",
                        "RecoveryFromLowScore",
                        "NegativeLiquidityScore"
                    ],
                    ascending=[False, False, False]
                )
                .reset_index(drop=True)
            )
            if not strict_rebound.empty:
                strict_rebound["DipBuyPriority"] = np.arange(
                    1, len(strict_rebound) + 1
                )
            else:
                strict_rebound["DipBuyPriority"] = pd.Series(dtype=int)

            # ---------- NEAR / WATCH ----------
            rebound_watch = (
                negative_universe[~negative_universe["ReboundQualified"]]
                .sort_values(
                    [
                        "ReboundFailedRuleCount",
                        "ReboundScore",
                        "RecoveryFromLowScore",
                        "NegativeLiquidityScore"
                    ],
                    ascending=[True, False, False, False]
                )
                .reset_index(drop=True)
            )
            if not rebound_watch.empty:
                rebound_watch["ReboundWatchRank"] = np.arange(
                    1, len(rebound_watch) + 1
                )
            else:
                rebound_watch["ReboundWatchRank"] = pd.Series(dtype=int)

            # One existing negative sheet/dashboard: strict candidates first,
            # followed by the closest watches. No new worksheet is added.
            _strict_top = strict_rebound.head(10).copy()
            _remaining = max(0, 10 - len(_strict_top))
            _watch_fill = rebound_watch.head(_remaining).copy()

            all_latest_losers = pd.concat(
                [_strict_top, _watch_fill],
                ignore_index=True,
                sort=False
            )

            if not all_latest_losers.empty:
                all_latest_losers.insert(
                    0, "DipRank",
                    np.arange(1, len(all_latest_losers) + 1)
                )
            else:
                all_latest_losers.insert(
                    0, "DipRank", pd.Series(dtype=int)
                )

            dip_candidates = strict_rebound.head(10).copy()

        else:
            all_latest_losers = negative_universe.copy()
            all_latest_losers.insert(
                0, "DipRank", pd.Series(dtype=int)
            )
            dip_candidates = negative_universe.copy()

    else:
        negative_universe = pd.DataFrame()
        all_latest_losers = pd.DataFrame(columns=[
            "DipRank", "Symbol", "Theme", "NSE_ChangePct",
            "DayVolume", "TodayTurnoverCr", "DipSignal"
        ])
        dip_candidates = all_latest_losers.copy()

    NEGATIVE_UNIVERSE_COUNT = len(negative_universe)

    # ------------------------------------------------------------
    # DISPLAY-ONLY DEEP FALL WATCH — GROUP 1 ONLY
    # ------------------------------------------------------------
    # This does NOT change Rebound Score, qualification or Buy Priority.
    # It simply makes important liquid Equity / International ETFs down >= 1%
    # visible even when they are not yet safe rebound buys.
    if not negative_universe.empty:
        deep_fall_watch = negative_universe[
            negative_universe["AssetClass"].isin(["EQUITY", "INTERNATIONAL"]) &
            (
                pd.to_numeric(
                    negative_universe["NSE_ChangePct"], errors="coerce"
                ).fillna(0) <= -1.00
            )
        ].copy()

        if not deep_fall_watch.empty:
            deep_fall_watch = (
                deep_fall_watch
                .sort_values(
                    ["AbsPctFall", "NegativeLiquidityScore", "ReboundScore"],
                    ascending=[False, False, False]
                )
                .reset_index(drop=True)
            )
            deep_fall_watch.insert(
                0, "DeepFallRank",
                np.arange(1, len(deep_fall_watch) + 1)
            )
    else:
        deep_fall_watch = pd.DataFrame()

    DEEP_FALL_COUNT = len(deep_fall_watch)

    print(
        f"   Negative universe mode : {NEGATIVE_UNIVERSE_MODE}"
    )
    print(
        f"   Negative ETFs eligible : {NEGATIVE_UNIVERSE_COUNT}"
    )
    print(
        f"   Deep Fall Watch >=1%   : {DEEP_FALL_COUNT} "
        f"(Equity / International only)"
    )
    if CURRENT_SESSION_MODE == "INTRADAY":
        print(
            f"   Negative liquidity guard: Avg30 turnover >= "
            f"₹{NEG_INTRADAY_MIN_AVG30_TURNOVER_CR:.2f} Cr | "
            f"Current turnover >= ₹{NEG_INTRADAY_MIN_TODAY_TURNOVER_CR:.2f} Cr"
        )

    # Classification quality control
    commodity_in_equity = equity[
        (
            equity["Symbol"].astype(str) + " " +
            equity["Underlying"].astype(str) + " " +
            equity["SecurityName"].astype(str)
        ).str.upper().str.contains("GOLD|SILVER|SILV", regex=True, na=False)
    ]
    if len(commodity_in_equity) > 0:
        print("WARNING: Commodity ETF classification issue detected:")
        print(commodity_in_equity[["Symbol","Underlying","SecurityName"]].to_string(index=False))
    else:
        print("   Classification QC      : PASS — no Gold/Silver ETF in Equity universe")

    trade_candidates = (
        equity[
            equity["Theme_Quadrant"].isin(["LEADING","IMPROVING"]) &
            equity["Quadrant"].isin(["LEADING","IMPROVING"]) &
            (equity["TrendPoints"] >= 3)
        ]
        .sort_values(["TradeScore","DayVolume","Avg30TurnoverCr"], ascending=[False,False,False])
        .reset_index(drop=True)
    )

    # -----------------------------
    # STEP 8 — Rotation changes
    # -----------------------------
    rotation_rows = []
    for theme, rr in theme_rrg_history.items():
        if len(rr) < 2:
            continue
        prevq = quadrant(rr.iloc[-2]["RS_Ratio"], rr.iloc[-2]["RS_Momentum"])
        nowq = quadrant(rr.iloc[-1]["RS_Ratio"], rr.iloc[-1]["RS_Momentum"])
        rotation_rows.append({
            "Theme": theme,
            "PreviousQuadrant": prevq,
            "CurrentQuadrant": nowq,
            "ChangedInLatestSession": prevq != nowq,
            "RS_Ratio": rr.iloc[-1]["RS_Ratio"],
            "RS_Momentum": rr.iloc[-1]["RS_Momentum"],
            "DeltaRatio1D": rr.iloc[-1]["RS_Ratio"] - rr.iloc[-2]["RS_Ratio"],
            "DeltaMomentum1D": rr.iloc[-1]["RS_Momentum"] - rr.iloc[-2]["RS_Momentum"],
        })
    rotation_changes = pd.DataFrame(rotation_rows).sort_values(
        ["ChangedInLatestSession","DeltaMomentum1D"], ascending=[False,False]
    ) if rotation_rows else pd.DataFrame()

    # -----------------------------
    # STEP 9 — Charts
    # -----------------------------
    print("\n[6/10] Creating colorful RRG charts ...")

    QUAD_COLORS = {
        "LEADING": "#22C55E",
        "IMPROVING": "#3B82F6",
        "WEAKENING": "#F59E0B",
        "LAGGING": "#EF4444",
        "NO DATA": "#94A3B8",
    }

    def save_plotly(fig, name):
        html_path = DASH_DIR / f"{name}.html"
        fig.write_html(str(html_path), include_plotlyjs="cdn")
        png_path = CHART_DIR / f"{name}.png"
        try:
            fig.write_image(str(png_path), width=1600, height=950, scale=1.5)
        except Exception:
            print(f"   PNG export skipped for {name}; interactive HTML chart saved.")
        return html_path, png_path

    def make_rrg_chart(df, title, label_col, histories, top_n=24):
        d = df.copy()
        if d.empty:
            return go.Figure()

        # Prefer points farthest from center + higher liquidity if available
        d["Distance"] = np.sqrt((d["RS_Ratio"]-100)**2 + (d["RS_Momentum"]-100)**2)
        d = d.sort_values("Distance", ascending=False).head(top_n)

        fig = go.Figure()

        # background quadrants
        xmin = min(98.0, d["RS_Ratio"].min()-0.5)
        xmax = max(102.0, d["RS_Ratio"].max()+0.5)
        ymin = min(98.0, d["RS_Momentum"].min()-0.5)
        ymax = max(102.0, d["RS_Momentum"].max()+0.5)

        fig.add_shape(type="rect", x0=100, x1=xmax, y0=100, y1=ymax, fillcolor="rgba(34,197,94,0.07)", line_width=0, layer="below")
        fig.add_shape(type="rect", x0=xmin, x1=100, y0=100, y1=ymax, fillcolor="rgba(59,130,246,0.07)", line_width=0, layer="below")
        fig.add_shape(type="rect", x0=100, x1=xmax, y0=ymin, y1=100, fillcolor="rgba(245,158,11,0.07)", line_width=0, layer="below")
        fig.add_shape(type="rect", x0=xmin, x1=100, y0=ymin, y1=100, fillcolor="rgba(239,68,68,0.07)", line_width=0, layer="below")

        for q, g in d.groupby("Quadrant"):
            fig.add_trace(go.Scatter(
                x=g["RS_Ratio"], y=g["RS_Momentum"],
                mode="markers+text",
                text=g[label_col],
                textposition="top center",
                marker=dict(size=13, color=QUAD_COLORS.get(q, "#94A3B8"), line=dict(width=1, color="white")),
                name=q,
                hovertemplate="<b>%{text}</b><br>RS-Ratio=%{x:.2f}<br>RS-Momentum=%{y:.2f}<extra></extra>"
            ))

        # trails
        for lbl in d[label_col]:
            rr = histories.get(lbl)
            if rr is None or rr.empty:
                continue
            trail = rr.tail(RRG_TRAIL_DAYS)
            fig.add_trace(go.Scatter(
                x=trail["RS_Ratio"], y=trail["RS_Momentum"],
                mode="lines",
                line=dict(width=1.5, color="rgba(100,116,139,0.55)"),
                showlegend=False,
                hoverinfo="skip"
            ))

        fig.add_vline(x=100, line_dash="dash", line_color="#64748B")
        fig.add_hline(y=100, line_dash="dash", line_color="#64748B")

        fig.add_annotation(x=xmax, y=ymax, text="LEADING", showarrow=False, xanchor="right", yanchor="top", font=dict(color="#16A34A", size=15))
        fig.add_annotation(x=xmin, y=ymax, text="IMPROVING", showarrow=False, xanchor="left", yanchor="top", font=dict(color="#2563EB", size=15))
        fig.add_annotation(x=xmax, y=ymin, text="WEAKENING", showarrow=False, xanchor="right", yanchor="bottom", font=dict(color="#D97706", size=15))
        fig.add_annotation(x=xmin, y=ymin, text="LAGGING", showarrow=False, xanchor="left", yanchor="bottom", font=dict(color="#DC2626", size=15))

        fig.update_layout(
            title=title,
            template="plotly_white",
            height=800,
            xaxis_title="RS-Ratio (vs NIFTY 50)",
            yaxis_title="RS-Momentum",
            legend=dict(orientation="h", y=1.04, x=0),
            margin=dict(l=60,r=40,t=90,b=60)
        )
        return fig

    # Theme RRG
    theme_chart_df = theme_rrg[theme_rrg["AssetClass"].isin(["EQUITY","INTERNATIONAL"])].copy()
    fig_theme = make_rrg_chart(
        theme_chart_df,
        f"Equity ETF Theme RRG vs NIFTY 50 — {RUN_DATE}",
        "Theme",
        theme_rrg_history,
        TOP_THEME_CHART
    )
    save_plotly(fig_theme, "01_EQUITY_THEME_RRG")

    # ETF RRG
    equity_rrg_chart = rrg_etf[rrg_etf["AssetClass"].isin(["EQUITY","INTERNATIONAL"])].copy()
    fig_etf = make_rrg_chart(
        equity_rrg_chart,
        f"Equity + International Equity ETF RRG vs NIFTY 50 — {RUN_DATE}",
        "Symbol",
        rrg_history,
        TOP_ETF_CHART
    )
    save_plotly(fig_etf, "02_EQUITY_ETF_RRG")

    # Non-equity RRG
    non_eq_rrg_chart = rrg_etf[rrg_etf["AssetClass"].isin(["GOLD","SILVER","DEBT","LIQUID"])].copy()
    fig_asset = make_rrg_chart(
        non_eq_rrg_chart,
        f"Gold / Silver / Debt / Liquid ETF RRG — {RUN_DATE}",
        "Symbol",
        rrg_history,
        30
    )
    save_plotly(fig_asset, "03_NON_EQUITY_ETF_RRG")

    # Performance heatmap of top themes
    theme_perf = best_by_theme[best_by_theme["AssetClass"].isin(["EQUITY","INTERNATIONAL"])].copy()
    theme_perf = theme_perf.sort_values("TradeScore", ascending=False).head(20)
    heat_cols = ["DayReturnPct","Week1ReturnPct","Month1ReturnPct","Month3ReturnPct","Month6ReturnPct","Year1ReturnPct"]
    if len(theme_perf):
        z = theme_perf[heat_cols].astype(float).values
        fig_heat = go.Figure(data=go.Heatmap(
            z=z,
            x=["1D","1W","1M","3M","6M","1Y"],
            y=theme_perf["Theme"],
            colorscale="RdYlGn",
            zmid=0,
            colorbar=dict(title="Return %"),
            hovertemplate="Theme=%{y}<br>Period=%{x}<br>Return=%{z:.2f}%<extra></extra>"
        ))
        fig_heat.update_layout(
            title=f"Theme Performance Heatmap — Best-Liquidity ETF Proxy — {RUN_DATE}",
            template="plotly_white",
            height=760,
            margin=dict(l=160,r=40,t=80,b=60)
        )
        save_plotly(fig_heat, "04_THEME_RETURN_HEATMAP")

    # Top candidates bar
    tc = trade_candidates.head(20).sort_values("TradeScore")
    if len(tc):
        fig_candidates = go.Figure(go.Bar(
            x=tc["TradeScore"],
            y=tc["Symbol"] + " | " + tc["Theme"],
            orientation="h",
            text=tc["TradeScore"].round(1),
            textposition="outside",
            customdata=np.stack([tc["Theme_Quadrant"], tc["Quadrant"], tc["Avg30TurnoverCr"]], axis=-1),
            hovertemplate="<b>%{y}</b><br>Trade Score=%{x:.1f}<br>Theme RRG=%{customdata[0]}<br>ETF RRG=%{customdata[1]}<br>30D Turnover=₹%{customdata[2]:.2f} Cr<extra></extra>"
        ))
        fig_candidates.update_layout(
            title=f"Top Group 1 Equity ETF Trade Candidates — {RUN_DATE}",
            xaxis_title="Data-Driven Trade Score (0-100)",
            template="plotly_white",
            height=760,
            margin=dict(l=220,r=60,t=80,b=60)
        )
        save_plotly(fig_candidates, "05_TOP_TRADE_CANDIDATES")

    # Dip candidates bar
    dc = dip_candidates.head(20).sort_values("DipBuyScore")
    if len(dc):
        fig_dips = go.Figure(go.Bar(
            x=dc["DipBuyScore"],
            y=dc["Symbol"] + " | " + dc["Theme"],
            orientation="h",
            text=dc["DipBuyScore"].round(1),
            textposition="outside",
            customdata=np.stack([
                dc["DayReturnPct"], dc["Theme_Quadrant"], dc["Quadrant"],
                dc["Avg30TurnoverCr"], dc["VolumeMultiple"]
            ], axis=-1),
            hovertemplate="<b>%{y}</b><br>Dip Score=%{x:.1f}<br>Today=%{customdata[0]:.2f}%<br>Theme RRG=%{customdata[1]}<br>ETF RRG=%{customdata[2]}<br>30D Liquidity=₹%{customdata[3]:.2f} Cr<br>Volume x=%{customdata[4]:.2f}<extra></extra>"
        ))
        fig_dips.update_layout(
            title=f"Top 10 Daily ETF Losers — {RUN_DATE}",
            xaxis_title="Dip Buy Score (0-100)",
            template="plotly_white",
            height=760,
            margin=dict(l=220,r=60,t=80,b=60)
        )
        save_plotly(fig_dips, "06_TODAYS_DIP_OPPORTUNITIES")

    # -----------------------------
    # STEP 10 — HTML Dashboard
    # -----------------------------
    print("\n[7/10] Building dashboard ...")

    def qcount(df, q):
        return int((df["Quadrant"] == q).sum()) if "Quadrant" in df else 0

    eq_theme = theme_rrg[theme_rrg["AssetClass"].isin(["EQUITY","INTERNATIONAL"])].copy()
    top_themes = (
        best_by_theme[best_by_theme["AssetClass"].isin(["EQUITY","INTERNATIONAL"])]
        .sort_values("TradeScore", ascending=False)
        .head(12)
    )

    top_etfs = trade_candidates.head(15)

    def table_html(df, cols, rename=None):
        if df is None or df.empty:
            return "<p>No data available.</p>"
        cols = [c for c in cols if c in df.columns]
        if not cols:
            return "<p>No compatible columns available.</p>"
        x = df[cols].copy()
        if rename:
            x = x.rename(columns=rename)
        for c in x.columns:
            if pd.api.types.is_numeric_dtype(x[c]):
                x[c] = x[c].map(lambda v: "" if pd.isna(v) else f"{v:,.2f}")
        return x.to_html(index=False, classes="data-table", border=0, escape=False)

    top_sector_html = table_html(
        top5_sectors,
        ["SectorRank","Theme","SectorRRG","SectorStrengthScore","TodaySectorPct",
         "BreadthPct","MedianVolumeVs30D","SectorTurnoverCr"],
        {
            "SectorRank":"Sector Rank",
            "Theme":"Sector / Theme",
            "SectorRRG":"Sector RRG",
            "SectorStrengthScore":"Sector Strength",
            "TodaySectorPct":"Today Sector %",
            "BreadthPct":"Breadth %",
            "MedianVolumeVs30D":"Volume Pace vs Normal",
            "SectorTurnoverCr":"Sector Turnover Cr"
        }
    )

    buy_priority_html = table_html(
        all_latest_gainers,
        ["BuyRank","Symbol","Theme","SectorRank","SectorRRG","NSE_ChangePct",
         "DayVolume","VolumePaceVs30D","TodayTurnoverCr","Tradability",
         "RangePositionPct","Quadrant","TrendPoints","MomentumPhase","EntryStretch",
         "BuyQualityScore","BuyPrioritySignal"],
        {
            "BuyRank":"Buy Priority",
            "Symbol":"ETF",
            "Theme":"Sector / Theme",
            "NSE_ChangePct":"Today %",
            "DayVolume":"Today Volume",
            "VolumePaceVs30D":"Volume Pace vs Normal",
            "TodayTurnoverCr":"Turnover Cr",
            "RangePositionPct":"Near Day High %",
            "Quadrant":"ETF RRG",
            "TrendPoints":"Trend /4",
            "MomentumPhase":"Momentum Phase",
            "EntryStretch":"Entry Stretch",
            "BuyQualityScore":"Buy Quality",
            "BuyPrioritySignal":"Buy Signal"
        }
    )

    near_buy_html = table_html(
        positive_watchlist.head(5),
        ["NearBuyRank","Symbol","Theme","SectorRank","SectorRRG","NSE_ChangePct",
         "DayVolume","VolumePaceVs30D","TodayTurnoverCr","Tradability",
         "RangePositionPct","Quadrant","TrendPoints","MomentumPhase","EntryStretch",
         "BuyQualityScore","WhyNotQualified"],
        {
            "NearBuyRank":"Watch Rank",
            "Symbol":"ETF",
            "Theme":"Sector / Theme",
            "SectorRank":"Sector Rank",
            "SectorRRG":"Sector RRG",
            "NSE_ChangePct":"Today %",
            "DayVolume":"Today Volume",
            "VolumePaceVs30D":"Volume vs Normal",
            "TodayTurnoverCr":"Turnover Cr",
            "RangePositionPct":"Near Day High %",
            "Quadrant":"ETF RRG",
            "TrendPoints":"Trend /4",
            "MomentumPhase":"Momentum Phase",
            "EntryStretch":"Entry Stretch",
            "BuyQualityScore":"Buy Quality",
            "WhyNotQualified":"Why Not Qualified"
        }
    )

    intraday_group_html = table_html(
        top5_intraday_groups,
        ["IntradayGroupRank","Theme","SectorRRG","IntradayGroupStrengthScore",
         "TodaySectorPct","BreadthPct","MedianVolumeVs30D","SectorTurnoverCr"],
        {
            "IntradayGroupRank":"Intraday Group Rank",
            "Theme":"Equity Theme / Market Group",
            "SectorRRG":"Group RRG",
            "IntradayGroupStrengthScore":"Intraday Group Strength",
            "TodaySectorPct":"Today Group %",
            "BreadthPct":"Breadth %",
            "MedianVolumeVs30D":"Volume Pace vs Normal",
            "SectorTurnoverCr":"Group Turnover Cr"
        }
    )

    intraday_buy_html = table_html(
        intraday_buys,
        ["IntradayRank","Symbol","Theme","IntradayGroupRank","SectorRRG",
         "NSE_ChangePct","IntradayOpenDrivePct","DayVolume","VolumePaceVs30D",
         "TodayTurnoverCr","IntradayTradability","RangePositionPct","Quadrant",
         "TrendPoints","IntradayMoveATR","IntradayBuyScore","IntradaySignal"],
        {
            "IntradayRank":"Intraday Priority",
            "Symbol":"ETF",
            "Theme":"Equity Theme / Market Group",
            "IntradayGroupRank":"Group Rank",
            "SectorRRG":"Group RRG",
            "NSE_ChangePct":"Today %",
            "IntradayOpenDrivePct":"Vs Open %",
            "DayVolume":"Today Volume",
            "VolumePaceVs30D":"Volume Pace",
            "TodayTurnoverCr":"Turnover Cr",
            "IntradayTradability":"Tradability",
            "RangePositionPct":"Near Day High %",
            "Quadrant":"ETF RRG",
            "TrendPoints":"Trend /4",
            "IntradayMoveATR":"Move ATR",
            "IntradayBuyScore":"Intraday Score",
            "IntradaySignal":"Intraday Signal"
        }
    )

    intraday_near_html = table_html(
        intraday_watchlist.head(5),
        ["IntradayWatchRank","Symbol","Theme","SectorRRG","NSE_ChangePct",
         "IntradayOpenDrivePct","VolumePaceVs30D","IntradayTradability",
         "RangePositionPct","Quadrant","IntradayMoveATR","IntradayBuyScore",
         "WhyNotIntradayQualified"],
        {
            "IntradayWatchRank":"Watch Rank",
            "Symbol":"ETF",
            "Theme":"Equity Theme / Market Group",
            "SectorRRG":"Group RRG",
            "NSE_ChangePct":"Today %",
            "IntradayOpenDrivePct":"Vs Open %",
            "VolumePaceVs30D":"Volume Pace",
            "IntradayTradability":"Tradability",
            "RangePositionPct":"Near Day High %",
            "Quadrant":"ETF RRG",
            "IntradayMoveATR":"Move ATR",
            "IntradayBuyScore":"Intraday Score",
            "WhyNotIntradayQualified":"Why Not Qualified"
        }
    )

    theme_table = table_html(
        top_themes,
        ["Theme","Symbol","TradeScore","Theme_Quadrant","Quadrant","Avg30TurnoverCr","Week1ReturnPct","Month1ReturnPct","Month3ReturnPct"],
        {"Symbol":"Best ETF","Avg30TurnoverCr":"30D Turnover ₹Cr","Week1ReturnPct":"1W %","Month1ReturnPct":"1M %","Month3ReturnPct":"3M %"}
    )

    etf_table = table_html(
        top_etfs,
        ["RankOverall","Symbol","Theme","TradeScore","Signal","Theme_Quadrant","Quadrant","LTP","Avg30TurnoverCr","VolumeMultiple","DayReturnPct","Week1ReturnPct","Month1ReturnPct"],
        {"Avg30TurnoverCr":"30D Liquidity ₹Cr","VolumeMultiple":"Vol x","DayReturnPct":"1D %","Week1ReturnPct":"1W %","Month1ReturnPct":"1M %"}
    )

    dip_table_cols = [
        c for c in [
            "DipRank","Symbol","Theme","NSE_ChangePct","DayVolume",
            "PctMoveRankScore","CurrentVolumeRankScore","DailyLoserRankScore",
            "CurrentVolumeMultiple30D","DipVolumeCharacter",
            "AggressiveDipScore","DipSignal","LTP","TodayTurnoverCr"
        ]
        if c in dip_candidates.columns
    ]
    dip_table = table_html(
        dip_candidates.head(10),
        dip_table_cols,
        {
            "NSE_ChangePct":"Today %",
            "DayVolume":"NSE Volume",
            "PctMoveRankScore":"% Move Rank",
            "CurrentVolumeRankScore":"Liquidity / Activity Rank",
            "DailyLoserRankScore":"Final Rank Score",
            "CurrentVolumeMultiple30D":"Vol Pace / Normal",
            "DipVolumeCharacter":"Dip Type",
            "AggressiveDipScore":"Dip Score",
            "TodayTurnoverCr":"Turnover Cr"
        }
    )

    asset_table = table_html(
        non_equity.sort_values("TradeScore", ascending=False).head(20),
        ["Symbol","AssetClass","Theme","TradeScore","Quadrant","LTP","Avg30TurnoverCr","DayReturnPct","Month1ReturnPct","Month3ReturnPct"],
        {"Avg30TurnoverCr":"30D Turnover ₹Cr","DayReturnPct":"1D %","Month1ReturnPct":"1M %","Month3ReturnPct":"3M %"}
    )

    changed_today = rotation_changes[rotation_changes["ChangedInLatestSession"]].copy() if not rotation_changes.empty else pd.DataFrame()
    change_table = table_html(
        changed_today.head(20),
        ["Theme","PreviousQuadrant","CurrentQuadrant","RS_Ratio","RS_Momentum","DeltaRatio1D","DeltaMomentum1D"],
        {"DeltaRatio1D":"Δ Ratio 1D","DeltaMomentum1D":"Δ Momentum 1D"}
    )

    theme_rrg_html = fig_theme.to_html(full_html=False, include_plotlyjs="cdn")
    etf_rrg_html = fig_etf.to_html(full_html=False, include_plotlyjs=False)
    asset_rrg_html = fig_asset.to_html(full_html=False, include_plotlyjs=False)

    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <title>NSE ETF RRG Dashboard</title>
    <style>
    body {{
        font-family: Arial, Helvetica, sans-serif;
        margin: 0;
        background: #0B1220;
        color: #E5E7EB;
    }}
    .container {{ max-width: 1500px; margin: auto; padding: 24px; }}
    h1 {{ margin-bottom: 4px; }}
    .sub {{ color:#94A3B8; margin-bottom:22px; }}
    .cards {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(190px,1fr)); gap:14px; margin:18px 0; }}
    .card {{ background:#111827; border:1px solid #243044; border-radius:14px; padding:18px; }}
    .card .k {{ color:#94A3B8; font-size:12px; text-transform:uppercase; letter-spacing:.6px; }}
    .card .v {{ font-size:29px; font-weight:700; margin-top:7px; }}
    .green {{ color:#22C55E; }} .blue {{ color:#60A5FA; }} .amber {{ color:#F59E0B; }} .red {{ color:#EF4444; }}
    .panel {{ background:#FFFFFF; color:#111827; border-radius:16px; padding:15px; margin:18px 0; overflow:auto; }}
    .panel-dark {{ background:#111827; border:1px solid #243044; border-radius:16px; padding:18px; margin:18px 0; }}
    .data-table {{ width:100%; border-collapse:collapse; font-size:13px; }}
    .data-table th {{ background:#172033; color:#F8FAFC; position:sticky; top:0; }}
    .data-table th,.data-table td {{ padding:8px 10px; border-bottom:1px solid #D1D5DB; text-align:right; white-space:nowrap; }}
    .data-table th:first-child,.data-table td:first-child {{ text-align:left; }}
    .note {{ background:#172033; border-left:4px solid #60A5FA; padding:13px 15px; border-radius:8px; color:#CBD5E1; }}
    small {{ color:#94A3B8; }}
    </style>
    </head>
    <body>
    <div class="container">
      <h1>NSE ETF RRG Market Intelligence Dashboard</h1>
      <div class="sub">Analysis Time: {ANALYSIS_TIME_IST} | Market Mode: {CURRENT_SESSION_MODE} | NSE snapshot: {NSE_DOWNLOAD_TIMESTAMP} | Benchmark: {BENCHMARK_LABEL} | Universe: Current NSE Volume ≥ 1,00,000 units</div>

      <div class="cards">
        <div class="card"><div class="k">NSE Snapshot</div><div class="v" style="font-size:16px">{NSE_DOWNLOAD_TIMESTAMP}</div></div>
        <div class="card"><div class="k">NSE ETF Master</div><div class="v">{len(master_all)}</div></div>
        <div class="card"><div class="k">NSE Volume ≥ 1L</div><div class="v">{len(eligible)}</div></div>
        <div class="card"><div class="k">Equity Themes Leading</div><div class="v green">{qcount(eq_theme,'LEADING')}</div></div>
        <div class="card"><div class="k">Equity Themes Improving</div><div class="v blue">{qcount(eq_theme,'IMPROVING')}</div></div>
        <div class="card"><div class="k">Equity Themes Weakening</div><div class="v amber">{qcount(eq_theme,'WEAKENING')}</div></div>
        <div class="card"><div class="k">Equity Themes Lagging</div><div class="v red">{qcount(eq_theme,'LAGGING')}</div></div>
        <div class="card"><div class="k">Momentum Candidates</div><div class="v">{len(trade_candidates)}</div></div>
        <div class="card"><div class="k">Top Daily Losers</div><div class="v amber">{len(dip_candidates)}</div></div>
      </div>

      <div class="note">
      <b>How to read this dashboard:</b>
      First look at the <b>Theme RRG</b>. Leading + Improving themes are the preferred hunting area.
      For swing trades use <b>Equity Swing Buy Priority</b>. Rank 1 is reserved for an ETF that passes every strict swing-entry rule. If none qualify, use the <b>Near Buy Watchlist</b> only as a monitoring list.
      For intraday trades use <b>Intraday Buy Priority</b>; it emphasizes current price strength, price vs open, near-day-high, time-adjusted volume pace and tradability. Intraday scan window: <b>{INTRADAY_SCAN_WINDOW}</b>.
      Gold/Silver/Debt/Liquid ETFs are kept in Group 2. International Equity is included with Equity in Group 1.
      </div>

      <div class="panel-dark">
        <h2>Top 5 Equity Themes / Market Groups — LEADING / IMPROVING only</h2>
        {top_sector_html}
      </div>

      <div class="panel-dark">
        <h2>Equity Swing Buy Priority — Strict Qualified Candidates</h2>
        <p style="color:#94A3B8">Rank 1 is the strongest qualified swing candidate at the analysis time. If no ETF meets all hard conditions, no Buy Rank 1 is forced.</p>
        {buy_priority_html}
      </div>

      <div class="panel-dark">
        <h2>Near Buy Watchlist — Closest Swing Candidates</h2>
        <p style="color:#94A3B8">These are NOT buy signals. They are the five closest ETFs to strict swing qualification, ranked first by the fewest failed rules and then by Buy Quality.</p>
        {near_buy_html}
      </div>

      <div class="panel-dark">
        <h2>Top 5 Intraday Equity Themes / Market Groups</h2>
        <p style="color:#94A3B8">Intraday group ranking gives more weight to today's group performance, breadth and time-adjusted participation. Scan window: <b>{INTRADAY_SCAN_WINDOW}</b>.</p>
        {intraday_group_html}
      </div>

      <div class="panel-dark">
        <h2>Equity Intraday Buy Priority — Current Session</h2>
        <p style="color:#94A3B8">Intraday Rank 1 is the strongest ETF passing all current-session entry rules. Long-term Trend /4 is shown as context but is not a hard intraday gate.</p>
        {intraday_buy_html}
      </div>

      <div class="panel-dark">
        <h2>Intraday Near Buy Watchlist — NOT a Buy Signal</h2>
        {intraday_near_html}
      </div>

      <div class="panel">{theme_rrg_html}</div>

      <div class="panel-dark">
        <h2>Group 1 — Equity Themes + Best ETF in Each Theme</h2>
        {theme_table}
      </div>

      <div class="panel-dark">
        <h2>Group 1 — Equity ETF Momentum / Rotation Candidates</h2>
        {etf_table}
      </div>

      <div class="panel-dark">
        <h2>Negative ETF / Dip Opportunities — Dynamic Liquidity Universe</h2>
        <p style="color:#94A3B8">A falling ETF is not automatically a buy. Highest preference is given when the sector/theme remains LEADING or IMPROVING and the ETF still has a healthy structural trend.</p>
        {dip_table}
      </div>

      <div class="panel">{etf_rrg_html}</div>

      <div class="panel-dark">
        <h2>Theme Quadrant Changes Today</h2>
        {change_table}
      </div>

      <div class="panel">{asset_rrg_html}</div>

      <div class="panel-dark">
        <h2>Group 2 — Gold / Silver / Debt / Liquid</h2>
        {asset_table}
      </div>

      <div class="note">
      <b>Data-quality note:</b> NSE is used for the current ETF master universe. Price and volume history is downloaded from Yahoo Finance using NSE symbols.
      Theme RRG uses the highest-liquidity eligible ETF in that theme as a tradable proxy for the underlying theme/index.
      The workbook includes source and quality-control sheets so you can see what was included or excluded.
      </div>

      <small>Generated automatically by NSE ETF RRG Scanner.</small>
    </div>
    </body>
    </html>
    """

    dashboard_file = DASH_DIR / f"NSE_ETF_RRG_DASHBOARD_{RUN_DATE}.html"
    dashboard_file.write_text(html, encoding="utf-8")

    # -----------------------------
    # STEP 11 — Excel workbook
    # -----------------------------
    print("\n[8/10] Creating professional Excel workbook ...")

    excel_file = EXCEL_DIR / f"NSE_ETF_RRG_ANALYSIS_V4_8_3_{RUN_DATE}.xlsx"

    readme = pd.DataFrame({
        "Item": [
            "Build Version","Run Date","Analysis Time","Market Mode","Intraday Scan Window","NSE Snapshot Timestamp","NSE Market As-Of","Benchmark",
            "Swing Universe","Intraday Universe","Negative / Dip Universe",
            "Master Source","Current Market Source","Historical Source","RRG Method","Important"
        ],
        "Value": [
            "V4.8.3 — FORCED SHEET ORDER",
            RUN_DATE,
            ANALYSIS_TIME_IST,
            CURRENT_SESSION_MODE,
            INTRADAY_SCAN_WINDOW,
            NSE_DOWNLOAD_TIMESTAMP,
            analysis["NSE_AsOf"].dropna().astype(str).replace("", np.nan).dropna().iloc[0] if len(analysis["NSE_AsOf"].dropna().astype(str).replace("", np.nan).dropna()) else "",
            BENCHMARK_LABEL,
            "GROUP 1 — Equity + International: Current NSE Volume >= 1,00,000 units",
            "GROUP 1 — Top 50 Active Equity + International: 50% turnover + 30% volume + 20% time-adjusted volume pace",
            f"GROUP 1 — {NEGATIVE_UNIVERSE_MODE}; GROUP 2 Gold / Silver / Debt / Liquid stay separate in ALTERNATIVE_ASSETS",
            MASTER_SOURCE,
            "NSE ETF Market Data snapshot (authoritative current fields)",
            "Yahoo Finance NSE tickers (.NS) for historical series only",
            f"RS-Ratio: relative strength / {RRG_RS_WINDOW}D mean, centered at 100; RS-Momentum: {RRG_MOM_LAG}D change in RS-Ratio, centered at 100",
            "Analytical scanner only; verify live market price/spread before placing any trade"
        ]
    })

    summary = pd.DataFrame({
        "Metric": [
            "Master ETFs","ETFs with price data","NSE-volume-selected ETFs",
            "Group 1 Equity + International","Group 2 Gold/Silver/Debt/Liquid",
            "Leading Group 1 sectors/themes","Improving Group 1 sectors/themes",
            "Momentum trade candidates","Swing qualified buys","Intraday qualified buys",
            "Negative liquidity universe","Strict Rebound Buys",
            "Deep Fall Watch — Equity / Intl down >=1%","Failed Yahoo batch tickers"
        ],
        "Value": [
            len(master_all), len(etf), len(eligible),
            int(analysis["AssetClass"].isin(["EQUITY","INTERNATIONAL"]).sum()),
            int(analysis["AssetClass"].isin(["GOLD","SILVER","DEBT","LIQUID"]).sum()),
            qcount(eq_theme,"LEADING"), qcount(eq_theme,"IMPROVING"),
            len(trade_candidates), len(all_latest_gainers), len(intraday_buys),
            NEGATIVE_UNIVERSE_COUNT, len(dip_candidates),
            DEEP_FALL_COUNT, len(failed_batches) + len(neg_extra_failed)
        ]
    })

    # ----- Simplified sheet datasets -----
    theme_rrg_valid = theme_rrg.dropna(subset=["RS_Ratio","RS_Momentum"]).copy()
    # Display separation only: THEME_RRG is Group 1 (Equity + International).
    # Gold / Silver / Debt / Liquid remain in ALTERNATIVE_ASSETS.
    theme_rrg_valid = theme_rrg_valid[
        theme_rrg_valid["AssetClass"].isin(["EQUITY","INTERNATIONAL"])
    ].copy()
    theme_rrg_valid = theme_rrg_valid.sort_values(
        ["RS_Ratio","RS_Momentum"], ascending=[False,False]
    ).reset_index(drop=True)
    theme_rrg_valid.insert(0, "ThemeRank", np.arange(1, len(theme_rrg_valid)+1))

    analysis_full = analysis.sort_values(["RankOverall","TradeScore","AssetClass","Theme"], ascending=[True,False,True,True]).copy()

    # Columns used in the two focused trading sheets.
    # Defined here explicitly so the simplified workbook is self-contained.
    ranking_cols = [
        "RankOverall","RankInTheme","Symbol","Underlying","Theme","AssetClass",
        "NSE_DownloadTimestamp","NSE_AsOf","NSE_PrevClose","NSE_Open","NSE_High","NSE_Low",
        "NSE_Change","NSE_ChangePct","LTP","RangePositionPct",
        "DayVolume","Avg30Volume","CurrentVolumeMultiple30D","CurrentVsAvg30VolumePct",
        "VolumeSurgeFlag","LatestDataMode",
        "TodayTurnoverCr","Avg30TurnoverCr","VolumeMultiple",
        "DayReturnPct","Week1ReturnPct","Month1ReturnPct","Month3ReturnPct",
        "DMA20","DMA50","DMA100","DMA200","TrendPoints",
        "RS_Ratio","RS_Momentum","Quadrant",
        "Theme_RS_Ratio","Theme_RS_Momentum","Theme_Quadrant",
        "LiquidityScore","VolumeImpulseScore","MomentumScore","TrendScore",
        "ETF_RRG_Score","ThemeRRGScore","TradeScore","Signal",
        "BuyingPressureScore","BuyingSignal",
        "DipBuyScore","DipSignal","DipVolumeCharacter","DipBucket"
    ]
    ranking_cols = [c for c in ranking_cols if c in analysis.columns]

    trade_candidate_cols = [c for c in ranking_cols if c in trade_candidates.columns]
    trade_candidates_sheet = trade_candidates[trade_candidate_cols].copy()
    if not trade_candidates_sheet.empty:
        trade_candidates_sheet = trade_candidates_sheet.sort_values(
            ["TradeScore","Avg30TurnoverCr"], ascending=[False,False]
        ).reset_index(drop=True)
        trade_candidates_sheet.insert(0, "TradeRank", np.arange(1, len(trade_candidates_sheet)+1))
    if trade_candidates_sheet.empty:
        trade_candidates_sheet = pd.DataFrame({
            "Status": ["No momentum trade candidates in the latest NSE session."],
            "NSE_DownloadTimestamp": [NSE_DOWNLOAD_TIMESTAMP]
        })

    dip_excel_cols = [
        c for c in [
            "DipRank","Symbol","Underlying","Theme","AssetClass",
            "NSE_ChangePct","DayVolume",
            "PctMoveRankScore","CurrentVolumeRankScore","DailyLoserRankScore",
            "LTP","NSE_Open","NSE_High","NSE_Low","RangePositionPct",
            "Avg30Volume","CurrentVolumeMultiple30D","CurrentVsAvg30VolumePct",
            "TodayTurnoverCr","LiquidityStatus","LiquidityTier","LiquidEligible",
            "DipEligibilityResult","DipVolumeCharacter","AggressiveDipScore","DipSignal"
        ]
        if c in dip_candidates.columns
    ]
    dip_candidates_sheet = dip_candidates[dip_excel_cols].copy() if dip_excel_cols else pd.DataFrame()

    if dip_candidates_sheet.empty:
        dip_candidates_sheet = pd.DataFrame({
            "Status": ["No negative Equity/International ETF in the selected NSE Volume >= 1 lakh universe."],
            "Rule": ["Top 10 losers ranked 60% by % fall + 40% by current NSE volume"],
            "NSE_DownloadTimestamp": [NSE_DOWNLOAD_TIMESTAMP]
        })

    alternative_assets = analysis_full[analysis_full["AssetClass"].isin(["GOLD","SILVER","DEBT","LIQUID"])].copy()
    if not alternative_assets.empty:
        _alt_order = {"GOLD":0, "SILVER":1, "DEBT":2, "LIQUID":3}
        alternative_assets["_AltClassOrder"] = (
            alternative_assets["AssetClass"].map(_alt_order).fillna(9)
        )
        alternative_assets = alternative_assets.sort_values(
            ["_AltClassOrder","TradeScore","TodayTurnoverCr"],
            ascending=[True,False,False]
        ).reset_index(drop=True)
        alternative_assets.drop(columns=["_AltClassOrder"], inplace=True)
        alternative_assets.insert(0, "AltRank", np.arange(1, len(alternative_assets)+1))

    # For graphical dashboard
    turnover_by_theme = (
        analysis_full[analysis_full["AssetClass"].isin(["EQUITY","INTERNATIONAL"])]
        .groupby("Theme", dropna=False)["TodayTurnoverCr"]
        .sum()
        .sort_values(ascending=False)
        .head(10)
        .reset_index()
    )
    turnover_by_theme.columns = ["Theme","LatestSessionTurnoverCr"]

    top_etf_turnover = (
        analysis_full[analysis_full["AssetClass"].isin(["EQUITY","INTERNATIONAL"])]
        .sort_values("TodayTurnoverCr", ascending=False)
        [["Symbol","Theme","AssetClass","TodayTurnoverCr","NSE_ChangePct"]]
        .head(10)
        .copy()
    )

    top_volume_surges = (
        analysis_full[
            analysis_full["AssetClass"].isin(["EQUITY","INTERNATIONAL"]) &
            analysis_full["CurrentVsAvg30VolumePct"].notna()
        ]
        .sort_values(["CurrentVsAvg30VolumePct","TodayTurnoverCr"], ascending=[False,False])
        [["Symbol","Theme","AssetClass","DayVolume","Avg30Volume",
          "CurrentVsAvg30VolumePct","VolumeSurgeFlag","NSE_ChangePct","TodayTurnoverCr"]]
        .head(10)
        .reset_index(drop=True)
        .copy()
    )
    top_volume_surges.insert(0, "VolumeRank", np.arange(1, len(top_volume_surges)+1))
    top_volume_surges = top_volume_surges.rename(columns={
        "CurrentVsAvg30VolumePct":"VolVs30DAvgPct",
        "VolumeSurgeFlag":"VolumeFlag",
        "NSE_ChangePct":"ChgPct",
        "TodayTurnoverCr":"TurnoverCr"
    })

    dashboard_volume_surges = top_volume_surges.copy().rename(columns={
        "VolumeRank":"Volume Rank",
        "Symbol":"ETF",
        "AssetClass":"Asset Class",
        "DayVolume":"Today Volume",
        "Avg30Volume":"Avg 30D Volume",
        "VolVs30DAvgPct":"Vol vs 30D %",
        "VolumeFlag":"Volume Flag",
        "ChgPct":"Today %",
        "TurnoverCr":"Turnover Cr"
    })

    top5_sectors_sheet = top5_sectors[[
        c for c in [
            "SectorRank","Theme","SectorRRG","SectorStrengthScore",
            "RRGStrengthScore","RRGMomentumScore","TodaySectorPct","TodayPerformanceScore",
            "BreadthPct","BreadthScore","MedianVolumeVs30D","ParticipationScore",
            "SectorTurnoverCr"
        ] if c in top5_sectors.columns
    ]].copy()

    if not top5_sectors_sheet.empty:
        top5_sectors_sheet = top5_sectors_sheet.rename(columns={
            "SectorRank":"Sector Rank",
            "Theme":"Sector / Theme",
            "SectorRRG":"Sector RRG",
            "SectorStrengthScore":"Sector Strength",
            "RRGStrengthScore":"RRG Strength",
            "RRGMomentumScore":"RRG Momentum",
            "TodaySectorPct":"Today Sector %",
            "TodayPerformanceScore":"Today Performance",
            "BreadthPct":"Breadth %",
            "BreadthScore":"Breadth Score",
            "MedianVolumeVs30D":"Volume Pace vs Normal",
            "ParticipationScore":"Participation",
            "SectorTurnoverCr":"Sector Turnover Cr"
        })

    dashboard_top_sectors = top5_sectors.copy()
    if not dashboard_top_sectors.empty:
        dashboard_top_sectors = dashboard_top_sectors[[
            c for c in [
                "SectorRank","Theme","SectorRRG","SectorStrengthScore",
                "TodaySectorPct","BreadthPct","MedianVolumeVs30D","SectorTurnoverCr"
            ] if c in dashboard_top_sectors.columns
        ]].rename(columns={
            "SectorRank":"Sector Rank",
            "Theme":"Sector / Theme",
            "SectorRRG":"Sector RRG",
            "SectorStrengthScore":"Sector Strength",
            "TodaySectorPct":"Today Sector %",
            "BreadthPct":"Breadth %",
            "MedianVolumeVs30D":"Sector Vol vs 30D",
            "SectorTurnoverCr":"Sector Turnover Cr"
        })

    dashboard_buying = all_latest_gainers.head(10).copy()
    if not dashboard_buying.empty:
        dashboard_buying = dashboard_buying[[
            c for c in [
                "BuyRank","Symbol","Theme","SectorRRG","SectorRank",
                "NSE_ChangePct","DayVolume","VolumePaceVs30D",
                "TodayTurnoverCr","Tradability","RangePositionPct","Quadrant",
                "TrendPoints","MomentumPhase","EntryStretch",
                "BuyQualityScore","BuyPrioritySignal"
            ] if c in dashboard_buying.columns
        ]].rename(columns={
            "BuyRank":"Buy Priority",
            "Symbol":"ETF",
            "Theme":"Sector / Theme",
            "SectorRRG":"Sector RRG",
            "SectorRank":"Sector Rank",
            "NSE_ChangePct":"Today %",
            "DayVolume":"Today Volume",
            "VolumePaceVs30D":"Volume vs Normal",
            "TodayTurnoverCr":"Turnover Cr",
            "Tradability":"Tradability",
            "RangePositionPct":"Near Day High %",
            "Quadrant":"ETF RRG",
            "TrendPoints":"Trend /4",
            "MomentumPhase":"Momentum Phase",
            "EntryStretch":"Entry Stretch",
            "BuyQualityScore":"Buy Quality",
            "BuyPrioritySignal":"Buy Signal"
        })

    dashboard_near_buy = positive_watchlist.head(5).copy()
    if not dashboard_near_buy.empty:
        dashboard_near_buy = dashboard_near_buy[[
            c for c in [
                "NearBuyRank","Symbol","Theme","SectorRRG","SectorRank",
                "NSE_ChangePct","DayVolume","VolumePaceVs30D",
                "TodayTurnoverCr","Tradability","RangePositionPct","Quadrant",
                "TrendPoints","MomentumPhase","EntryStretch",
                "BuyQualityScore","WhyNotQualified"
            ] if c in dashboard_near_buy.columns
        ]].rename(columns={
            "NearBuyRank":"Watch Rank",
            "Symbol":"ETF",
            "Theme":"Sector / Theme",
            "SectorRRG":"Sector RRG",
            "SectorRank":"Sector Rank",
            "NSE_ChangePct":"Today %",
            "DayVolume":"Today Volume",
            "VolumePaceVs30D":"Volume vs Normal",
            "TodayTurnoverCr":"Turnover Cr",
            "Tradability":"Tradability",
            "RangePositionPct":"Near Day High %",
            "Quadrant":"ETF RRG",
            "TrendPoints":"Trend /4",
            "MomentumPhase":"Momentum Phase",
            "EntryStretch":"Entry Stretch",
            "BuyQualityScore":"Buy Quality",
            "WhyNotQualified":"Why Not Qualified"
        })

    dashboard_intraday_groups = top5_intraday_groups.head(5).copy()
    if not dashboard_intraday_groups.empty:
        dashboard_intraday_groups = dashboard_intraday_groups[[
            c for c in [
                "IntradayGroupRank","Theme","SectorRRG","IntradayGroupStrengthScore",
                "TodaySectorPct","BreadthPct","MedianVolumeVs30D","SectorTurnoverCr"
            ] if c in dashboard_intraday_groups.columns
        ]].rename(columns={
            "IntradayGroupRank":"Group Rank",
            "Theme":"Equity Theme / Market Group",
            "SectorRRG":"Group RRG",
            "IntradayGroupStrengthScore":"Intraday Group Strength",
            "TodaySectorPct":"Today Group %",
            "BreadthPct":"Breadth %",
            "MedianVolumeVs30D":"Volume Pace",
            "SectorTurnoverCr":"Turnover Cr"
        })

    dashboard_intraday_buying = intraday_buys.head(10).copy()
    if not dashboard_intraday_buying.empty:
        dashboard_intraday_buying = dashboard_intraday_buying[[
            c for c in [
                "IntradayRank","Symbol","Theme","IntradayGroupRank","SectorRRG",
                "NSE_ChangePct","IntradayOpenDrivePct","DayVolume","VolumePaceVs30D",
                "TodayTurnoverCr","IntradayTradability","RangePositionPct","Quadrant",
                "TrendPoints","IntradayMoveATR","IntradayBuyScore","IntradaySignal"
            ] if c in dashboard_intraday_buying.columns
        ]].rename(columns={
            "IntradayRank":"Intraday Priority",
            "Symbol":"ETF",
            "Theme":"Equity Theme / Market Group",
            "IntradayGroupRank":"Group Rank",
            "SectorRRG":"Group RRG",
            "NSE_ChangePct":"Today %",
            "IntradayOpenDrivePct":"Vs Open %",
            "DayVolume":"Today Volume",
            "VolumePaceVs30D":"Volume Pace",
            "TodayTurnoverCr":"Turnover Cr",
            "IntradayTradability":"Tradability",
            "RangePositionPct":"Near Day High %",
            "Quadrant":"ETF RRG",
            "TrendPoints":"Trend /4",
            "IntradayMoveATR":"Move ATR",
            "IntradayBuyScore":"Intraday Score",
            "IntradaySignal":"Intraday Signal"
        })

    dashboard_intraday_near = intraday_watchlist.head(5).copy()
    if not dashboard_intraday_near.empty:
        dashboard_intraday_near = dashboard_intraday_near[[
            c for c in [
                "IntradayWatchRank","Symbol","Theme","SectorRRG","NSE_ChangePct",
                "IntradayOpenDrivePct","VolumePaceVs30D","IntradayTradability",
                "RangePositionPct","Quadrant","IntradayMoveATR","IntradayBuyScore",
                "WhyNotIntradayQualified"
            ] if c in dashboard_intraday_near.columns
        ]].rename(columns={
            "IntradayWatchRank":"Watch Rank",
            "Symbol":"ETF",
            "Theme":"Equity Theme / Market Group",
            "SectorRRG":"Group RRG",
            "NSE_ChangePct":"Today %",
            "IntradayOpenDrivePct":"Vs Open %",
            "VolumePaceVs30D":"Volume Pace",
            "IntradayTradability":"Tradability",
            "RangePositionPct":"Near Day High %",
            "Quadrant":"ETF RRG",
            "IntradayMoveATR":"Move ATR",
            "IntradayBuyScore":"Intraday Score",
            "WhyNotIntradayQualified":"Why Not Qualified"
        })

    dashboard_all_losers = all_latest_losers.head(10).copy()
    if not dashboard_all_losers.empty:
        dashboard_all_losers = dashboard_all_losers[[
            c for c in [
                "DipRank","Symbol","Theme","Theme_Quadrant","Quadrant",
                "NSE_ChangePct","RangePositionPct","ReboundVsOpenPct",
                "CurrentVolumeMultiple30D","TrendPoints","ReboundScore","DipSignal"
            ] if c in dashboard_all_losers.columns
        ]].rename(columns={
            "DipRank":"Rebound Rank",
            "Symbol":"ETF",
            "Theme":"Theme",
            "Theme_Quadrant":"Group RRG",
            "Quadrant":"ETF RRG",
            "NSE_ChangePct":"Today %",
            "RangePositionPct":"Recovery %",
            "ReboundVsOpenPct":"Vs Open %",
            "CurrentVolumeMultiple30D":"Vol Pace",
            "TrendPoints":"Trend /4",
            "ReboundScore":"Rebound Score",
            "DipSignal":"Signal"
        })

    # Display-only list of liquidity-qualified Group 1 ETFs down >= 1%.
    dashboard_deep_fall = deep_fall_watch.head(10).copy()
    if not dashboard_deep_fall.empty:
        dashboard_deep_fall = dashboard_deep_fall[[
            c for c in [
                "DeepFallRank","Symbol","Theme","Theme_Quadrant","Quadrant",
                "NSE_ChangePct","RangePositionPct","ReboundVsOpenPct",
                "CurrentVolumeMultiple30D","TrendPoints","ReboundScore","DipSignal"
            ] if c in dashboard_deep_fall.columns
        ]].rename(columns={
            "DeepFallRank":"Fall Rank",
            "Symbol":"ETF",
            "Theme":"Theme",
            "Theme_Quadrant":"Group RRG",
            "Quadrant":"ETF RRG",
            "NSE_ChangePct":"Today %",
            "RangePositionPct":"Recovery %",
            "ReboundVsOpenPct":"Vs Open %",
            "CurrentVolumeMultiple30D":"Vol Pace",
            "TrendPoints":"Trend /4",
            "ReboundScore":"Rebound Score",
            "DipSignal":"Status"
        })

    dashboard_top_candidates = trade_candidates.head(12).copy()
    if not dashboard_top_candidates.empty:
        dashboard_top_candidates = dashboard_top_candidates[[
            "RankOverall","Symbol","Theme","TradeScore","Signal","Theme_Quadrant","Quadrant","NSE_ChangePct","TodayTurnoverCr","Avg30TurnoverCr"
        ]].rename(columns={
            "TodayTurnoverCr":"LatestSessionTurnoverCr",
            "Avg30TurnoverCr":"Avg30TurnoverCr"
        })

    dashboard_dips = dip_candidates.head(10).copy()
    if not dashboard_dips.empty:
        dashboard_dips = dashboard_dips[[
            c for c in [
                "Symbol","Theme","NSE_ChangePct","ReboundScore","DipSignal",
                "Theme_Quadrant","Quadrant","TodayTurnoverCr"
            ] if c in dashboard_dips.columns
        ]].rename(columns={
            "ReboundScore":"DipBuyScore",
            "TodayTurnoverCr":"LatestSessionTurnoverCr"
        })

    excel_file = EXCEL_DIR / f"NSE_ETF_RRG_ANALYSIS_{RUN_DATE}.xlsx"

    with pd.ExcelWriter(excel_file, engine="xlsxwriter") as writer:
        # Sheet order matters: user wants the visual dashboard first and info last.
        pd.DataFrame({"Dashboard": ["See formatted dashboard sheet."]}).to_excel(
            writer, sheet_name="DASHBOARD", index=False
        )
        # User-facing navigation: INTRADAY immediately after DASHBOARD.
        pd.DataFrame().to_excel(
            writer, sheet_name="INTRADAY", index=False, header=False
        )

        # Hidden helper sheet used only as chart source.
        # This keeps the visible DASHBOARD completely clean (no P:W helper data).
        pd.DataFrame({"Helper": []}).to_excel(writer, sheet_name="_CHART_DATA", index=False)
        analysis_full.to_excel(writer, sheet_name="ETF_ANALYSIS", index=False)
        top5_sectors_sheet.to_excel(writer, sheet_name="TOP_5_SECTORS", index=False)
        theme_rrg_valid.to_excel(writer, sheet_name="THEME_RRG", index=False)
        trade_candidates_sheet.to_excel(writer, sheet_name="TRADE_CANDIDATES", index=False)

        buying_sheet = all_latest_gainers[[
            c for c in [
                "BuyRank","Symbol","Underlying","Theme","AssetClass",
                "SectorRank","SectorRRG","SectorStrengthScore",
                "NSE_ChangePct","DayVolume","Prior30AvgVolume","VolumePaceVs30D","TodayTurnoverCr",
                "Tradability","TradabilityScore","RangePositionPct",
                "Quadrant","TrendPoints","MomentumPhase","MomentumImprovementScore",
                "EntryStretch","EntryStretchATR","ATR20",
                "TodayPriceStrengthScore","CurrentNSEVolumeScore","TurnoverScore",
                "VolumeVs30DScore","NearDayHighScore","TrendStrengthScore",
                "ETFRRGStrengthScore","EntryStretchScore","BuyQualityScore",
                "BuyPrioritySignal","LTP","NSE_Open","NSE_High","NSE_Low",
                "Week1ReturnPct","Month1ReturnPct","Month3ReturnPct",
                "DMA20","DMA50","DMA100","DMA200",
                "NSE_DownloadTimestamp","NSE_AsOf","LatestDataMode"
            ] if c in all_latest_gainers.columns
        ]].copy()

        if buying_sheet.empty:
            buying_sheet = pd.DataFrame({
                "Status":["NO QUALIFIED BUY NOW"],
                "Reason":["No Equity / International Equity ETF passed every strict positive-side swing condition."],
                "Analysis Time":[ANALYSIS_TIME_IST],
                "Market Mode":[CURRENT_SESSION_MODE],
                "NSE_DownloadTimestamp":[NSE_DOWNLOAD_TIMESTAMP]
            })
        else:
            buying_sheet = buying_sheet.rename(columns={
                "BuyRank":"Buy Priority",
                "Symbol":"ETF",
                "Theme":"Sector / Theme",
                "SectorRank":"Sector Rank",
                "SectorRRG":"Sector RRG",
                "SectorStrengthScore":"Sector Strength",
                "NSE_ChangePct":"Today %",
                "DayVolume":"Today Volume",
                "VolumePaceVs30D":"Volume vs 30D",
                "TodayTurnoverCr":"Turnover Cr",
                "RangePositionPct":"Near Day High %",
                "Quadrant":"ETF RRG",
                "TrendPoints":"Trend /4",
                "MomentumPhase":"Momentum Phase",
                "MomentumImprovementScore":"Momentum Improvement",
                "EntryStretch":"Entry Stretch",
                "EntryStretchATR":"Stretch ATR",
                "BuyQualityScore":"Buy Quality",
                "BuyPrioritySignal":"Buy Signal",
                "LatestDataMode":"Market Mode"
            })

        buying_sheet.to_excel(writer, sheet_name="BUYING_ANALYSIS", index=False)

        near_buy_sheet = positive_watchlist.head(20)[[
            c for c in [
                "NearBuyRank","Symbol","Underlying","Theme","SectorRank","SectorRRG",
                "SectorStrengthScore","Quadrant","NSE_ChangePct","DayVolume",
                "VolumePaceVs30D","TodayTurnoverCr","Tradability","TradabilityScore",
                "RangePositionPct","TrendPoints","MomentumPhase",
                "MomentumImprovementScore","EntryStretch","EntryStretchATR",
                "BuyQualityScore","FailedRuleCount","WhyNotQualified",
                "NSE_DownloadTimestamp","NSE_AsOf","LatestDataMode"
            ] if c in positive_watchlist.columns
        ]].copy()

        near_buy_sheet = near_buy_sheet.rename(columns={
            "NearBuyRank":"Watch Rank",
            "Symbol":"ETF",
            "Theme":"Sector / Theme",
            "SectorRank":"Sector Rank",
            "SectorRRG":"Sector RRG",
            "SectorStrengthScore":"Sector Strength",
            "Quadrant":"ETF RRG",
            "NSE_ChangePct":"Today %",
            "DayVolume":"Today Volume",
            "VolumePaceVs30D":"Volume vs Normal",
            "TodayTurnoverCr":"Turnover Cr",
            "TradabilityScore":"Tradability Score",
            "RangePositionPct":"Near Day High %",
            "TrendPoints":"Trend /4",
            "MomentumPhase":"Momentum Phase",
            "MomentumImprovementScore":"Momentum Improvement",
            "EntryStretch":"Entry Stretch",
            "EntryStretchATR":"Stretch ATR",
            "BuyQualityScore":"Buy Quality",
            "FailedRuleCount":"Failed Rules",
            "WhyNotQualified":"Why Not Qualified",
            "LatestDataMode":"Market Mode"
        })
        near_buy_sheet.to_excel(writer, sheet_name="NEAR_BUY_WATCHLIST", index=False)

        # ------------------------------------------------------------
        # ONE COMPACT INTRADAY SHEET
        # All detailed intraday calculations remain in memory / ETF_ANALYSIS.
        # The visible INTRADAY sheet shows only the minimum decision fields.
        # ------------------------------------------------------------
        intraday_group_compact = top5_intraday_groups[[
            c for c in [
                "IntradayGroupRank","Theme","SectorRRG","IntradayGroupStrengthScore",
                "TodaySectorPct","BreadthPct","MedianVolumeVs30D"
            ] if c in top5_intraday_groups.columns
        ]].copy().rename(columns={
            "IntradayGroupRank":"Group Rank",
            "Theme":"Equity Group",
            "SectorRRG":"Group RRG",
            "IntradayGroupStrengthScore":"Group Strength",
            "TodaySectorPct":"Today %",
            "BreadthPct":"Breadth %",
            "MedianVolumeVs30D":"Volume Pace"
        })

        intraday_buy_compact = intraday_buys.head(5)[[
            c for c in [
                "IntradayRank","Symbol","Theme","SectorRRG","NSE_ChangePct",
                "IntradayOpenDrivePct","VolumePaceVs30D","IntradayTradability",
                "RangePositionPct","Quadrant","IntradayBuyScore","IntradaySignal"
            ] if c in intraday_buys.columns
        ]].copy().rename(columns={
            "IntradayRank":"Priority",
            "Symbol":"ETF",
            "Theme":"Equity Group",
            "SectorRRG":"Group RRG",
            "NSE_ChangePct":"Today %",
            "IntradayOpenDrivePct":"Vs Open %",
            "VolumePaceVs30D":"Volume Pace",
            "IntradayTradability":"Tradability",
            "RangePositionPct":"Near Day High %",
            "Quadrant":"ETF RRG",
            "IntradayBuyScore":"Intraday Score",
            "IntradaySignal":"Signal"
        })

        intraday_near_compact = intraday_watchlist.head(5)[[
            c for c in [
                "IntradayWatchRank","Symbol","Theme","NSE_ChangePct",
                "VolumePaceVs30D","RangePositionPct","IntradayBuyScore",
                "WhyNotIntradayQualified"
            ] if c in intraday_watchlist.columns
        ]].copy().rename(columns={
            "IntradayWatchRank":"Watch Rank",
            "Symbol":"ETF",
            "Theme":"Equity Group",
            "NSE_ChangePct":"Today %",
            "VolumePaceVs30D":"Volume Pace",
            "RangePositionPct":"Near Day High %",
            "IntradayBuyScore":"Intraday Score",
            "WhyNotIntradayQualified":"Why Not Qualified"
        })

        # Create only ONE visible intraday worksheet. Detailed calculations are not
        # repeated across separate worksheets.
        # INTRADAY was created immediately after DASHBOARD for tab order.

        all_losers_sheet = all_latest_losers[[
            c for c in [
                # Decision fields first
                "DipRank","DipBuyPriority","ReboundWatchRank","Symbol","Theme",
                "Theme_Quadrant","Quadrant","NSE_ChangePct",
                "RangePositionPct","ReboundVsOpenPct","CurrentVolumeMultiple30D",
                "TodayTurnoverCr","Avg30TurnoverCr","TrendPoints",
                "ReboundScore","ReboundQualified","DipSignal",
                "WhyNotReboundQualified",
                # Supporting / background fields
                "Underlying","AssetClass","DayVolume",
                "LTP","NSE_Open","NSE_High","NSE_Low",
                "Avg30Volume","Prior30AvgVolume",
                "NegativeActivityScore","NegativeLiquidityScore",
                "Week1ReturnPct","Month1ReturnPct","Month3ReturnPct",
                "ATR20","DayFallATR","DipVolumeCharacter",
                "RecoveryFromLowScore","ReboundVsOpenScore",
                "ReboundThemeRRGScore","ReboundETFRRGScore",
                "ReboundTrendScore","ReboundMomentumScore",
                "ReboundVolumeCharacterScore","ReboundLiquidityScore",
                "ReboundDipMagnitudeScore","ReboundATRScore",
                "LiquidityStatus","LiquidityTier","LiquidEligible",
                "DipEligibilityResult","AggressiveDipScore"
            ] if c in all_latest_losers.columns
        ]].copy().rename(columns={
            "DipRank":"Rebound Rank",
            "DipBuyPriority":"Buy Priority",
            "ReboundWatchRank":"Watch Rank",
            "Symbol":"ETF",
            "Theme":"Theme",
            "Theme_Quadrant":"Group RRG",
            "Quadrant":"ETF RRG",
            "NSE_ChangePct":"Today %",
            "RangePositionPct":"Recovery %",
            "ReboundVsOpenPct":"Vs Open %",
            "CurrentVolumeMultiple30D":"Vol Pace",
            "TodayTurnoverCr":"Turnover Cr",
            "Avg30TurnoverCr":"30D Avg Turnover Cr",
            "TrendPoints":"Trend /4",
            "ReboundScore":"Rebound Score",
            "ReboundQualified":"Qualified",
            "DipSignal":"Signal",
            "WhyNotReboundQualified":"Why Not Qualified"
        })
        if "Qualified" in all_losers_sheet.columns:
            all_losers_sheet["Qualified"] = (
                all_losers_sheet["Qualified"].map({True:"YES", False:"NO"}).fillna("")
            )
        all_losers_sheet.to_excel(
            writer, sheet_name="LOSERS_DIP_ANALYSIS", index=False
        )
        alternative_assets.to_excel(writer, sheet_name="ALTERNATIVE_ASSETS", index=False)

        # Combine README + RUN SUMMARY into a single final sheet.
        # Blank creation avoids a duplicate PROJECT INFO row.
        pd.DataFrame().to_excel(
            writer, sheet_name="INFO", index=False, header=False
        )
        readme.to_excel(writer, sheet_name="INFO", index=False, startrow=2)
        summary.to_excel(writer, sheet_name="INFO", index=False, startrow=len(readme)+5)

        book = writer.book

        # ---------- formats ----------
        fmt_title = book.add_format({"bold": True, "font_color": "white", "bg_color": "#0F172A", "font_size": 18, "align":"center", "valign":"vcenter"})
        fmt_sub = book.add_format({"bold": False, "font_color": "#0F172A", "bg_color": "#DCE6F1", "font_size": 11})
        fmt_card_head = book.add_format({"bold": True, "font_color": "white", "bg_color": "#1E3A8A", "align":"center", "valign":"vcenter", "border":1})
        fmt_card_val = book.add_format({"bold": True, "font_color": "#0F172A", "bg_color": "#EAF2F8", "align":"center", "valign":"vcenter", "border":1, "font_size": 15})
        fmt_header = book.add_format({"bold": True, "font_color": "white", "bg_color": "#1E3A8A", "border":2, "border_color":"#000000", "text_wrap": True, "align":"center", "valign":"vcenter"})
        fmt_pct = book.add_format({"num_format": "0.00;[Red](0.00);-"})
        fmt_num = book.add_format({"num_format": "#,##0.00;[Red](#,##0.00);-"})
        fmt_cr = book.add_format({"num_format": '₹0.00" Cr";[Red](₹0.00" Cr");-'})
        fmt_wrap = book.add_format({"text_wrap": True, "valign": "top"})
        fmt_note = book.add_format({"text_wrap": True, "valign": "top", "font_color":"#334155", "bg_color":"#F8FAFC", "border":1})
        fmt_body = book.add_format({"border":2, "border_color":"#000000", "valign":"vcenter"})
        fmt_body_wrap = book.add_format({"border":2, "border_color":"#000000", "valign":"top", "text_wrap":True})
        fmt_body_pct = book.add_format({"border":2, "border_color":"#000000", "num_format":"0.00;[Red](0.00);-", "valign":"vcenter"})
        fmt_body_num = book.add_format({"border":2, "border_color":"#000000", "num_format":"#,##0.00;[Red](#,##0.00);-", "valign":"vcenter"})
        fmt_body_int = book.add_format({"border":2, "border_color":"#000000", "num_format":"#,##0;[Red](#,##0);-", "valign":"vcenter"})
        fmt_negative_pct = book.add_format({"bg_color":"#FDE9E7","font_color":"#C00000","border":2,"border_color":"#000000","num_format":"0.00;[Red](0.00);-","valign":"vcenter"})
        fmt_positive_pct = book.add_format({"bg_color":"#E2F0D9","font_color":"#006100","border":2,"border_color":"#000000","num_format":"0.00;[Red](0.00);-","valign":"vcenter"})
        fmt_vol_extreme = book.add_format({"bg_color":"#F4B183","font_color":"#9C0006","bold":True,"border":2,"border_color":"#000000","valign":"vcenter","num_format":"0.00;[Red](0.00);-"})
        fmt_vol_very_high = book.add_format({"bg_color":"#FFD966","font_color":"#7F6000","bold":True,"border":2,"border_color":"#000000","valign":"vcenter","num_format":"0.00;[Red](0.00);-"})
        fmt_vol_high = book.add_format({"bg_color":"#FFF2CC","font_color":"#7F6000","border":2,"border_color":"#000000","valign":"vcenter","num_format":"0.00;[Red](0.00);-"})
        fmt_vol_above = book.add_format({"bg_color":"#D9EAF7","font_color":"#1F4E78","border":2,"border_color":"#000000","valign":"vcenter","num_format":"0.00;[Red](0.00);-"})
        fmt_grid_blank = book.add_format({"border":2,"border_color":"#000000","valign":"vcenter"})
        fmt_qualified_fill = book.add_format({"bg_color":"#C6EFCE","border":2,"border_color":"#000000","valign":"vcenter"})
        fmt_emerging_fill = book.add_format({"bg_color":"#FFF2CC","border":2,"border_color":"#000000","valign":"vcenter"})
        fmt_watch_fill = book.add_format({"bg_color":"#F2F2F2","border":2,"border_color":"#000000","valign":"vcenter"})
        fmt_buy_a = book.add_format({"bg_color":"#70AD47","font_color":"#FFFFFF","bold":True,"border":2,"border_color":"#000000","valign":"vcenter"})
        fmt_buy_strong = book.add_format({"bg_color":"#A9D18E","bold":True,"border":2,"border_color":"#000000","valign":"vcenter"})
        fmt_buy_pressure = book.add_format({"bg_color":"#C6E0B4","border":2,"border_color":"#000000","valign":"vcenter"})
        fmt_buy_watch = book.add_format({"bg_color":"#D9EAF7","border":2,"border_color":"#000000","valign":"vcenter"})
        fmt_buy_fading = book.add_format({"bg_color":"#F4CCCC","border":2,"border_color":"#000000","valign":"vcenter"})
        fmt_dip_a = book.add_format({"bg_color":"#548235","font_color":"#FFFFFF","bold":True,"border":2,"border_color":"#000000","valign":"vcenter"})
        fmt_dip_strong = book.add_format({"bg_color":"#A9D18E","bold":True,"border":2,"border_color":"#000000","valign":"vcenter"})
        fmt_dip_volume = book.add_format({"bg_color":"#C6E0B4","border":2,"border_color":"#000000","valign":"vcenter"})
        fmt_dip_watch = book.add_format({"bg_color":"#FFF2CC","border":2,"border_color":"#000000","valign":"vcenter"})
        fmt_dip_low = book.add_format({"bg_color":"#F2F2F2","border":2,"border_color":"#000000","valign":"vcenter"})
        fmt_sector_leading = book.add_format({"bg_color":"#70AD47","font_color":"#FFFFFF","bold":True,"border":2,"border_color":"#000000","valign":"vcenter"})
        fmt_sector_improving = book.add_format({"bg_color":"#9DC3E6","font_color":"#000000","bold":True,"border":2,"border_color":"#000000","valign":"vcenter"})
        fmt_priority_1 = book.add_format({"bg_color":"#FFD966","font_color":"#000000","bold":True,"border":2,"border_color":"#000000","valign":"vcenter"})
        fmt_buy_top = book.add_format({"bg_color":"#548235","font_color":"#FFFFFF","bold":True,"border":2,"border_color":"#000000","valign":"vcenter"})
        fmt_buy_good = book.add_format({"bg_color":"#C6E0B4","bold":True,"border":2,"border_color":"#000000","valign":"vcenter"})
        fmt_body_cr = book.add_format({"border":2, "border_color":"#000000", "num_format":"#,##0.00;[Red](#,##0.00);-", "valign":"vcenter"})
        fmt_good = book.add_format({"bg_color":"#DCFCE7","font_color":"#166534","bold":True,"border":2,"border_color":"#000000"})
        fmt_imp = book.add_format({"bg_color":"#DBEAFE","font_color":"#1D4ED8","bold":True,"border":2,"border_color":"#000000"})
        fmt_weak = book.add_format({"bg_color":"#FEF3C7","font_color":"#92400E","bold":True,"border":2,"border_color":"#000000"})
        fmt_lag = book.add_format({"bg_color":"#FEE2E2","font_color":"#991B1B","bold":True,"border":2,"border_color":"#000000"})

        # ---------- dashboard sheet ----------
        ws = writer.sheets["DASHBOARD"]
        ws.hide_gridlines(2)
        ws.set_zoom(90)
        ws.set_default_row(22)
        # Positive-side dashboard is intentionally wide but plain-English.
        ws.set_column("A:A", 11)   # Buy Priority
        ws.set_column("B:B", 15)   # ETF
        ws.set_column("C:C", 24)   # Sector / Theme
        ws.set_column("D:D", 14)   # Sector RRG
        ws.set_column("E:E", 12)   # Sector Rank
        ws.set_column("F:F", 11)   # Today %
        ws.set_column("G:G", 15)   # Today Volume
        ws.set_column("H:H", 15)   # Volume vs 30D
        ws.set_column("I:I", 13)   # Turnover
        ws.set_column("J:J", 14)   # Tradability
        ws.set_column("K:K", 16)   # Near day high
        ws.set_column("L:L", 13)   # ETF RRG
        ws.set_column("M:M", 11)   # Trend
        ws.set_column("N:N", 21)   # Momentum phase
        ws.set_column("O:O", 16)   # Entry stretch
        ws.set_column("P:P", 13)   # Buy Quality
        ws.set_column("Q:Q", 42)   # Buy Signal / Why Not Qualified
        ws.merge_range("A1:Q2", "NSE ETF RRG — Equity Swing Buy Dashboard", fmt_title)
        ws.merge_range("A3:Q3", f"Build V4.8.3 | Analysis Time: {ANALYSIS_TIME_IST} | Market Mode: {CURRENT_SESSION_MODE}", fmt_sub)
        ws.merge_range("A4:Q4", f"NSE Snapshot: {NSE_DOWNLOAD_TIMESTAMP} | Intraday: {INTRADAY_SCAN_WINDOW} | Benchmark: {BENCHMARK_LABEL} | Group 1: Equity + International Equity | Group 2: Gold / Silver / Debt / Liquid", fmt_sub)

        cards = [
            ("NSE Master ETFs", len(master_all)),
            ("NSE Vol >= 1L", len(eligible)),
            ("Leading Groups", qcount(eq_theme,"LEADING")),
            ("Improving Groups", qcount(eq_theme,"IMPROVING")),
            ("Swing Buy ETFs", len(all_latest_gainers)),
            ("Intraday Buy ETFs", len(intraday_buys)),
            ("Excluded < 1L Vol", len(master_excluded_low_volume)),
            ("Rebound Watch Rows", len(all_latest_losers)),
            ("Strict Rebound Buys", len(dip_candidates)),
            ("Deep Fall >=1%", DEEP_FALL_COUNT),
        ]
        card_cols = [0,2,4,6,8,0]
        card_rows = [4,4,4,4,4,7]
        # place six cards in 2 rows
        positions = [(5,0),(5,2),(5,4),(5,6),(5,8),(8,0)]
        positions += [(8,2),(8,4),(8,6),(8,8)]
        for i,(label,val) in enumerate(cards):
            r,c = positions[i]
            ws.merge_range(r, c, r, c+1, label, fmt_card_head)
            ws.merge_range(r+1, c, r+1, c+1, val, fmt_card_val)

        # Positive side begins after KPI cards.
        ws.set_row(10, 8)

        # ---- TOP 5 SECTORS / THEMES ----
        sector_row = 11
        ws.set_row(sector_row, 34)
        ws.merge_range(
            sector_row, 0, sector_row, 16,
            "TOP 5 EQUITY THEMES / MARKET GROUPS — LEADING / IMPROVING ONLY",
            fmt_header
        )

        if not dashboard_top_sectors.empty:
            for j, col in enumerate(dashboard_top_sectors.columns):
                ws.write(sector_row+1, j, col, fmt_header)
            for i, row in dashboard_top_sectors.reset_index(drop=True).iterrows():
                ws.set_row(sector_row+2+i, 24)
                for j, col in enumerate(dashboard_top_sectors.columns):
                    val = "" if pd.isna(row[col]) else row[col]
                    if col in ["Sector / Theme","Sector RRG"]:
                        q = str(row.get("Sector RRG",""))
                        cf = fmt_sector_leading if q == "LEADING" else fmt_sector_improving
                    elif col in ["Today Sector %","Breadth %"]:
                        cf = fmt_body_pct
                    elif col == "Sector Rank":
                        cf = fmt_body_int
                    elif isinstance(val,(int,float,np.integer,np.floating)) and not isinstance(val,bool):
                        cf = fmt_body_num
                    else:
                        cf = fmt_body
                    ws.write(sector_row+2+i, j, val, cf)
        else:
            ws.merge_range(
                sector_row+1,0,sector_row+1,16,
                "No LEADING / IMPROVING Equity theme / market group qualified at this analysis time.",
                fmt_body
            )

        # ---- STRICT BUY PRIORITY ----
        buying_row = sector_row + max(len(dashboard_top_sectors)+3, 8)
        ws.set_row(buying_row, 36)
        ws.merge_range(
            buying_row, 0, buying_row, 16,
            "GROUP 1 — EQUITY SWING BUY PRIORITY | Rank 1 = Best Strict Qualified Candidate Now",
            fmt_header
        )

        if not dashboard_buying.empty:
            for j,col in enumerate(dashboard_buying.columns):
                ws.write(buying_row+1, j, col, fmt_header)

            for i in range(len(dashboard_buying)):
                for j in range(17):
                    ws.write_blank(buying_row+2+i, j, None, fmt_grid_blank)

            for i,row in dashboard_buying.reset_index(drop=True).iterrows():
                ws.set_row(buying_row+2+i, 25)
                for j,col in enumerate(dashboard_buying.columns):
                    val = "" if pd.isna(row[col]) else row[col]

                    if col == "Buy Signal":
                        sval = str(val)
                        if "TOP BUY SETUP" in sval:
                            cf = fmt_buy_top
                        elif "STRONG BUY SETUP" in sval:
                            cf = fmt_buy_strong
                        else:
                            cf = fmt_buy_good

                    elif col in ["Sector / Theme","Sector RRG"]:
                        q = str(row.get("Sector RRG",""))
                        cf = fmt_sector_leading if q == "LEADING" else fmt_sector_improving

                    elif col == "Buy Priority":
                        cf = fmt_priority_1 if float(val) == 1 else fmt_body_int

                    elif col == "Momentum Phase":
                        sval = str(val)
                        if sval == "STRONG TURNAROUND":
                            cf = fmt_buy_top
                        elif sval == "FRESH MOMENTUM":
                            cf = fmt_buy_strong
                        elif sval == "HEALTHY TREND":
                            cf = fmt_buy_good
                        elif sval == "EARLY IMPROVEMENT":
                            cf = fmt_emerging_fill
                        elif sval == "FADING":
                            cf = fmt_buy_fading
                        else:
                            cf = fmt_buy_watch

                    elif col == "Entry Stretch":
                        sval = str(val)
                        if sval == "FRESH":
                            cf = fmt_buy_strong
                        elif sval == "HEALTHY":
                            cf = fmt_buy_good
                        elif sval == "EXTENDED":
                            cf = fmt_emerging_fill
                        else:
                            cf = fmt_body

                    elif col in ["Today %","Near Day High %"]:
                        cf = fmt_body_pct
                    elif col in ["Sector Rank","Trend /4"]:
                        cf = fmt_body_int

                    elif isinstance(val,(int,float,np.integer,np.floating)) and not isinstance(val,bool):
                        cf = fmt_body_num
                    else:
                        cf = fmt_body

                    ws.write(buying_row+2+i, j, val, cf)
        else:
            ws.merge_range(
                buying_row+1, 0, buying_row+2, 16,
                "NO QUALIFIED BUY NOW — No Equity / International Equity ETF passed every strict swing-entry condition at this analysis time.",
                fmt_buy_fading
            )

        # ---- NEAR BUY WATCHLIST ----
        near_row = buying_row + max(len(dashboard_buying) + 4, 7)
        ws.set_row(near_row, 36)
        ws.merge_range(
            near_row, 0, near_row, 16,
            "NEAR BUY WATCHLIST — Closest ETFs to Strict Qualification | NOT a Buy Signal",
            fmt_header
        )

        if not dashboard_near_buy.empty:
            for j,col in enumerate(dashboard_near_buy.columns):
                ws.write(near_row+1, j, col, fmt_header)

            for i in range(len(dashboard_near_buy)):
                for j in range(17):
                    ws.write_blank(near_row+2+i, j, None, fmt_grid_blank)

            for i,row in dashboard_near_buy.reset_index(drop=True).iterrows():
                ws.set_row(near_row+2+i, 34)
                for j,col in enumerate(dashboard_near_buy.columns):
                    val = "" if pd.isna(row[col]) else row[col]

                    if col in ["Sector / Theme","Sector RRG"]:
                        q = str(row.get("Sector RRG",""))
                        cf = fmt_sector_leading if q == "LEADING" else fmt_sector_improving
                    elif col == "Momentum Phase":
                        sval = str(val)
                        if sval == "STRONG TURNAROUND":
                            cf = fmt_buy_top
                        elif sval == "FRESH MOMENTUM":
                            cf = fmt_buy_strong
                        elif sval == "HEALTHY TREND":
                            cf = fmt_buy_good
                        elif sval == "EARLY IMPROVEMENT":
                            cf = fmt_emerging_fill
                        elif sval == "FADING":
                            cf = fmt_buy_fading
                        else:
                            cf = fmt_buy_watch
                    elif col == "Entry Stretch":
                        sval = str(val)
                        if sval == "FRESH":
                            cf = fmt_buy_strong
                        elif sval == "HEALTHY":
                            cf = fmt_buy_good
                        elif sval == "EXTENDED":
                            cf = fmt_emerging_fill
                        elif sval == "OVEREXTENDED":
                            cf = fmt_buy_fading
                        else:
                            cf = fmt_body
                    elif col == "Why Not Qualified":
                        cf = fmt_body_wrap
                    elif col in ["Today %","Near Day High %"]:
                        cf = fmt_body_pct
                    elif col in ["Watch Rank","Sector Rank","Trend /4","Failed Rules"]:
                        cf = fmt_body_int
                    elif isinstance(val,(int,float,np.integer,np.floating)) and not isinstance(val,bool):
                        cf = fmt_body_num
                    else:
                        cf = fmt_body
                    ws.write(near_row+2+i, j, val, cf)
        else:
            ws.merge_range(
                near_row+1,0,near_row+1,16,
                "No additional positive Equity / International Equity ETF is available for the near-buy watchlist.",
                fmt_body
            )

        # Intraday details are intentionally NOT repeated on the main Excel dashboard.
        # Use the single dedicated INTRADAY sheet for the compact intraday decision view.

        # ------------------------------------------------------------
        # GROUP 1 — NEGATIVE / REBOUND
        # Calculation logic is unchanged; this section is display-only refinement.
        # ------------------------------------------------------------
        losers_row = near_row + max(len(dashboard_near_buy) + 3, 5)
        # Signal is deliberately given three dashboard columns so full text is visible.
        neg_base_cols = len(dashboard_all_losers.columns) if not dashboard_all_losers.empty else 12
        neg_cols = neg_base_cols + 1  # zero-based last column after 3-column Signal merge
        ws.set_row(losers_row, 36)
        ws.set_row(losers_row+1, 42)
        ws.merge_range(
            losers_row, 0, losers_row, neg_cols,
            "GROUP 1 — EQUITY + INTERNATIONAL | REBOUND RANKING — Rank 1 = Best Available Candidate; Signal confirms BUY vs WATCH",
            fmt_header
        )

        if not dashboard_all_losers.empty:
            _signal_j = dashboard_all_losers.columns.get_loc("Signal")
            for j,col in enumerate(dashboard_all_losers.columns):
                if col == "Signal":
                    ws.merge_range(losers_row+1, j, losers_row+1, j+2, col, fmt_header)
                else:
                    ws.write(losers_row+1, j, col, fmt_header)

            for i,row in dashboard_all_losers.reset_index(drop=True).iterrows():
                ws.set_row(losers_row+2+i, 29)
                for j,col in enumerate(dashboard_all_losers.columns):
                    val = "" if pd.isna(row[col]) else row[col]

                    if col == "Rebound Rank":
                        cf = fmt_priority_1 if str(val) in ["1", "1.0"] else fmt_body_int

                    elif col in ["Group RRG","ETF RRG"]:
                        sval = str(val)
                        if sval == "LEADING":
                            cf = fmt_sector_leading
                        elif sval == "IMPROVING":
                            cf = fmt_sector_improving
                        elif sval == "WEAKENING":
                            cf = fmt_weak
                        else:
                            cf = fmt_body

                    elif col == "Signal":
                        sval = str(val)
                        if "TOP REBOUND BUY SETUP" in sval:
                            cf = fmt_dip_a
                        elif "STRONG REBOUND BUY SETUP" in sval:
                            cf = fmt_dip_strong
                        elif "GOOD REBOUND BUY SETUP" in sval:
                            cf = fmt_dip_volume
                        elif "REBOUND WATCH" in sval:
                            cf = fmt_dip_watch
                        elif "FALLING KNIFE" in sval or "WAIT / WEAK REBOUND" in sval:
                            cf = fmt_dip_low
                        else:
                            cf = fmt_body
                        # Give Signal a wide merged display area without changing the
                        # widths of columns used by the already-frozen positive sections.
                        ws.merge_range(losers_row+2+i, j, losers_row+2+i, j+2, val, cf)
                        continue

                    elif col in ["Today %","Recovery %","Vs Open %"]:
                        cf = fmt_body_pct
                    elif col == "Trend /4":
                        cf = fmt_body_int

                    elif isinstance(val,(int,float,np.integer,np.floating)) and not isinstance(val,bool):
                        cf = fmt_body_num
                    else:
                        cf = fmt_body

                    ws.write(losers_row+2+i, j, val, cf)
        else:
            ws.merge_range(
                losers_row+1, 0, losers_row+1, neg_cols,
                "No negative Equity / International ETF passed the current dynamic-liquidity universe.",
                fmt_body
            )

        # ------------------------------------------------------------
        # GROUP 1 — DEEP FALL VISIBILITY
        # Shows liquidity-qualified Equity / International ETFs down >=1% even
        # when rebound confirmation is not yet strong enough for Buy Priority.
        # ------------------------------------------------------------
        deep_row = losers_row + max(len(dashboard_all_losers) + 4, 7)
        deep_base_cols = len(dashboard_deep_fall.columns) if not dashboard_deep_fall.empty else 12
        deep_cols = deep_base_cols + 1  # room for 3-column Status merge
        ws.set_row(deep_row, 36)
        ws.set_row(deep_row+1, 42)
        ws.merge_range(
            deep_row, 0, deep_row, deep_cols,
            "GROUP 1 — EQUITY + INTERNATIONAL | DEEP FALL WATCH — Liquidity-Qualified ETFs Down 1% or More",
            fmt_header
        )

        if not dashboard_deep_fall.empty:
            for j,col in enumerate(dashboard_deep_fall.columns):
                if col == "Status":
                    ws.merge_range(deep_row+1, j, deep_row+1, j+2, col, fmt_header)
                else:
                    ws.write(deep_row+1, j, col, fmt_header)

            for i,row in dashboard_deep_fall.reset_index(drop=True).iterrows():
                ws.set_row(deep_row+2+i, 29)
                for j,col in enumerate(dashboard_deep_fall.columns):
                    val = "" if pd.isna(row[col]) else row[col]

                    if col in ["Group RRG","ETF RRG"]:
                        sval = str(val)
                        if sval == "LEADING":
                            cf = fmt_sector_leading
                        elif sval == "IMPROVING":
                            cf = fmt_sector_improving
                        elif sval == "WEAKENING":
                            cf = fmt_weak
                        else:
                            cf = fmt_body

                    elif col == "Status":
                        sval = str(val)
                        if "TOP REBOUND BUY SETUP" in sval:
                            cf = fmt_dip_a
                        elif "STRONG REBOUND BUY SETUP" in sval:
                            cf = fmt_dip_strong
                        elif "GOOD REBOUND BUY SETUP" in sval:
                            cf = fmt_dip_volume
                        elif "REBOUND WATCH" in sval:
                            cf = fmt_dip_watch
                        elif "FALLING KNIFE" in sval or "WAIT / WEAK REBOUND" in sval:
                            cf = fmt_dip_low
                        else:
                            cf = fmt_body
                        ws.merge_range(deep_row+2+i, j, deep_row+2+i, j+2, val, cf)
                        continue

                    elif col == "Today %":
                        cf = fmt_negative_pct

                    elif col in ["Recovery %","Vs Open %"]:
                        cf = fmt_body_pct
                    elif col in ["Fall Rank","Trend /4"]:
                        cf = fmt_body_int

                    elif isinstance(val,(int,float,np.integer,np.floating)) and not isinstance(val,bool):
                        cf = fmt_body_num
                    else:
                        cf = fmt_body

                    ws.write(deep_row+2+i, j, val, cf)
        else:
            ws.merge_range(
                deep_row+1, 0, deep_row+1, deep_cols,
                "No liquidity-qualified Equity / International ETF closed down 1% or more.",
                fmt_body
            )

        # Clear separation from Group 2.
        note_row = deep_row + max(len(dashboard_deep_fall) + 3, 4)
        ws.merge_range(
            note_row, 0, note_row, deep_cols,
            "GROUP 2 — GOLD / SILVER / DEBT / LIQUID remain separate in ALTERNATIVE_ASSETS and never compete for Equity / International rebound priority.",
            fmt_note
        )

        # ------------------------------------------------------------
        # TOP 10 UNUSUAL VOLUME — GROUP 1 ONLY
        # Wide merged layout prevents clipped flags and #### volume cells.
        # ------------------------------------------------------------
        volume_row = note_row + 2
        ws.set_row(volume_row, 36)
        ws.set_row(volume_row+1, 42)
        ws.merge_range(
            volume_row, 0, volume_row, 15,
            "TOP 10 UNUSUAL VOLUME — GROUP 1 EQUITY + INTERNATIONAL | Current NSE Volume vs Prior 30D Average",
            fmt_header
        )

        # Physical layout:
        # A Rank | B ETF | C:D Theme | E Asset | F:G Today Volume |
        # H:I Avg30 Volume | J:K Vol vs30D | L:N Flag | O Today% | P Turnover
        _uv_headers = [
            ("Volume Rank",0,0), ("ETF",1,1), ("Theme",2,3),
            ("Asset Class",4,4), ("Today Volume",5,6),
            ("Avg 30D Volume",7,8), ("Vol vs 30D %",9,10),
            ("Volume Flag",11,13), ("Today %",14,14), ("Turnover Cr",15,15)
        ]
        for _label,_c1,_c2 in _uv_headers:
            if _c1 == _c2:
                ws.write(volume_row+1, _c1, _label, fmt_header)
            else:
                ws.merge_range(volume_row+1, _c1, volume_row+1, _c2, _label, fmt_header)

        if not dashboard_volume_surges.empty:
            for _i,_row in dashboard_volume_surges.reset_index(drop=True).iterrows():
                _rr = volume_row + 2 + _i
                ws.set_row(_rr, 26)

                _rank = _row.get("Volume Rank", "")
                _etf = _row.get("ETF", "")
                _theme = _row.get("Theme", "")
                _asset = _row.get("Asset Class", "")
                _dayvol = _row.get("Today Volume", "")
                _avgvol = _row.get("Avg 30D Volume", "")
                _volpct = _row.get("Vol vs 30D %", "")
                _flag = str(_row.get("Volume Flag", ""))
                _chg = _row.get("Today %", "")
                _turn = _row.get("Turnover Cr", "")

                if "EXTREME" in _flag:
                    _vf = fmt_vol_extreme
                elif "VERY HIGH" in _flag:
                    _vf = fmt_vol_very_high
                elif "HIGH" in _flag:
                    _vf = fmt_vol_high
                elif "ABOVE" in _flag:
                    _vf = fmt_vol_above
                else:
                    _vf = fmt_body

                if pd.notna(_chg) and _chg < 0:
                    _cf = fmt_negative_pct
                elif pd.notna(_chg) and _chg > 0:
                    _cf = fmt_positive_pct
                else:
                    _cf = fmt_body_pct

                ws.write(_rr, 0, _rank, fmt_body_int)
                ws.write(_rr, 1, _etf, fmt_body)
                ws.merge_range(_rr, 2, _rr, 3, _theme, fmt_body)
                ws.write(_rr, 4, _asset, fmt_body)
                ws.merge_range(_rr, 5, _rr, 6, _dayvol, fmt_body_int)
                ws.merge_range(_rr, 7, _rr, 8, _avgvol, fmt_body_int)
                ws.merge_range(_rr, 9, _rr, 10, _volpct, _vf)
                ws.merge_range(_rr, 11, _rr, 13, _flag, _vf)
                ws.write(_rr, 14, _chg, _cf)
                ws.write(_rr, 15, _turn, fmt_body_num)
        else:
            ws.merge_range(
                volume_row+2, 0, volume_row+2, 15,
                "No usable unusual-volume data for Group 1 Equity / International.",
                fmt_body
            )

        # Chart source data goes to a hidden worksheet, not the visible dashboard.
        chart_ws = writer.sheets["_CHART_DATA"]
        chart_ws.hide()

        # Top 10 theme/category turnover
        chart_ws.write(0, 0, "Theme")
        chart_ws.write(0, 1, "LatestSessionTurnoverCr")
        for i, row in turnover_by_theme.reset_index(drop=True).iterrows():
            chart_ws.write(i+1, 0, row["Theme"])
            chart_ws.write(i+1, 1, row["LatestSessionTurnoverCr"])

        # Top 10 ETF turnover
        chart_ws.write(0, 3, "Symbol")
        chart_ws.write(0, 4, "TurnoverCr")
        for i, row in top_etf_turnover.reset_index(drop=True).iterrows():
            chart_ws.write(i+1, 3, row["Symbol"])
            chart_ws.write(i+1, 4, row["TodayTurnoverCr"])

        # Top 10 unusual volume
        chart_ws.write(0, 6, "VolumeSurgeETF")
        chart_ws.write(0, 7, "VolumeVs30DAvgPct")
        for i, row in top_volume_surges.reset_index(drop=True).iterrows():
            chart_ws.write(i+1, 6, row["Symbol"])
            chart_ws.write(i+1, 7, row["VolVs30DAvgPct"])

        # Pie chart: top 10 traded ETF categories/themes
        pie = book.add_chart({"type":"pie"})
        pie.add_series({
            "name":"Top 10 ETF Categories by Latest Session Turnover",
            "categories":["_CHART_DATA",1,0,max(1,len(turnover_by_theme)),0],
            "values":["_CHART_DATA",1,1,max(1,len(turnover_by_theme)),1],
            "data_labels":{
                "percentage": True,
                "category": True,
                "position": "best_fit",
                "leader_lines": True,
                "num_format": "0.00%"
            }
        })
        pie.set_title({"name":"Top 10 ETF Categories by Latest Session Turnover"})
        pie.set_legend({"position":"right"})
        ws.insert_chart("S2", pie, {"x_scale": 1.35, "y_scale": 1.35})

        # Bar chart: top 10 traded ETFs
        bar = book.add_chart({"type":"bar"})
        bar.add_series({
            "name":"Top 10 ETFs by Latest Session Turnover",
            "categories":["_CHART_DATA",1,3,max(1,len(top_etf_turnover)),3],
            "values":["_CHART_DATA",1,4,max(1,len(top_etf_turnover)),4],
            "data_labels":{"value": True, "num_format":"0.00"}
        })
        bar.set_title({"name":"Top 10 ETFs by Latest Session Turnover"})
        bar.set_x_axis({"name":"Turnover (Crore)", "num_format":"0.00"})
        bar.set_y_axis({"reverse": True})
        ws.insert_chart("S20", bar, {"x_scale": 1.2, "y_scale": 1.25})

        # Unusual volume chart
        vol_chart = book.add_chart({"type":"bar"})
        vol_chart.add_series({
            "name":"Current Volume vs Prior 30D Average %",
            "categories":["_CHART_DATA",1,6,max(1,len(top_volume_surges)),6],
            "values":["_CHART_DATA",1,7,max(1,len(top_volume_surges)),7],
            "data_labels":{"value": True, "num_format":"0.00"}
        })
        vol_chart.set_title({"name":"Top 10 Unusual Volume ETFs vs Prior 30D Average"})
        vol_chart.set_x_axis({"name":"Volume above 30D average (%)", "num_format":"0.00"})
        vol_chart.set_y_axis({"reverse": True})
        vol_chart.set_legend({"none": True})
        ws.insert_chart("S38", vol_chart, {"x_scale": 1.2, "y_scale": 1.25})

        # ---------- professional sizing + border helpers ----------
        def excel_width(series, header, min_w=10, max_w=50):
            vals = [str(header)]
            try:
                vals += [str(x) for x in series.dropna().tolist()]
            except Exception:
                pass
            longest = max([len(x) for x in vals] or [min_w])
            return max(min_w, min(max_w, longest + 3))

        def write_table_with_professional_format(ws, dfw):
            if dfw is None:
                return
            # Header
            for ci, col in enumerate(dfw.columns):
                ws.write(0, ci, col, fmt_header)

            # Body — rewrite values so every used cell has a clean grid/border.
            for ri, row in dfw.reset_index(drop=True).iterrows():
                for ci, col in enumerate(dfw.columns):
                    val = row[col]
                    if pd.isna(val):
                        val = ""
                    if col in ["SecurityName","Underlying","Theme","Sector / Theme","Equity Theme / Market Group","Signal","DipSignal","Buy Signal","Intraday Signal","Momentum Phase","Entry Stretch","Status","Rule","Reason","Why Not Qualified"]:
                        cell_fmt = fmt_body_wrap
                    elif col in ["NSE_ChangePct","Today %","Today Sector %","Today Group %","Vs Open %","Breadth %","RangePositionPct","Near Day High %","CurrentVsAvg30VolumePct","DayReturnPct","Week1ReturnPct","Week2ReturnPct","Month1ReturnPct","Month3ReturnPct","Month6ReturnPct","Year1ReturnPct","PctFrom52WHigh","PctAbove52WLow"]:
                        cell_fmt = fmt_body_pct
                    elif col in [
                        "AltRank","ThemeRank","TradeRank","RankOverall","RankInTheme",
                        "Sector Rank","Buy Priority","Watch Rank","Rebound Rank",
                        "Fall Rank","Group Rank","Priority","Trend /4","TrendPoints",
                        "Failed Rules","HistoryDays"
                    ]:
                        cell_fmt = fmt_body_int
                    elif col in ["TodayTurnoverCr","LatestSessionTurnoverCr","Avg30TurnoverCr","ADTV5Cr","ADTV10Cr","ADTV20Cr","ADTV30Cr","LTPxAvg5VolCr","LTPxAvg10VolCr","LTPxAvg20VolCr","Turnover Cr","30D Avg Turnover Cr"]:
                        cell_fmt = fmt_body_num
                    elif isinstance(val, (int, float, np.integer, np.floating)) and not isinstance(val, bool):
                        cell_fmt = fmt_body_num
                    else:
                        cell_fmt = fmt_body
                    ws.write(ri + 1, ci, val, cell_fmt)

            # Auto-size each column with sensible caps.
            for ci, col in enumerate(dfw.columns):
                if col in ["SecurityName","Underlying"]:
                    width = excel_width(dfw[col], col, 18, 50)
                elif col in ["Theme","Signal","DipSignal","Status","Rule"]:
                    width = excel_width(dfw[col], col, 14, 40)
                elif col in ["Why Not Qualified","WhyNotQualified"]:
                    width = excel_width(dfw[col], col, 24, 48)
                elif "Timestamp" in col or col == "NSE_AsOf":
                    width = 30
                else:
                    width = excel_width(dfw[col], col, 10, 32)
                # Ensure long analytical headers are fully visible.
                if col in ["CurrentVsAvg30VolumePct","LatestSessionTurnoverCr","Avg30TurnoverCr",
                           "Theme_RS_Momentum","Theme_RS_Ratio","NSE_DownloadTimestamp"]:
                    width = max(width, 24)
                ws.set_column(ci, ci, width)

            # Adjust all rows to readable heights.
            ws.set_row(0, 30)
            for ri in range(1, len(dfw) + 1):
                row_height = 20
                try:
                    text_len = max(len(str(x)) for x in dfw.iloc[ri-1].tolist() if pd.notna(x))
                    if text_len > 70:
                        row_height = 42
                    elif text_len > 40:
                        row_height = 30
                    elif text_len > 24:
                        row_height = 24
                except Exception:
                    pass
                ws.set_row(ri, row_height)

        # ---------- compact single INTRADAY worksheet ----------
        intraday_ws = writer.sheets["INTRADAY"]
        intraday_ws.hide_gridlines(2)
        intraday_ws.set_zoom(90)
        intraday_ws.freeze_panes(3, 0)

        # Use only the columns actually needed by the visible intraday tables.
        # This removes stray border/line fragments in unused columns when the
        # strict-buy block is empty after market close.
        _intraday_visible_cols = max(
            len(intraday_group_compact.columns),
            len(intraday_near_compact.columns),
            len(intraday_buy_compact.columns) if not intraday_buy_compact.empty else 0,
            8
        )
        _intraday_last_col = _intraday_visible_cols - 1

        # Sensible compact widths. Detailed scores remain in the background.
        intraday_ws.set_column("A:A", 11)
        intraday_ws.set_column("B:B", 14)
        intraday_ws.set_column("C:C", 24)
        intraday_ws.set_column("D:D", 13)
        intraday_ws.set_column("E:G", 13)
        intraday_ws.set_column("H:H", 15)
        intraday_ws.set_column("I:I", 16)
        intraday_ws.set_column("J:J", 13)
        intraday_ws.set_column("K:K", 15)
        intraday_ws.set_column("L:L", 22)

        # Closed-market compact view normally needs only A:H.
        # Hide any unused columns to the right; they automatically remain visible
        # during a live run when the strict-buy table requires them.
        if _intraday_visible_cols < 12:
            intraday_ws.set_column(
                _intraday_visible_cols, 11, None, None, {"hidden": True}
            )

        intraday_ws.set_row(0, 30)
        intraday_ws.merge_range(0, 0, 0, _intraday_last_col, "NSE ETF RRG — INTRADAY POSITIVE SCANNER", fmt_title)
        intraday_ws.set_row(1, 24)
        intraday_ws.merge_range(
            1, 0, 1, _intraday_last_col,
            f"Build V4.8.3 | Analysis: {ANALYSIS_TIME_IST} | Market Mode: {CURRENT_SESSION_MODE} | Scan Window: {INTRADAY_SCAN_WINDOW} | Universe: TOP 50 ACTIVE EQUITY | NSE Snapshot: {NSE_DOWNLOAD_TIMESTAMP}",
            fmt_sub
        )

        # Section 1: only the five strongest eligible equity groups.
        grp_row = 3
        intraday_ws.merge_range(grp_row, 0, grp_row, _intraday_last_col, "TOP 5 INTRADAY EQUITY GROUPS", fmt_header)
        grp_header_row = grp_row + 1
        for j, col in enumerate(intraday_group_compact.columns):
            intraday_ws.write(grp_header_row, j, col, fmt_header)
        for i, row in intraday_group_compact.reset_index(drop=True).iterrows():
            rr = grp_header_row + 1 + i
            intraday_ws.set_row(rr, 22)
            for j, col in enumerate(intraday_group_compact.columns):
                val = "" if pd.isna(row[col]) else row[col]
                if col in ["Equity Group", "Group RRG"]:
                    q = str(row.get("Group RRG", ""))
                    cf = fmt_sector_leading if q == "LEADING" else fmt_sector_improving
                elif col == "Group Rank":
                    cf = fmt_body_int
                elif col in ["Today %", "Breadth %"]:
                    cf = fmt_body_pct
                elif isinstance(val, (int,float,np.integer,np.floating)) and not isinstance(val,bool):
                    cf = fmt_body_num
                else:
                    cf = fmt_body
                intraday_ws.write(rr, j, val, cf)

        # Section 2: strict intraday buys. Maximum five displayed.
        buy_row = grp_header_row + max(len(intraday_group_compact), 1) + 3
        intraday_ws.merge_range(
            buy_row, 0, buy_row, _intraday_last_col,
            "INTRADAY BUY PRIORITY — ONLY STRICT QUALIFIED CANDIDATES",
            fmt_header
        )
        if not intraday_buy_compact.empty:
            buy_header_row = buy_row + 1
            for j, col in enumerate(intraday_buy_compact.columns):
                intraday_ws.write(buy_header_row, j, col, fmt_header)
            for i, row in intraday_buy_compact.reset_index(drop=True).iterrows():
                rr = buy_header_row + 1 + i
                intraday_ws.set_row(rr, 23)
                for j, col in enumerate(intraday_buy_compact.columns):
                    val = "" if pd.isna(row[col]) else row[col]
                    if col == "Priority":
                        cf = fmt_priority_1 if float(val) == 1 else fmt_body_int
                    elif col == "Signal":
                        sval = str(val)
                        if "TOP INTRADAY" in sval:
                            cf = fmt_buy_top
                        elif "STRONG INTRADAY" in sval:
                            cf = fmt_buy_strong
                        else:
                            cf = fmt_buy_good
                    elif col in ["Equity Group", "Group RRG"]:
                        q = str(row.get("Group RRG", ""))
                        cf = fmt_sector_leading if q == "LEADING" else fmt_sector_improving
                    elif col in ["Today %", "Vs Open %", "Near Day High %"]:
                        cf = fmt_body_pct
                    elif isinstance(val, (int,float,np.integer,np.floating)) and not isinstance(val,bool):
                        cf = fmt_body_num
                    else:
                        cf = fmt_body
                    intraday_ws.write(rr, j, val, cf)
            buy_end_row = buy_header_row + len(intraday_buy_compact)
        else:
            intraday_ws.merge_range(
                buy_row+1, 0, buy_row+2, _intraday_last_col,
                f"NO QUALIFIED INTRADAY BUY NOW — {INTRADAY_SCAN_WINDOW}",
                fmt_buy_fading
            )
            buy_end_row = buy_row + 2

        # Section 3: only the closest five alternatives, with the exact reason.
        watch_row = buy_end_row + 2
        intraday_ws.merge_range(
            watch_row, 0, watch_row, _intraday_last_col,
            "NEAR BUY WATCHLIST — NOT A BUY SIGNAL",
            fmt_header
        )
        if not intraday_near_compact.empty:
            watch_header_row = watch_row + 1
            for j, col in enumerate(intraday_near_compact.columns):
                intraday_ws.write(watch_header_row, j, col, fmt_header)
            for i, row in intraday_near_compact.reset_index(drop=True).iterrows():
                rr = watch_header_row + 1 + i
                intraday_ws.set_row(rr, 34)
                for j, col in enumerate(intraday_near_compact.columns):
                    val = "" if pd.isna(row[col]) else row[col]
                    if col == "Why Not Qualified":
                        cf = fmt_body_wrap
                    elif col == "Watch Rank":
                        cf = fmt_body_int
                    elif col in ["Today %", "Near Day High %"]:
                        cf = fmt_body_pct
                    elif isinstance(val, (int,float,np.integer,np.floating)) and not isinstance(val,bool):
                        cf = fmt_body_num
                    else:
                        cf = fmt_body
                    intraday_ws.write(rr, j, val, cf)
            reason_col = intraday_near_compact.columns.get_loc("Why Not Qualified") if "Why Not Qualified" in intraday_near_compact.columns else None
            if reason_col is not None:
                intraday_ws.set_column(reason_col, reason_col, 48)
        else:
            intraday_ws.merge_range(watch_row+1, 0, watch_row+1, _intraday_last_col, "No near-buy intraday candidate available.", fmt_body)

        intraday_ws.set_row(watch_row + max(len(intraday_near_compact),1) + 3, 30)
        note_row = watch_row + max(len(intraday_near_compact),1) + 3
        intraday_ws.merge_range(
            note_row, 0, note_row, _intraday_last_col,
            f"Background: Intraday universe = Top {INTRADAY_ACTIVITY_COUNT} active Equity / International ETFs (50% turnover + 30% volume + 20% volume pace). Swing remains on >=1 lakh current volume; Negative / Dip uses its separate dynamic-liquidity universe. Only decision fields are displayed here.",
            fmt_note
        )

        # ---------- standard sheets formatting ----------
        df_map = {
            "ETF_ANALYSIS": analysis_full,
            "TOP_5_SECTORS": top5_sectors_sheet,
            "THEME_RRG": theme_rrg_valid,
            "TRADE_CANDIDATES": trade_candidates_sheet,
            "BUYING_ANALYSIS": buying_sheet,
            "NEAR_BUY_WATCHLIST": near_buy_sheet,
            "LOSERS_DIP_ANALYSIS": all_losers_sheet,
            "ALTERNATIVE_ASSETS": alternative_assets,
            "INFO": None,
        }

        for sname, dfw in df_map.items():
            ws = writer.sheets[sname]
            ws.hide_gridlines(2)
            ws.freeze_panes(1, 0)
            ws.set_zoom(90)

            if sname != "INFO":
                write_table_with_professional_format(ws, dfw)
                if len(dfw) > 0:
                    ws.autofilter(0, 0, len(dfw), len(dfw.columns)-1)
                    n = len(dfw) + 1

                    # RRG quadrant coloring
                    for qcol in ["Quadrant","Theme_Quadrant","CurrentQuadrant","PreviousQuadrant","Sector RRG","Group RRG","ETF RRG"]:
                        if qcol in dfw.columns:
                            c = dfw.columns.get_loc(qcol)
                            ws.conditional_format(1,c,n-1,c,{"type":"text","criteria":"containing","value":"LEADING","format":fmt_good})
                            ws.conditional_format(1,c,n-1,c,{"type":"text","criteria":"containing","value":"IMPROVING","format":fmt_imp})
                            ws.conditional_format(1,c,n-1,c,{"type":"text","criteria":"containing","value":"WEAKENING","format":fmt_weak})
                            ws.conditional_format(1,c,n-1,c,{"type":"text","criteria":"containing","value":"LAGGING","format":fmt_lag})

                    # Highlight unusual current volume vs the prior 30D average.
                    if "CurrentVsAvg30VolumePct" in dfw.columns:
                        c = dfw.columns.get_loc("CurrentVsAvg30VolumePct")
                        ws.conditional_format(1,c,n-1,c,{
                            "type":"3_color_scale",
                            "min_color":"#F8FAFC",
                            "mid_color":"#FDE68A",
                            "max_color":"#F97316"
                        })

                    # Eligibility is already guaranteed by the initial NSE >=1 lakh
                    # volume filter. Highlight only the eligibility cell so signal
                    # colors remain visible across BUYING/LOSERS sheets.
                    for ecol in ["BuyingEligibilityResult","DipEligibilityResult"]:
                        if ecol in dfw.columns:
                            ec = dfw.columns.get_loc(ecol)
                            ws.conditional_format(1,ec,n-1,ec,{
                                "type":"text",
                                "criteria":"containing",
                                "value":"ELIGIBLE",
                                "format":fmt_qualified_fill
                            })

                    # Directly rewrite signal cells with colors so the visible fill
                    # is guaranteed even if other conditional formatting applies.
                    for bscol in ["BuyingSignal","BuyingSignalFull"]:
                        if bscol in dfw.columns:
                            bc = dfw.columns.get_loc(bscol)
                            for ri, sval in enumerate(dfw[bscol].astype(str).tolist(), start=1):
                                if "A+ HEAVY BUYING" in sval:
                                    sf = fmt_buy_a
                                elif "STRONG BUYING" in sval:
                                    sf = fmt_buy_strong
                                elif "BUYING PRESSURE" in sval:
                                    sf = fmt_buy_pressure
                                elif "POSITIVE BUY WATCH" in sval:
                                    sf = fmt_buy_watch
                                elif "WEAK POSITIVE / FADING" in sval:
                                    sf = fmt_buy_fading
                                else:
                                    sf = fmt_body
                                ws.write(ri, bc, sval, sf)

                    # Positive Buy Priority color fields in BUYING_ANALYSIS.
                    if "Buy Signal" in dfw.columns:
                        bc = dfw.columns.get_loc("Buy Signal")
                        for ri, sval in enumerate(dfw["Buy Signal"].astype(str).tolist(), start=1):
                            if "TOP BUY SETUP" in sval:
                                sf = fmt_buy_top
                            elif "STRONG BUY SETUP" in sval:
                                sf = fmt_buy_strong
                            elif "GOOD BUY SETUP" in sval:
                                sf = fmt_buy_good
                            else:
                                sf = fmt_buy_watch
                            ws.write(ri, bc, sval, sf)

                    if "Intraday Signal" in dfw.columns:
                        ic = dfw.columns.get_loc("Intraday Signal")
                        for ri, sval in enumerate(dfw["Intraday Signal"].astype(str).tolist(), start=1):
                            if "TOP INTRADAY SETUP" in sval:
                                sf = fmt_buy_top
                            elif "STRONG INTRADAY SETUP" in sval:
                                sf = fmt_buy_strong
                            elif "GOOD INTRADAY SETUP" in sval:
                                sf = fmt_buy_good
                            else:
                                sf = fmt_buy_watch
                            ws.write(ri, ic, sval, sf)

                    if "Why Not Qualified" in dfw.columns:
                        wc = dfw.columns.get_loc("Why Not Qualified")
                        ws.set_column(wc, wc, 48, fmt_body_wrap)

                    if "Momentum Phase" in dfw.columns:
                        mc = dfw.columns.get_loc("Momentum Phase")
                        for ri, sval in enumerate(dfw["Momentum Phase"].astype(str).tolist(), start=1):
                            if sval == "STRONG TURNAROUND":
                                sf = fmt_buy_top
                            elif sval == "FRESH MOMENTUM":
                                sf = fmt_buy_strong
                            elif sval == "HEALTHY TREND":
                                sf = fmt_buy_good
                            elif sval == "EARLY IMPROVEMENT":
                                sf = fmt_emerging_fill
                            elif sval == "FADING":
                                sf = fmt_buy_fading
                            else:
                                sf = fmt_buy_watch
                            ws.write(ri, mc, sval, sf)

                    if "Entry Stretch" in dfw.columns:
                        ec2 = dfw.columns.get_loc("Entry Stretch")
                        for ri, sval in enumerate(dfw["Entry Stretch"].astype(str).tolist(), start=1):
                            if sval == "FRESH":
                                sf = fmt_buy_strong
                            elif sval == "HEALTHY":
                                sf = fmt_buy_good
                            elif sval == "EXTENDED":
                                sf = fmt_emerging_fill
                            elif sval == "OVEREXTENDED":
                                sf = fmt_buy_fading
                            else:
                                sf = fmt_body
                            ws.write(ri, ec2, sval, sf)

                    for dscol in ["DipSignal"]:
                        if dscol in dfw.columns:
                            dc = dfw.columns.get_loc(dscol)
                            for ri, sval in enumerate(dfw[dscol].astype(str).tolist(), start=1):
                                if "TOP REBOUND BUY SETUP" in sval:
                                    sf = fmt_dip_a
                                elif "STRONG REBOUND BUY SETUP" in sval:
                                    sf = fmt_dip_strong
                                elif "GOOD REBOUND BUY SETUP" in sval:
                                    sf = fmt_dip_volume
                                elif "REBOUND WATCH" in sval:
                                    sf = fmt_dip_watch
                                elif "FALLING KNIFE" in sval or "WAIT / WEAK REBOUND" in sval:
                                    sf = fmt_dip_low
                                else:
                                    sf = fmt_body
                                ws.write(ri, dc, sval, sf)

                    # Buying signal colors
                    for bscol in ["BuyingSignal","BuyingSignalFull"]:
                        if bscol in dfw.columns:
                            bc = dfw.columns.get_loc(bscol)
                            ws.conditional_format(1,bc,n-1,bc,{"type":"text","criteria":"containing","value":"A+ HEAVY BUYING","format":fmt_buy_a})
                            ws.conditional_format(1,bc,n-1,bc,{"type":"text","criteria":"containing","value":"STRONG BUYING","format":fmt_buy_strong})
                            ws.conditional_format(1,bc,n-1,bc,{"type":"text","criteria":"containing","value":"BUYING PRESSURE","format":fmt_buy_pressure})
                            ws.conditional_format(1,bc,n-1,bc,{"type":"text","criteria":"containing","value":"POSITIVE BUY WATCH","format":fmt_buy_watch})
                            ws.conditional_format(1,bc,n-1,bc,{"type":"text","criteria":"containing","value":"WEAK POSITIVE / FADING","format":fmt_buy_fading})

                    # Negative rebound signal colors
                    for _sig_col in ["DipSignal","Signal"]:
                        if _sig_col in dfw.columns:
                            dc = dfw.columns.get_loc(_sig_col)
                            ws.conditional_format(1,dc,n-1,dc,{"type":"text","criteria":"containing","value":"TOP REBOUND BUY SETUP","format":fmt_dip_a})
                            ws.conditional_format(1,dc,n-1,dc,{"type":"text","criteria":"containing","value":"STRONG REBOUND BUY SETUP","format":fmt_dip_strong})
                            ws.conditional_format(1,dc,n-1,dc,{"type":"text","criteria":"containing","value":"GOOD REBOUND BUY SETUP","format":fmt_dip_volume})
                            ws.conditional_format(1,dc,n-1,dc,{"type":"text","criteria":"containing","value":"REBOUND WATCH","format":fmt_dip_watch})
                            ws.conditional_format(1,dc,n-1,dc,{"type":"text","criteria":"containing","value":"FALLING KNIFE","format":fmt_dip_low})
                            ws.conditional_format(1,dc,n-1,dc,{"type":"text","criteria":"containing","value":"WAIT / WEAK REBOUND","format":fmt_dip_low})

                    # Score heatmaps
                    for scol in ["TradeScore","DipBuyScore","ReboundScore","Rebound Score","AggressiveDipScore","LatestBuyingScore","RS_Ratio","RS_Momentum","Theme_RS_Ratio","Theme_RS_Momentum"]:
                        if scol in dfw.columns:
                            c = dfw.columns.get_loc(scol)
                            ws.conditional_format(1,c,n-1,c,{"type":"3_color_scale","min_color":"#FCA5A5","mid_color":"#FDE68A","max_color":"#86EFAC"})

        # ------------------------------------------------------------
        # NEGATIVE SHEET — presentation upgrade only
        # ------------------------------------------------------------
        neg_ws = writer.sheets["LOSERS_DIP_ANALYSIS"]
        neg_ws.set_zoom(82)
        neg_ws.freeze_panes(1, 3)

        # Decision columns are intentionally wider / easier to scan.
        _neg_widths = {
            "Rebound Rank": 13,
            "Buy Priority": 12,
            "Watch Rank": 11,
            "ETF": 18,
            "Theme": 24,
            "Group RRG": 13,
            "ETF RRG": 13,
            "Today %": 11,
            "Recovery %": 12,
            "Vs Open %": 11,
            "Vol Pace": 11,
            "Turnover Cr": 13,
            "30D Avg Turnover Cr": 17,
            "Trend /4": 10,
            "Rebound Score": 14,
            "Qualified": 11,
            "Signal": 34,
            "Why Not Qualified": 44,
        }
        for _cname, _width in _neg_widths.items():
            if _cname in all_losers_sheet.columns:
                _ci = all_losers_sheet.columns.get_loc(_cname)
                neg_ws.set_column(_ci, _ci, _width)

        # Strong visual cue for strict Priority 1.
        if "Buy Priority" in all_losers_sheet.columns and len(all_losers_sheet) > 0:
            _pc = all_losers_sheet.columns.get_loc("Buy Priority")
            neg_ws.conditional_format(
                1, _pc, len(all_losers_sheet), _pc,
                {"type":"cell","criteria":"==","value":1,"format":fmt_priority_1}
            )

        # Append the complete liquidity-qualified >=1% Deep Fall Watch
        # below the existing rebound table, on the SAME worksheet.
        deep_sheet_start = len(all_losers_sheet) + 4
        deep_sheet = deep_fall_watch[[
            c for c in [
                "DeepFallRank","Symbol","Theme","Theme_Quadrant","Quadrant",
                "NSE_ChangePct","RangePositionPct","ReboundVsOpenPct",
                "CurrentVolumeMultiple30D","TodayTurnoverCr","Avg30TurnoverCr",
                "TrendPoints","ReboundScore","DipSignal",
                "WhyNotReboundQualified"
            ] if c in deep_fall_watch.columns
        ]].copy().rename(columns={
            "DeepFallRank":"Fall Rank",
            "Symbol":"ETF",
            "Theme":"Theme",
            "Theme_Quadrant":"Group RRG",
            "Quadrant":"ETF RRG",
            "NSE_ChangePct":"Today %",
            "RangePositionPct":"Recovery %",
            "ReboundVsOpenPct":"Vs Open %",
            "CurrentVolumeMultiple30D":"Vol Pace",
            "TodayTurnoverCr":"Turnover Cr",
            "Avg30TurnoverCr":"30D Avg Turnover Cr",
            "TrendPoints":"Trend /4",
            "ReboundScore":"Rebound Score",
            "DipSignal":"Status",
            "WhyNotReboundQualified":"Why Not Qualified"
        })

        _deep_end_col = max(0, len(deep_sheet.columns) - 1)
        neg_ws.set_row(deep_sheet_start, 34)
        neg_ws.merge_range(
            deep_sheet_start, 0, deep_sheet_start, max(_deep_end_col, 10),
            "GROUP 1 — EQUITY + INTERNATIONAL | DEEP FALL WATCH — Liquidity-Qualified ETFs Down 1% or More",
            fmt_header
        )

        if not deep_sheet.empty:
            for _j, _col in enumerate(deep_sheet.columns):
                neg_ws.write(deep_sheet_start+1, _j, _col, fmt_header)

            for _i, _row in deep_sheet.reset_index(drop=True).iterrows():
                _rr = deep_sheet_start + 2 + _i
                neg_ws.set_row(_rr, 26)
                for _j, _col in enumerate(deep_sheet.columns):
                    _val = "" if pd.isna(_row[_col]) else _row[_col]

                    if _col in ["Group RRG","ETF RRG"]:
                        _sv = str(_val)
                        if _sv == "LEADING":
                            _fmt = fmt_sector_leading
                        elif _sv == "IMPROVING":
                            _fmt = fmt_sector_improving
                        elif _sv == "WEAKENING":
                            _fmt = fmt_weak
                        else:
                            _fmt = fmt_body

                    elif _col == "Status":
                        _sv = str(_val)
                        if "TOP REBOUND BUY SETUP" in _sv:
                            _fmt = fmt_dip_a
                        elif "STRONG REBOUND BUY SETUP" in _sv:
                            _fmt = fmt_dip_strong
                        elif "GOOD REBOUND BUY SETUP" in _sv:
                            _fmt = fmt_dip_volume
                        elif "REBOUND WATCH" in _sv:
                            _fmt = fmt_dip_watch
                        elif "FALLING KNIFE" in _sv or "WAIT / WEAK REBOUND" in _sv:
                            _fmt = fmt_dip_low
                        else:
                            _fmt = fmt_body

                    elif _col == "Today %":
                        _fmt = fmt_negative_pct

                    elif _col in ["Recovery %","Vs Open %"]:
                        _fmt = fmt_body_pct
                    elif _col in ["Fall Rank","Trend /4"]:
                        _fmt = fmt_body_int

                    elif isinstance(_val, (int,float,np.integer,np.floating)) and not isinstance(_val, bool):
                        _fmt = fmt_body_num
                    else:
                        _fmt = fmt_body_wrap if _col == "Why Not Qualified" else fmt_body

                    neg_ws.write(_rr, _j, _val, _fmt)

            # Column widths for the appended Deep Fall table.
            for _j, _col in enumerate(deep_sheet.columns):
                _w = {
                    "Fall Rank":10, "ETF":18, "Theme":24,
                    "Group RRG":13, "ETF RRG":13,
                    "Today %":11, "Recovery %":12, "Vs Open %":11,
                    "Vol Pace":11, "Turnover Cr":13, "30D Avg Turnover Cr":17,
                    "Trend /4":10, "Rebound Score":14,
                    "Status":34, "Why Not Qualified":44
                }.get(_col, 14)
                neg_ws.set_column(_j, _j, _w)

        else:
            neg_ws.merge_range(
                deep_sheet_start+1, 0, deep_sheet_start+1, 10,
                "No liquidity-qualified Equity / International ETF closed down 1% or more.",
                fmt_body
            )

        _group2_note_row = deep_sheet_start + max(len(deep_sheet) + 3, 4)
        neg_ws.merge_range(
            _group2_note_row, 0, _group2_note_row, max(_deep_end_col, 10),
            "GROUP 2 — Gold / Silver / Debt / Liquid remain separate in ALTERNATIVE_ASSETS and never compete for Group 1 rebound priority.",
            fmt_note
        )

        # INFO sheet formatting and layout
        ws = writer.sheets["INFO"]
        ws.hide_gridlines(2)
        ws.set_zoom(95)
        ws.set_column("A:A", 30)
        ws.set_column("B:B", 90)
        ws.set_row(0, 30)
        ws.write(0,0,"PROJECT INFO", fmt_header)

        # First table
        ws.write(2,0,"Item", fmt_header)
        ws.write(2,1,"Value", fmt_header)
        for i, row in readme.reset_index(drop=True).iterrows():
            ws.write(i+3, 0, row["Item"], fmt_body)
            ws.write(i+3, 1, row["Value"], fmt_body_wrap)
            _info_len = len(str(row["Value"]))
            ws.set_row(
                i+3,
                38 if _info_len > 110 else (30 if _info_len > 75 else 24)
            )

        # Summary table
        start2 = len(readme)+5
        ws.write(start2,0,"Metric", fmt_header)
        ws.write(start2,1,"Value", fmt_header)
        for i, row in summary.reset_index(drop=True).iterrows():
            ws.write(start2+1+i, 0, row["Metric"], fmt_body)
            _sv = row["Value"]
            if isinstance(_sv, (int, np.integer)) and not isinstance(_sv, bool):
                _sf = fmt_body_int
            elif isinstance(_sv, (float, np.floating)) and pd.notna(_sv) and float(_sv).is_integer():
                _sf = fmt_body_int
            elif isinstance(_sv, (float, np.floating)):
                _sf = fmt_body_num
            else:
                _sf = fmt_body
            ws.write(start2+1+i, 1, _sv, _sf)
            ws.set_row(start2+1+i, 22)

    # ------------------------------------------------------------
    # FINAL WORKSHEET ORDER — FORCE IMMEDIATELY BEFORE SAVE/CLOSE
    # ------------------------------------------------------------
    # Visible order requested by user:
    # DASHBOARD -> INTRADAY -> ETF_ANALYSIS -> TOP_5_SECTORS ->
    # THEME_RRG -> TRADE_CANDIDATES -> BUYING_ANALYSIS ->
    # NEAR_BUY_WATCHLIST -> LOSERS_DIP_ANALYSIS ->
    # ALTERNATIVE_ASSETS -> INFO
    # Hidden _CHART_DATA is moved to the very end.
    _desired_sheet_order = [
        "DASHBOARD",
        "INTRADAY",
        "ETF_ANALYSIS",
        "TOP_5_SECTORS",
        "THEME_RRG",
        "TRADE_CANDIDATES",
        "BUYING_ANALYSIS",
        "NEAR_BUY_WATCHLIST",
        "LOSERS_DIP_ANALYSIS",
        "ALTERNATIVE_ASSETS",
        "INFO",
        "_CHART_DATA",
    ]

    _sheet_map = {ws.get_name(): ws for ws in book.worksheets_objs}
    _missing_sheets = [s for s in _desired_sheet_order if s not in _sheet_map]
    if _missing_sheets:
        raise RuntimeError(
            f"Cannot force worksheet order; missing sheets: {_missing_sheets}"
        )

    book.worksheets_objs[:] = [_sheet_map[s] for s in _desired_sheet_order]

    # Keep helper data hidden and last.
    _sheet_map["_CHART_DATA"].hide()

    print("   Final sheet order:", [ws.get_name() for ws in book.worksheets_objs])

    print("   Excel saved:", excel_file)
    print("   BUILD       : V4.8.3 — FORCED SHEET ORDER")


    # -----------------------------
    # STEP 12 — Daily CSV snapshots
    # -----------------------------
    print("\n[9/10] Saving daily CSV snapshots ...")
    analysis.sort_values("RankOverall").to_csv(OUTPUT_DIR / f"ETF_RANKING_{RUN_DATE}.csv", index=False)
    theme_rrg.to_csv(OUTPUT_DIR / f"THEME_RRG_{RUN_DATE}.csv", index=False)
    trade_candidates.to_csv(OUTPUT_DIR / f"TRADE_CANDIDATES_{RUN_DATE}.csv", index=False)
    dip_candidates.to_csv(OUTPUT_DIR / f"DIP_OPPORTUNITIES_{RUN_DATE}.csv", index=False)

    # -----------------------------
    # STEP 13 — ZIP all outputs
    # -----------------------------
    zip_file = OUTPUT_DIR / f"NSE_ETF_RRG_OUTPUT_{RUN_DATE}.zip"
    with zipfile.ZipFile(zip_file, "w", zipfile.ZIP_DEFLATED) as z:
        for p in [dashboard_file, excel_file]:
            if p.exists():
                z.write(p, arcname=p.name)
        for p in CHART_DIR.glob("*.png"):
            z.write(p, arcname=f"CHARTS/{p.name}")
        for p in OUTPUT_DIR.glob(f"*_{RUN_DATE}.csv"):
            z.write(p, arcname=f"CSV/{p.name}")

    print("\n[10/10] FINISHED")
    print("="*72)
    print("Dashboard :", dashboard_file)
    print("Excel     :", excel_file)
    print("ZIP       :", zip_file)
    print("Charts    :", CHART_DIR)
    print("="*72)

    # Show the actual positive-side decision summary in Colab.
    print("\nTOP 5 QUALIFIED EQUITY THEMES / MARKET GROUPS")
    sector_console_cols = [
        c for c in [
            "SectorRank","Theme","SectorRRG","SectorStrengthScore",
            "TodaySectorPct","BreadthPct","MedianVolumeVs30D","SectorTurnoverCr"
        ] if c in top5_sectors.columns
    ]
    if len(top5_sectors):
        try:
            from IPython.display import display, HTML
            display(
                top5_sectors[sector_console_cols].style
                .background_gradient(subset=["SectorStrengthScore"], cmap="RdYlGn")
                .format({
                    "SectorStrengthScore":"{:.1f}",
                    "TodaySectorPct":"{:.2f}",
                    "BreadthPct":"{:.1f}",
                    "MedianVolumeVs30D":"{:.2f}",
                    "SectorTurnoverCr":"{:.2f}",
                })
            )
        except Exception:
            print(top5_sectors[sector_console_cols].to_string(index=False))
    else:
        print("No LEADING / IMPROVING Equity theme / market group qualified.")

    print("\nEQUITY SWING BUY PRIORITY")
    buy_console_cols = [
        c for c in [
            "BuyRank","Symbol","Theme","SectorRank","SectorRRG","NSE_ChangePct",
            "VolumePaceVs30D","TodayTurnoverCr","Tradability","RangePositionPct",
            "Quadrant","TrendPoints","MomentumPhase","EntryStretch",
            "BuyQualityScore","BuyPrioritySignal"
        ] if c in all_latest_gainers.columns
    ]
    if len(all_latest_gainers):
        try:
            from IPython.display import display
            display(
                all_latest_gainers[buy_console_cols].style
                .background_gradient(subset=["BuyQualityScore"], cmap="RdYlGn")
                .format({
                    "NSE_ChangePct":"{:.2f}",
                    "VolumePaceVs30D":"{:.2f}",
                    "TodayTurnoverCr":"{:.2f}",
                    "RangePositionPct":"{:.1f}",
                    "BuyQualityScore":"{:.1f}",
                })
            )
        except Exception:
            print(all_latest_gainers[buy_console_cols].to_string(index=False))
    else:
        print("NO QUALIFIED BUY NOW — no ETF passed every strict positive-side swing rule.")

    print("\nNEAR BUY WATCHLIST — NOT BUY SIGNALS")
    near_console_cols = [
        c for c in [
            "NearBuyRank","Symbol","Theme","SectorRRG","NSE_ChangePct",
            "VolumePaceVs30D","Tradability","RangePositionPct","Quadrant",
            "TrendPoints","MomentumPhase","EntryStretch","BuyQualityScore",
            "WhyNotQualified"
        ] if c in positive_watchlist.columns
    ]
    if len(positive_watchlist):
        try:
            from IPython.display import display, HTML
            display(
                positive_watchlist[near_console_cols].head(5).style
                .background_gradient(subset=["BuyQualityScore"], cmap="RdYlGn")
                .format({
                    "NSE_ChangePct":"{:.2f}",
                    "VolumePaceVs30D":"{:.2f}",
                    "RangePositionPct":"{:.1f}",
                    "BuyQualityScore":"{:.1f}",
                })
            )
            display(HTML(f'<p><b>Dashboard file:</b> {dashboard_file}</p>'))
        except Exception:
            print(positive_watchlist[near_console_cols].head(5).to_string(index=False))
    else:
        print("No near-buy candidate available.")

    print("\nTOP 5 INTRADAY EQUITY THEMES / MARKET GROUPS")
    print("Intraday scan window:", INTRADAY_SCAN_WINDOW)
    print(
        f"Intraday universe: Top {INTRADAY_ACTIVITY_COUNT} active Equity / International ETFs "
        f"(50% turnover + 30% current volume + 20% volume pace)"
    )
    print(
        f"Swing / negative universe remains: NSE current volume >= "
        f"{NSE_MIN_CURRENT_VOLUME:,} units"
    )
    intraday_group_console_cols = [
        c for c in [
            "IntradayGroupRank","Theme","SectorRRG","IntradayGroupStrengthScore",
            "TodaySectorPct","BreadthPct","MedianVolumeVs30D","SectorTurnoverCr"
        ] if c in top5_intraday_groups.columns
    ]
    if len(top5_intraday_groups):
        try:
            from IPython.display import display
            display(
                top5_intraday_groups[intraday_group_console_cols].style
                .background_gradient(subset=["IntradayGroupStrengthScore"], cmap="RdYlGn")
                .format({
                    "IntradayGroupStrengthScore":"{:.1f}",
                    "TodaySectorPct":"{:.2f}",
                    "BreadthPct":"{:.1f}",
                    "MedianVolumeVs30D":"{:.2f}",
                    "SectorTurnoverCr":"{:.2f}",
                })
            )
        except Exception:
            print(top5_intraday_groups[intraday_group_console_cols].to_string(index=False))
    else:
        print("No positive LEADING / IMPROVING intraday equity group qualified.")

    print("\nEQUITY INTRADAY BUY PRIORITY")
    intraday_console_cols = [
        c for c in [
            "IntradayRank","Symbol","Theme","IntradayGroupRank","SectorRRG",
            "NSE_ChangePct","IntradayOpenDrivePct","VolumePaceVs30D",
            "IntradayTradability","RangePositionPct","Quadrant","TrendPoints",
            "IntradayMoveATR","IntradayBuyScore","IntradaySignal"
        ] if c in intraday_buys.columns
    ]
    if len(intraday_buys):
        try:
            from IPython.display import display
            display(
                intraday_buys[intraday_console_cols].style
                .background_gradient(subset=["IntradayBuyScore"], cmap="RdYlGn")
                .format({
                    "NSE_ChangePct":"{:.2f}",
                    "IntradayOpenDrivePct":"{:.2f}",
                    "VolumePaceVs30D":"{:.2f}",
                    "RangePositionPct":"{:.1f}",
                    "IntradayMoveATR":"{:.2f}",
                    "IntradayBuyScore":"{:.1f}",
                })
            )
        except Exception:
            print(intraday_buys[intraday_console_cols].to_string(index=False))
    else:
        print("NO QUALIFIED INTRADAY BUY NOW.")

    print("\nINTRADAY NEAR BUY WATCHLIST — NOT BUY SIGNALS")
    intraday_near_console_cols = [
        c for c in [
            "IntradayWatchRank","Symbol","Theme","NSE_ChangePct","IntradayOpenDrivePct",
            "VolumePaceVs30D","IntradayTradability","RangePositionPct","Quadrant",
            "IntradayMoveATR","IntradayBuyScore","WhyNotIntradayQualified"
        ] if c in intraday_watchlist.columns
    ]
    if len(intraday_watchlist):
        try:
            from IPython.display import display
            display(
                intraday_watchlist[intraday_near_console_cols].head(5).style
                .background_gradient(subset=["IntradayBuyScore"], cmap="RdYlGn")
                .format({
                    "NSE_ChangePct":"{:.2f}",
                    "IntradayOpenDrivePct":"{:.2f}",
                    "VolumePaceVs30D":"{:.2f}",
                    "RangePositionPct":"{:.1f}",
                    "IntradayMoveATR":"{:.2f}",
                    "IntradayBuyScore":"{:.1f}",
                })
            )
        except Exception:
            print(intraday_watchlist[intraday_near_console_cols].head(5).to_string(index=False))
    else:
        print("No intraday near-buy candidate available.")

    print("\nNEGATIVE ETF REBOUND OPPORTUNITIES — DYNAMIC LIQUIDITY + REBOUND QUALITY")
    print("Negative universe mode:", NEGATIVE_UNIVERSE_MODE)
    print("Eligible negative ETFs:", NEGATIVE_UNIVERSE_COUNT)
    dip_display_cols = [
        "Symbol","Theme","DayReturnPct","DipBuyScore","DipSignal",
        "Theme_Quadrant","Quadrant","LTP","Avg30TurnoverCr",
        "VolumeMultiple","Week1ReturnPct","Month1ReturnPct"
    ]
    if len(dip_candidates):
        try:
            from IPython.display import display
            display(dip_candidates[dip_display_cols].head(20).style
                    .background_gradient(subset=["DipBuyScore"], cmap="RdYlGn")
                    .format({
                        "DayReturnPct":"{:.2f}",
                        "DipBuyScore":"{:.1f}",
                        "LTP":"{:.2f}",
                        "Avg30TurnoverCr":"{:.2f}",
                        "VolumeMultiple":"{:.2f}",
                        "Week1ReturnPct":"{:.2f}",
                        "Month1ReturnPct":"{:.2f}",
                    }))
        except Exception:
            print(dip_candidates[dip_display_cols].head(20).to_string(index=False))
    else:
        print("No liquid equity ETF is down 1% or more in the latest NSE session.")

    # -----------------------------
    # MOBILE RETURN PAYLOAD
    # -----------------------------
    return {
        "run_date": RUN_DATE,
        "analysis_time": ANALYSIS_TIME_IST,
        "market_mode": CURRENT_SESSION_MODE,
        "intraday_window": INTRADAY_SCAN_WINDOW,
        "nse_snapshot": NSE_DOWNLOAD_TIMESTAMP,
        "top_groups": top5_sectors.copy(),
        "swing_buys": all_latest_gainers.copy(),
        "near_buys": positive_watchlist.copy(),
        "rebound": all_latest_losers.copy(),
        "deep_fall": deep_fall_watch.copy(),
        "unusual_volume": top_volume_surges.copy(),
        "intraday_groups": top5_intraday_groups.copy(),
        "intraday_buys": intraday_buys.copy(),
        "intraday_near": intraday_watchlist.copy(),
        "excel_file": str(excel_file),
    }


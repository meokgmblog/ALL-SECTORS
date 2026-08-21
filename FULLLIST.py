import os
import time
import zoneinfo
from datetime import datetime
import pandas as pd
import requests
import streamlit as st

# ==========================================
# PAGE CONFIGURATION
# ==========================================
st.set_page_config(
    page_title="Live Sector Scope Tracker & Alpha Radar",
    layout="wide",
    initial_sidebar_state="collapsed"
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FNO_EXCEL_PATH = os.path.join(BASE_DIR, "FNO all list.xlsx")
INSTRUMENTS_CSV_PATH = os.path.join(BASE_DIR, "instruments.csv")

ACCESS_TOKEN = st.secrets.get("ACCESS_TOKEN", "eyJ0eXAiOiJKV1QiLCJrZXlfaWQiOiJza192MS4wIiwiYWxnIjoiSFMyNTYifQ.eyJzdWIiOiI2M0FZSEUiLCJqdGkiOiI2YTMwY2UxNTY4ODI0Zjc3ZDc1NmU3NjgiLCJpc011bHRpQ2xpZW50IjpmYWxzZSwiaXNQbHVzUGxhbiI6ZmFsc2UsImlzRXh0ZW5kZWQiOnRydWUsImlhdCI6MTc4MTU4MzM4MSwiaXNzIjoidWRhcGktZ2F0ZXdheS1zZXJ2aWNlIiwiZXhwIjoxODEzMTgzMjAwfQ.IoRDQhbhcn3w9Fkw75N3eBSamLcaA8GcAhVjf5K-iL8")
REFRESH_INTERVAL_SECONDS = 60
IST = zoneinfo.ZoneInfo("Asia/Kolkata")

INDEX_KEY_MAP = {
    "NIFTY50": "NSE_INDEX|Nifty 50",
    "SENSEX": "BSE_INDEX|SENSEX",
    "IT": "NSE_INDEX|Nifty IT",
    "FMCG": "NSE_INDEX|Nifty FMCG",
    "PHARMA": "NSE_INDEX|Nifty Pharma",
    "PVT BANKS": "NSE_INDEX|Nifty Private Bank",
    "NIFTY MID SELECT": "NSE_INDEX|Nifty Midcap Select",
    "AUTO": "NSE_INDEX|Nifty Auto",
    "FIN SERVICE": "NSE_INDEX|Nifty Financial Services",
    "BANKS": "NSE_INDEX|Nifty Bank",
    "CEMENT": "NSE_INDEX|Nifty Commodities",
    "ENERGY": "NSE_INDEX|Nifty Energy",
    "METAL": "NSE_INDEX|Nifty Metal",
    "PSU BANK": "NSE_INDEX|Nifty PSU Bank",
    "REALITY": "NSE_INDEX|Nifty Realty"
}

NORMALIZED_INDEX_MAP = {
    str(k).strip().upper().replace("BANKS", "BANK").replace("REALITY", "REALTY"): v
    for k, v in INDEX_KEY_MAP.items()
}

# File Existence Guard
if not os.path.exists(FNO_EXCEL_PATH) or not os.path.exists(INSTRUMENTS_CSV_PATH):
    st.error("⚠️ **Required Data Files Missing!**")
    st.info(f"Please make sure `FNO all list.xlsx` and `instruments.csv` exist in: `{BASE_DIR}`")
    st.stop()

# ==========================================
# DATA PREPARATION & MAPPING
# ==========================================
@st.cache_data(ttl=86400)
def load_data_and_mappings():
    df = pd.read_excel(FNO_EXCEL_PATH, engine="openpyxl")
    df = df[["SYMBOL", "SECTOR"]].dropna(subset=["SECTOR"])
    df = df[df["SECTOR"].astype(str).str.strip() != ""]
    df = df.drop_duplicates(subset=["SYMBOL", "SECTOR"])

    inst_df = pd.read_csv(INSTRUMENTS_CSV_PATH)
    instrument_map = {}

    for _, row in inst_df.iterrows():
        try:
            seg = str(row["segment"]).strip()
            if seg in ["NSE_EQ", "NSE_INDEX", "BSE_INDEX"]:
                instrument_map[str(row["trading_symbol"]).strip()] = str(row["instrument_key"]).strip()
        except Exception:
            pass

    sector_map = {}
    all_keys_to_fetch = []

    for _, row in df.iterrows():
        sym = str(row["SYMBOL"]).strip()
        sec = str(row["SECTOR"]).strip()

        if sym not in instrument_map:
            continue

        sector_map.setdefault(sec, []).append(sym)
        key = instrument_map[sym]
        if key not in all_keys_to_fetch:
            all_keys_to_fetch.append(key)

    for idx_key in INDEX_KEY_MAP.values():
        if idx_key not in all_keys_to_fetch:
            all_keys_to_fetch.append(idx_key)

    return df, sector_map, instrument_map, all_keys_to_fetch

# ==========================================
# BATCH SNAPSHOT FETCH ENGINE
# ==========================================
def fetch_market_snapshots_batched(keys_list, access_token, batch_size=50):
    snapshot_data = {}
    headers = {
        "Accept": "application/json",
        "Authorization": f"Bearer {access_token}"
    }

    for i in range(0, len(keys_list), batch_size):
        batch = keys_list[i:i + batch_size]
        instrument_csv = ",".join(batch)
        url = f"https://api.upstox.com/v2/market-quote/quotes?instrument_key={instrument_csv}"

        try:
            res = requests.get(url, headers=headers, timeout=10)
            if res.status_code == 200:
                data = res.json().get("data", {})
                for returned_key, metrics in data.items():
                    actual_key = metrics.get("instrument_token", returned_key)
                    last_price = metrics.get("last_price")
                    net_change = metrics.get("net_change")

                    if last_price is not None and net_change is not None:
                        prev_close = last_price - net_change
                        if prev_close > 0:
                            change_pct = (net_change / prev_close) * 100
                            snapshot_data[actual_key] = change_pct
            elif res.status_code == 429:
                time.sleep(2)
        except Exception:
            pass

        time.sleep(0.05)

    return snapshot_data

# ==========================================
# METRICS & RADAR COMPUTATION
# ==========================================
def calculate_scope_metrics(sector_map, instrument_map, live_performance_matrix):
    nifty50_key = INDEX_KEY_MAP.get("NIFTY50")
    market_baseline = live_performance_matrix.get(nifty50_key, 0.0)
    divisor_safeguard = abs(market_baseline) if abs(market_baseline) >= 0.05 else 0.10

    raw_dashboard_data = []
    all_market_stocks = []

    for sector, stocks in sector_map.items():
        bullish, bearish, total = 0, 0, 0
        stock_movements = []

        norm_sector = str(sector).strip().upper().replace("BANKS", "BANK").replace("REALITY", "REALTY")
        idx_instrument = NORMALIZED_INDEX_MAP.get(norm_sector)
        true_index_change = live_performance_matrix.get(idx_instrument)

        for sym in stocks:
            key = instrument_map.get(sym)
            if key in live_performance_matrix:
                change = live_performance_matrix[key]
                total += 1

                if change > 0.001:
                    bullish += 1
                elif change < -0.001:
                    bearish += 1

                stock_movements.append((sym, change))

        if total == 0:
            continue

        if true_index_change is None and stock_movements:
            true_index_change = sum([s[1] for s in stock_movements]) / len(stock_movements)
        elif true_index_change is None:
            true_index_change = 0.00

        true_index_change = round(true_index_change, 2)
        bullish_pct = round((bullish / total) * 100, 2) if total > 0 else 0.0
        bearish_pct = round((bearish / total) * 100, 2) if total > 0 else 0.0
        multiplier_ratio = round(true_index_change / divisor_safeguard, 2)
        score = round(bullish_pct - bearish_pct, 2)

        raw_dashboard_data.append({
            "Sector Group": sector,
            "Bullish": bullish,
            "Bearish": bearish,
            "Total": total,
            "Bullish %": f"{bullish_pct:.2f}%",
            "Bearish %": f"{bearish_pct:.2f}%",
            "True Index %": true_index_change,
            "True Index % Str": f"{true_index_change:+.2f}%",
            "Scope Multiplier": f"{multiplier_ratio:+.2f}x",
            "Score": score
        })

        if sector not in ["SENSEX", "NIFTY50", "NIFTY MID SELECT"]:
            for sym, stock_chg in stock_movements:
                relative_alpha = stock_chg - true_index_change
                all_market_stocks.append({
                    "ticker": sym,
                    "sector": sector,
                    "change": stock_chg,
                    "index_change": true_index_change,
                    "alpha": relative_alpha
                })

    scope_df = pd.DataFrame(raw_dashboard_data).sort_values(by="True Index %", ascending=False)

    unique_market_stocks = {}
    for s in all_market_stocks:
        if s["ticker"] not in unique_market_stocks:
            unique_market_stocks[s["ticker"]] = s

    filtered_stocks = list(unique_market_stocks.values())
    return scope_df, filtered_stocks, market_baseline

# ==========================================
# TIME CONTROL & MARKET STATE
# ==========================================
def is_market_open():
    now = datetime.now(IST)
    if now.weekday() >= 5:
        return False
    start_time = now.replace(hour=9, minute=14, second=0, microsecond=0)
    end_time = now.replace(hour=15, minute=13, second=0, microsecond=0)
    return start_time <= now <= end_time

# ==========================================
# DASHBOARD RENDERING ENGINE
# ==========================================
st.title("📊 Live Sector Scope Tracker & Relative Multiplier Radar")

with st.spinner("Loading Mapping Files & Initializing Instrument Universe..."):
    _, sector_map, instrument_map, all_keys_to_fetch = load_data_and_mappings()

@st.fragment(run_every=REFRESH_INTERVAL_SECONDS if is_market_open() else None)
def render_scope_radar():
    now = datetime.now(IST)
    now_str = now.strftime("%H:%M:%S IST")
    today_str = now.strftime("%Y-%m-%d")
    market_active = is_market_open()

    # Reset cache on new session date
    if 'scope_frozen_date' in st.session_state and st.session_state['scope_frozen_date'] != today_str:
        st.session_state.pop('scope_live_matrix', None)
        st.session_state.pop('scope_frozen_time', None)
        st.session_state.pop('scope_frozen_date', None)

    if market_active:
        st.success(f"🟢 **MARKET LIVE** — Last Refresh: {now_str}")
        live_matrix = fetch_market_snapshots_batched(all_keys_to_fetch, ACCESS_TOKEN)
        if live_matrix:
            st.session_state['scope_live_matrix'] = live_matrix
            st.session_state['scope_frozen_time'] = now_str
            st.session_state['scope_frozen_date'] = today_str
    else:
        if 'scope_live_matrix' in st.session_state and st.session_state['scope_live_matrix']:
            live_matrix = st.session_state['scope_live_matrix']
            frozen_at = st.session_state.get('scope_frozen_time', '3:13:00 IST')
            st.warning(f"🔴 **MARKET CLOSED — DATA FROZEN AT 3:13 PM IST** (Snapshot time: {frozen_at})")
        else:
            st.warning(f"🔴 **MARKET CLOSED** — Fetching 3:13 PM closing market snapshot. Time: {now_str}")
            live_matrix = fetch_market_snapshots_batched(all_keys_to_fetch, ACCESS_TOKEN)
            if live_matrix:
                st.session_state['scope_live_matrix'] = live_matrix
                st.session_state['scope_frozen_time'] = now_str
                st.session_state['scope_frozen_date'] = today_str

    if not live_matrix:
        st.error("⚠️ Unable to load market quotes from Upstox API.")
        return

    scope_df, filtered_stocks, market_baseline = calculate_scope_metrics(sector_map, instrument_map, live_matrix)

    # --- Top Banner Metrics ---
    m1, m2, m3 = st.columns(3)
    m1.metric("Nifty 50 Baseline", f"{market_baseline:+.2f}%")
    m2.metric("Tracked Sectors", len(scope_df))
    m3.metric("Total F&O Universe", len(filtered_stocks))

    # --- Section 1: Scope Analysis Table ---
    st.subheader("Sector Scope Analysis Matrix")
    display_scope = scope_df.drop(columns=["True Index %"]).rename(columns={"True Index % Str": "True Index %"})
    st.dataframe(display_scope, use_container_width=True, hide_index=True)

    def prepare_radar_df(stock_list, label_text):
        records = []
        for s in stock_list:
            symbol = s["ticker"]
            records.append({
                "CHART_URL": f"https://www.tradingview.com/chart/?symbol=NSE:{symbol}&interval=5",
                "Sector Peer Group": s["sector"],
                "Stock Ret %": f"{s['change']:+.2f}%",
                "Sector Ret %": f"{s['index_change']:+.2f}%",
                "Net Alpha Edge": f"{s['alpha']:+.2f}%",
                "Signal Framework": label_text
            })
        return pd.DataFrame(records)

    table_config = {
        "CHART_URL": st.column_config.LinkColumn(
            "Ticker",
            help="Click to open TradingView chart",
            display_text=r"symbol=NSE:([^&]+)"
        )
    }

    # --- Section 2: Strategic Radars ---
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("🔥 Top True Alpha Momentum Leaders (Buy Focus)")
        filtered_stocks.sort(key=lambda x: x["alpha"], reverse=True)
        top_momentum = prepare_radar_df(filtered_stocks[:8], "INSTITUTIONAL ACCUMULATION")
        st.dataframe(top_momentum, column_config=table_config, use_container_width=True, hide_index=True)

    with col2:
        st.subheader("🛡️ Defensive Radar: Top Iron Domes")
        defensive = [s for s in filtered_stocks if s["index_change"] < -0.50]
        defensive.sort(key=lambda x: x["alpha"], reverse=True)
        top_defensive = prepare_radar_df(defensive[:8], "SHIELDED FROM MARKET SELLOFF")
        if not top_defensive.empty:
            st.dataframe(top_defensive, column_config=table_config, use_container_width=True, hide_index=True)
        else:
            st.info("No sectors are currently under systemic selling pressure (< -0.50%) to trigger defensive shields.")

    st.subheader("🚀 Trend Rider Radar: Top Sector Velocity Leaders")
    trend_riders = [s for s in filtered_stocks if s["index_change"] > 0.30]
    trend_riders.sort(key=lambda x: x["change"], reverse=True)
    top_riders = prepare_radar_df(trend_riders[:8], "STRONG MOMENTUM WITH SECTOR TAILWIND")
    if not top_riders.empty:
        st.dataframe(top_riders, column_config=table_config, use_container_width=True, hide_index=True)
    else:
        st.info("No sectors are currently showing strong upward velocity (> +0.30%) to trigger Trend Riders.")

    # --- Section 3: Sector-Wise Stocks Breakdown ---
    st.subheader("📊 Sector-Wise Breakdown Radar (All Stocks)")

    for sector_name, stock_symbols in sector_map.items():
        norm_sec = str(sector_name).strip().upper().replace("BANKS", "BANK").replace("REALITY", "REALTY")
        idx_inst = NORMALIZED_INDEX_MAP.get(norm_sec)
        sec_idx_change = live_matrix.get(idx_inst)

        sec_stocks_list = []
        for sym in stock_symbols:
            k = instrument_map.get(sym)
            chg = live_matrix.get(k)
            if chg is not None:
                sec_stocks_list.append((sym, chg))

        if sec_idx_change is None and sec_stocks_list:
            sec_idx_change = sum([s[1] for s in sec_stocks_list]) / len(sec_stocks_list)
        elif sec_idx_change is None:
            sec_idx_change = 0.00

        sec_idx_change = round(sec_idx_change, 2)
        sec_stocks_list.sort(key=lambda x: x[1], reverse=True)

        sec_table_data = []
        for sym, stock_chg in sec_stocks_list:
            alpha = round(stock_chg - sec_idx_change, 2)
            framework_label = "OUTPERFORMING" if alpha > 0 else "UNDERPERFORMING"
            sec_table_data.append({
                "CHART_URL": f"https://www.tradingview.com/chart/?symbol=NSE:{sym}&interval=5",
                "Sector Peer Group": sector_name,
                "Stock Ret %": f"{stock_chg:+.2f}%",
                "Sector Ret %": f"{sec_idx_change:+.2f}%",
                "Net Alpha Edge": f"{alpha:+.2f}%",
                "Signal Framework": framework_label
            })

        if sec_table_data:
            with st.expander(f"📁 **{sector_name.upper()}** ({len(sec_table_data)} Stocks)", expanded=False):
                st.dataframe(
                    pd.DataFrame(sec_table_data),
                    column_config=table_config,
                    use_container_width=True,
                    hide_index=True
                )

render_scope_radar()
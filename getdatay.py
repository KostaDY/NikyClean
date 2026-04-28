#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import time
import json
import logging
import datetime as dt
from pathlib import Path

import pandas as pd
import yfinance as yf
import requests

# ================= CONFIG =================

WORKBOOK = Path("/Users/kostayanev/NikyClean/YahooDataOutput.xlsx")
INPUT_SHEET = "Data"
OUTPUT_SHEET = "YahooData"

REQUEST_DELAY = 0.12
RETRY_ATTEMPTS = 2

# --- TTL CACHE (SAFE) ---
CACHE_FILE = Path("ticker_cache.json")
CACHE_TTL_HOURS = 12

# ==========================================

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger()

# ============ SESSION REUSE ===============

session = requests.Session()
yf.shared._requests = session

# ============ CACHE =======================

def load_cache():
    if CACHE_FILE.exists():
        try:
            return json.loads(CACHE_FILE.read_text())
        except:
            return {}
    return {}

def save_cache(cache):
    CACHE_FILE.write_text(json.dumps(cache))

def is_cache_valid(entry):
    if not isinstance(entry, dict):
        return False

    ts = entry.get("_ts")
    if not ts:
        return False

    age = (time.time() - ts) / 3600
    return age < CACHE_TTL_HOURS

CACHE = load_cache()

# ============ HELPERS =====================

def format_ts(ts):
    if ts:
        try:
            return dt.datetime.fromtimestamp(ts, dt.UTC).strftime("%Y-%m-%d")
        except:
            return None
    return None

def load_tickers():
    df = pd.read_excel(WORKBOOK, sheet_name=INPUT_SHEET)
    col = df.columns[0]

    return (
        df[col]
        .dropna()
        .astype(str)
        .str.strip()
        .tolist()
    )

# ============ FETCH (OPTIMIZED) ===========

def fetch_one(ticker):

    # --- CACHE ---
    if ticker in CACHE and is_cache_valid(CACHE[ticker]):
        return CACHE[ticker]["data"]

    for attempt in range(RETRY_ATTEMPTS):
        try:
            t = yf.Ticker(ticker)

            info = t.info
            if not info or info.get("regularMarketPrice") is None:
                raise ValueError("No data")

            # --- LOCAL EXTRACTION ---
            ex_div = info.get("exDividendDate")
            target_mean = info.get("targetMeanPrice")
            target_low = info.get("targetLowPrice")
            target_high = info.get("targetHighPrice")
            analysts = info.get("numberOfAnalystOpinions")
            div_yield = info.get("dividendYield")
            div_rate = info.get("dividendRate")

            # --- EARNINGS ---
            earnings = None

            try:
                cal = t.calendar
                if cal is not None and not cal.empty:
                    val = cal.iloc[0][0]
                    if pd.notna(val):
                        earnings = str(val)
            except:
                pass

            if earnings is None:
                try:
                    ed = t.earnings_dates
                    if ed is not None and not ed.empty:
                        earnings = str(ed.index[0].date())
                except:
                    pass

            result = {
                "Ticker": ticker,
                "Ex-Dividend Date": format_ts(ex_div),
                "Earnings Date": earnings,
                "1-Year Target (Mean)": target_mean,
                "Target Low": target_low,
                "Target High": target_high,
                "# of Analysts": analysts,
                "Div Yield (%)": (div_yield * 100 if div_yield is not None else None),
                "Div ($)": div_rate,
            }

            # --- STORE CACHE ---
            CACHE[ticker] = {
                "_ts": time.time(),
                "data": result
            }

            return result

        except Exception:
            if attempt == RETRY_ATTEMPTS - 1:
                return {
                    "Ticker": ticker,
                    "Ex-Dividend Date": None,
                    "Earnings Date": None,
                    "1-Year Target (Mean)": None,
                    "Target Low": None,
                    "Target High": None,
                    "# of Analysts": None,
                    "Div Yield (%)": None,
                    "Div ($)": None,
                }

            time.sleep(0.5)

# ============ MAIN ========================

def main():

    tickers = load_tickers()
    logger.info(f"{len(tickers)} tickers loaded")

    rows = []

    start = time.time()

    for i, ticker in enumerate(tickers, 1):
        logger.info(f"[{i}/{len(tickers)}] {ticker}")

        result = fetch_one(ticker)
        rows.append(result)

        time.sleep(REQUEST_DELAY)

    elapsed = time.time() - start
    logger.info(f"Completed in {elapsed:.1f}s")

    save_cache(CACHE)

    df_new = pd.DataFrame(rows)

    # ===== SAFE WRITE (PRESERVE ALL SHEETS) =====
    with pd.ExcelFile(WORKBOOK) as xls:
        sheets = {name: xls.parse(name) for name in xls.sheet_names}

    sheets[OUTPUT_SHEET] = df_new

    with pd.ExcelWriter(WORKBOOK, engine="openpyxl", mode="w") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)

    logger.info("Excel updated safely")

    os.system(f"open '{WORKBOOK}'")

# ============ RUN ==========================
if __name__ == "__main__":
    main()
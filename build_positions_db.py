#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import time
from pathlib import Path

import pandas as pd
import yfinance as yf


# ============================================================
# FILES
# ============================================================

INPUT_JSON = Path("Base_Portfolio.json")
OUTPUT_JSON = Path("Positions_DB.json")

DAYS_TO_KEEP = 200
DOWNLOAD_PERIOD = "18mo"      # enough to calculate 52-week low/high
AV_VOLUME_WINDOW = 20         # rolling average volume
SLEEP_SECONDS = 0.30


# ============================================================
# HELPERS
# ============================================================

def empty_if_na(x):
    if pd.isna(x):
        return ""
    if isinstance(x, pd.Timestamp):
        return x.strftime("%Y-%m-%d")
    if isinstance(x, float):
        return round(x, 6)
    return x


def load_tickers_from_base_portfolio(path: Path):
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    # Supports both:
    # 1) [ {"Ticker":"AAPL"}, ... ]
    # 2) { "Port": [ {"Ticker":"AAPL"}, ... ] }
    if isinstance(data, dict) and "Port" in data:
        rows = data["Port"]
    elif isinstance(data, list):
        rows = data
    else:
        raise ValueError("Unsupported JSON structure. Expected list or {'Port': list}.")

    tickers = []
    for row in rows:
        t = str(row.get("Ticker", "")).strip()
        if t and t.lower() != "nan":
            tickers.append(t)

    # preserve order, remove duplicates
    return list(dict.fromkeys(tickers))


def fetch_one_ticker(ticker: str):
    print(f"Fetching {ticker}...")

    result = []

    try:
        tk = yf.Ticker(ticker)

        info = {}
        try:
            info = tk.info or {}
        except Exception:
            info = {}

        beta = info.get("beta", "")
        pe = info.get("trailingPE", "")

        hist = tk.history(
            period=DOWNLOAD_PERIOD,
            interval="1d",
            auto_adjust=False
        )

        if hist.empty:
            return result

        hist = hist.reset_index()

        # normalize Date column
        if "Date" not in hist.columns:
            return result

        hist["Date"] = pd.to_datetime(hist["Date"]).dt.date

        hist["PREVIOUS_CLOSE"] = hist["Close"].shift(1)
        hist["Change"] = hist["Close"] - hist["PREVIOUS_CLOSE"]
        hist["Change (%)"] = (hist["Change"] / hist["PREVIOUS_CLOSE"]) * 100

        hist["AV_VOLUME"] = hist["Volume"].rolling(
            AV_VOLUME_WINDOW,
            min_periods=1
        ).mean()

        # rolling 52-week values, approx. 252 trading days
        hist["52 week low"] = hist["Low"].rolling(
            252,
            min_periods=1
        ).min()

        hist["52 week high"] = hist["High"].rolling(
            252,
            min_periods=1
        ).max()

        hist = hist.tail(DAYS_TO_KEEP)

        for _, r in hist.iterrows():
            record = {
                "Date": empty_if_na(pd.Timestamp(r["Date"])),
                "PRICE": empty_if_na(r.get("Close")),
                "VOLUME": empty_if_na(r.get("Volume")),
                "AV_VOLUME": empty_if_na(r.get("AV_VOLUME")),
                "OPEN": empty_if_na(r.get("Open")),
                "PREVIOUS_CLOSE": empty_if_na(r.get("PREVIOUS_CLOSE")),
                "HIGH": empty_if_na(r.get("High")),
                "LOW": empty_if_na(r.get("Low")),
                "BETA": empty_if_na(beta),
                "Change (%)": empty_if_na(r.get("Change (%)")),
                "Change": empty_if_na(r.get("Change")),
                "P/E": empty_if_na(pe),
                "52 week low": empty_if_na(r.get("52 week low")),
                "52 week high": empty_if_na(r.get("52 week high")),
            }

            result.append(record)

    except Exception as e:
        print(f"WARNING: {ticker} failed: {e}")

    return result


# ============================================================
# MAIN
# ============================================================

def main():
    if not INPUT_JSON.exists():
        raise FileNotFoundError(f"Missing file: {INPUT_JSON}")

    tickers = load_tickers_from_base_portfolio(INPUT_JSON)

    database = {
        "source": str(INPUT_JSON),
        "days_kept": DAYS_TO_KEEP,
        "records": {}
    }

    for ticker in tickers:
        database["records"][ticker] = fetch_one_ticker(ticker)
        time.sleep(SLEEP_SECONDS)

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(database, f, indent=4, ensure_ascii=False)

    print(f"\nDone.")
    print(f"Created: {OUTPUT_JSON}")
    print(f"Tickers: {len(tickers)}")


if __name__ == "__main__":
    main()
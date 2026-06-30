#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Positions_DB.json compact structure
===================================

This script creates a compact JSON database with this structure:

{
    "source": "Base_Portfolio.json",
    "days_kept": 200,
    "columns": [
        "Date",
        "PRICE",
        "VOLUME",
        ...
    ],
    "records": {
        "EOAN.DE": [
            ["2025-08-20", 15.615, 5414130, ...],
            ["2025-08-21", 15.72,  4300000, ...]
        ],
        "AAPL": [
            ...
        ]
    }
}

The column names are stored only once in "columns".
Each ticker contains a list of rows.
Each row follows the same order as "columns".

How to use this structure in another .py program
------------------------------------------------

import json
import pandas as pd

with open("Positions_DB.json", "r", encoding="utf-8") as f:
    db = json.load(f)

columns = db["columns"]

# Example: get one ticker as a pandas DataFrame
ticker = "EOAN.DE"
rows = db["records"][ticker]

df = pd.DataFrame(rows, columns=columns)

print(df.head())

# Example: get latest PRICE
latest_price = df.iloc[-1]["PRICE"]
print(latest_price)

# Example: loop over all tickers
for ticker, rows in db["records"].items():
    df = pd.DataFrame(rows, columns=columns)
    print(ticker, df.iloc[-1]["Date"], df.iloc[-1]["PRICE"])

"""

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
# OUTPUT COLUMNS
# ============================================================

COLUMNS = [
    "Date",
    "PRICE",
    "VOLUME",
    "AV_VOLUME",
    "OPEN",
    "PREVIOUS_CLOSE",
    "HIGH",
    "LOW",
    "BETA",
    "Change (%)",
    "Change",
    "P/E",
    "52 week low",
    "52 week high",
]


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

            # IMPORTANT:
            # This row must follow exactly the same order as COLUMNS.
            record = [
                empty_if_na(pd.Timestamp(r["Date"])),
                empty_if_na(r.get("Close")),
                empty_if_na(r.get("Volume")),
                empty_if_na(r.get("AV_VOLUME")),
                empty_if_na(r.get("Open")),
                empty_if_na(r.get("PREVIOUS_CLOSE")),
                empty_if_na(r.get("High")),
                empty_if_na(r.get("Low")),
                empty_if_na(beta),
                empty_if_na(r.get("Change (%)")),
                empty_if_na(r.get("Change")),
                empty_if_na(pe),
                empty_if_na(r.get("52 week low")),
                empty_if_na(r.get("52 week high")),
            ]

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
        "columns": COLUMNS,
        "records": {}
    }

    for ticker in tickers:
        database["records"][ticker] = fetch_one_ticker(ticker)
        time.sleep(SLEEP_SECONDS)

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(
            database,
            f,
            ensure_ascii=False,
            separators=(",", ":")
        )

    print("\nDone.")
    print(f"Created: {OUTPUT_JSON}")
    print(f"Tickers: {len(tickers)}")


if __name__ == "__main__":
    main()
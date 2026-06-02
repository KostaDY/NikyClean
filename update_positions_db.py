#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import time
from pathlib import Path

import pandas as pd
import yfinance as yf


# ============================================================
# SETTINGS
# ============================================================

DB_FILE = Path("Positions_DB.json")

DAYS_TO_KEEP = 200
FETCH_PERIOD = "5d"           # fast: enough to get latest trading day
AV_VOLUME_WINDOW = 20
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


def load_db(path: Path):
    if not path.exists():
        raise FileNotFoundError(f"Missing database file: {path}")

    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_db(path: Path, db: dict):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(db, f, indent=4, ensure_ascii=False)


def latest_saved_date(records):
    dates = [r.get("Date", "") for r in records if r.get("Date", "")]
    if not dates:
        return ""
    return max(dates)


def fetch_newest_market_row(ticker: str):
    try:
        tk = yf.Ticker(ticker)

        hist = tk.history(
            period=FETCH_PERIOD,
            interval="1d",
            auto_adjust=False
        )

        if hist.empty:
            return None, None

        hist = hist.reset_index()

        if "Date" not in hist.columns:
            return None, None

        hist["Date"] = pd.to_datetime(hist["Date"]).dt.date
        hist = hist.sort_values("Date")

        newest = hist.iloc[-1]

        date_str = pd.Timestamp(newest["Date"]).strftime("%Y-%m-%d")

        row = {
            "Date": date_str,
            "PRICE": empty_if_na(newest.get("Close")),
            "VOLUME": empty_if_na(newest.get("Volume")),
            "AV_VOLUME": "",          # recalculated later
            "OPEN": empty_if_na(newest.get("Open")),
            "PREVIOUS_CLOSE": "",     # filled from existing DB
            "HIGH": empty_if_na(newest.get("High")),
            "LOW": empty_if_na(newest.get("Low")),
            "BETA": "",              # optional, fetched separately below
            "Change (%)": "",         # calculated later
            "Change": "",             # calculated later
            "P/E": "",               # optional, fetched separately below
            "52 week low": "",        # recalculated later
            "52 week high": "",       # recalculated later
        }

        return date_str, row

    except Exception as e:
        print(f"WARNING: {ticker} market row failed: {e}")
        return None, None


def fetch_slow_info(ticker: str):
    """
    BETA and P/E do not need historical download.
    Still, tk.info may be slow, so failures return blanks.
    """
    try:
        tk = yf.Ticker(ticker)
        info = tk.info or {}

        return {
            "BETA": empty_if_na(info.get("beta", "")),
            "P/E": empty_if_na(info.get("trailingPE", "")),
        }

    except Exception:
        return {
            "BETA": "",
            "P/E": "",
        }


def recalculate_derived_fields(records):
    """
    Recalculate:
    - PREVIOUS_CLOSE
    - Change
    - Change (%)
    - AV_VOLUME
    - 52 week low
    - 52 week high

    using only stored rolling database.
    """

    if not records:
        return records

    df = pd.DataFrame(records)

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date"])
    df = df.sort_values("Date")

    for col in ["PRICE", "VOLUME", "HIGH", "LOW"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df["PREVIOUS_CLOSE"] = df["PRICE"].shift(1)
    df["Change"] = df["PRICE"] - df["PREVIOUS_CLOSE"]
    df["Change (%)"] = (df["Change"] / df["PREVIOUS_CLOSE"]) * 100

    df["AV_VOLUME"] = df["VOLUME"].rolling(
        AV_VOLUME_WINDOW,
        min_periods=1
    ).mean()

    df["52 week low"] = df["LOW"].rolling(
        252,
        min_periods=1
    ).min()

    df["52 week high"] = df["HIGH"].rolling(
        252,
        min_periods=1
    ).max()

    df = df.tail(DAYS_TO_KEEP)

    out = []

    for _, r in df.iterrows():
        out.append({
            "Date": empty_if_na(r.get("Date")),
            "PRICE": empty_if_na(r.get("PRICE")),
            "VOLUME": empty_if_na(r.get("VOLUME")),
            "AV_VOLUME": empty_if_na(r.get("AV_VOLUME")),
            "OPEN": empty_if_na(r.get("OPEN")),
            "PREVIOUS_CLOSE": empty_if_na(r.get("PREVIOUS_CLOSE")),
            "HIGH": empty_if_na(r.get("HIGH")),
            "LOW": empty_if_na(r.get("LOW")),
            "BETA": empty_if_na(r.get("BETA")),
            "Change (%)": empty_if_na(r.get("Change (%)")),
            "Change": empty_if_na(r.get("Change")),
            "P/E": empty_if_na(r.get("P/E")),
            "52 week low": empty_if_na(r.get("52 week low")),
            "52 week high": empty_if_na(r.get("52 week high")),
        })

    return out


# ============================================================
# MAIN
# ============================================================

def main():
    db = load_db(DB_FILE)

    if "records" not in db:
        raise ValueError("Invalid database structure: missing 'records' key.")

    tickers = list(db["records"].keys())

    updated = 0
    skipped = 0
    failed = 0

    for ticker in tickers:
        print(f"Checking {ticker}...")

        old_records = db["records"].get(ticker, [])
        last_date = latest_saved_date(old_records)

        newest_date, newest_row = fetch_newest_market_row(ticker)

        if newest_row is None:
            print(f"  {ticker}: no data")
            failed += 1
            continue

        if last_date and newest_date <= last_date:
            print(f"  {ticker}: already up to date ({last_date})")
            skipped += 1
            continue

        slow_info = fetch_slow_info(ticker)
        newest_row["BETA"] = slow_info["BETA"]
        newest_row["P/E"] = slow_info["P/E"]

        merged = old_records + [newest_row]

        db["records"][ticker] = recalculate_derived_fields(merged)

        print(f"  {ticker}: appended {newest_date}")

        updated += 1
        time.sleep(SLEEP_SECONDS)

    db["days_kept"] = DAYS_TO_KEEP
    db["last_update_attempt"] = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
    db["last_update_mode"] = "fast_newest_date_only"

    save_db(DB_FILE, db)

    print("\nDone.")
    print(f"Updated: {updated}")
    print(f"Skipped: {skipped}")
    print(f"Failed: {failed}")
    print(f"Saved: {DB_FILE}")


if __name__ == "__main__":
    main()
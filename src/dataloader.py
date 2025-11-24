#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
from yahooquery import Ticker
import requests
from bs4 import BeautifulSoup
import time
import random
import traceback
from concurrent.futures import ThreadPoolExecutor, as_completed

# ============================================================
# CONFIGURATION
# ============================================================

DATASOURCE_XLSX = "/Users/kostayanev/NikyClean/excel/DataSource.xlsx"
DATASOURCE_SHEET = "TickerList"
DATASOURCE_COL = "A"   # column containing ticker symbols

OUTPUT_XLSX = "/Users/kostayanev/NikyClean/excel/DataSource_Raw.xlsx"
OUTPUT_SHEET = "LatestData"

HEADERS = {"User-Agent": "Mozilla/5.0"}
TIMEOUT = 8
MAX_WORKERS = 6
RETRY_DELAY = (1.0, 2.0)

# ============================================================
# READ TICKERS (preserve blanks)
# ============================================================

def read_tickers():
    df = pd.read_excel(
        DATASOURCE_XLSX,
        sheet_name=DATASOURCE_SHEET,
        usecols=DATASOURCE_COL,
        header=0
    )

    col = df.columns[0]
    tickers = []

    for v in df[col]:
        if pd.isna(v) or str(v).strip() == "":
            tickers.append("")  # preserve blank row
        else:
            tickers.append(str(v).strip())

    print(f"✔ Loaded {len(tickers)} tickers from {DATASOURCE_SHEET}!{DATASOURCE_COL}")
    return tickers


# ============================================================
# HTML CALENDAR PAGE SCRAPER
# ============================================================

def get_calendar_page(ticker):
    try:
        url = f"https://finance.yahoo.com/quote/{ticker}/calendar"
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
        if r.status_code != 200:
            return None
        return BeautifulSoup(r.text, "lxml")
    except Exception:
        return None


def extract_earnings_date(soup):
    if not soup:
        return ""
    try:
        span = soup.find("span", string="Earnings Date")
        if span:
            nxt = span.find_next("span")
            return nxt.text.strip() if nxt else ""
    except Exception:
        pass
    return ""


# ============================================================
# EXTENDED (MEDY ENGINE)
# ============================================================

def fetch_extended_once(ticker):
    try:
        yq = Ticker(ticker)

        sd = yq.summary_detail
        sd = sd.get(ticker, {}) if isinstance(sd, dict) else {}

        fin = yq.financial_data
        fin = fin.get(ticker, {}) if isinstance(fin, dict) else {}

        price_info = yq.price
        price_info = price_info.get(ticker, {}) if isinstance(price_info, dict) else {}

        dividend_yield = sd.get("dividendYield")
        ex_div = sd.get("exDividendDate")
        target = fin.get("targetMeanPrice")
        currency = price_info.get("currency", "")

        try:
            if isinstance(ex_div, (float, int)):
                ex_div = pd.to_datetime(ex_div, unit="s").strftime("%Y-%m-%d")
            else:
                ex_div = ""
        except Exception:
            ex_div = ""

        soup = get_calendar_page(ticker)
        earnings_date = extract_earnings_date(soup)

        return dict(
            Ticker=ticker,
            OneYearTarget=target,
            ExDividendDate=ex_div,
            EarningsDate=earnings_date,
            DividendYield=dividend_yield,
            Currency=currency
        )

    except Exception:
        return dict(
            Ticker=ticker, OneYearTarget=None,
            ExDividendDate=None, EarningsDate=None,
            DividendYield=None, Currency=None
        )


def fetch_extended(ticker):
    if ticker == "":
        return dict(
            Ticker="", OneYearTarget=None, ExDividendDate=None,
            EarningsDate=None, DividendYield=None, Currency=None
        )

    r = fetch_extended_once(ticker)
    if r["EarningsDate"] or r["ExDividendDate"]:
        return r

    time.sleep(random.uniform(*RETRY_DELAY))
    return fetch_extended_once(ticker)


# ============================================================
# FAST PRICE DATA — FIXED PE + FIXED AVG VOLUME
# ============================================================

def fetch_fast(tickers):

    cleaned = [t for t in tickers if t.strip() != ""]

    if not cleaned:
        return pd.DataFrame(columns=[
            "Ticker", "RefreshTime", "Close", "Open", "Last",
            "Low", "High", "PE", "Change", "ChangePct",
            "Volume", "VolumeAverage", "Beta"
        ])

    tq = Ticker(cleaned)

    prices = tq.price if isinstance(tq.price, dict) else {}
    summary = tq.summary_detail if isinstance(tq.summary_detail, dict) else {}

    rows = []
    for t in tickers:

        if t.strip() == "":
            rows.append({
                "Ticker": "", "RefreshTime": None,
                "Close": None, "Open": None, "Last": None,
                "Low": None, "High": None, "PE": None,
                "Change": None, "ChangePct": None,
                "Volume": None, "VolumeAverage": None,
                "Beta": None
            })
            continue

        p = prices.get(t, {}) if isinstance(prices.get(t), dict) else {}
        s = summary.get(t, {}) if isinstance(summary.get(t), dict) else {}

        # FIX 1: trailing PE sometimes in summary_detail
        pe = p.get("trailingPE") or s.get("trailingPE")

        # FIX 2: average volume sometimes in summary_detail
        avg_vol = (
            p.get("averageDailyVolume10Day")
            or p.get("averageDailyVolume3Month")
            or s.get("averageDailyVolume10Day")
            or s.get("averageDailyVolume3Month")
        )

        rows.append(dict(
            Ticker=t,
            RefreshTime=p.get("regularMarketTime"),
            Close=p.get("regularMarketPrice"),
            Open=p.get("regularMarketOpen"),
            Last=p.get("regularMarketPreviousClose"),
            Low=p.get("regularMarketDayLow"),
            High=p.get("regularMarketDayHigh"),
            PE=pe,
            Change=p.get("regularMarketChange"),
            ChangePct=p.get("regularMarketChangePercent"),
            Volume=p.get("regularMarketVolume"),
            VolumeAverage=avg_vol,
            Beta=s.get("beta")
        ))

    return pd.DataFrame(rows)


# ============================================================
# MAIN
# ============================================================

def main():
    print("🕒 START — FAST+MEDY ENGINE")

    tickers = read_tickers()

    print("⏱ Fetching FAST YahooQuery data…")
    df_fast = fetch_fast(tickers)

    print("⏱ Fetching EXTENDED MEDY engine data…")
    ext_rows = []
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
        futures = {ex.submit(fetch_extended, t): i for i, t in enumerate(tickers)}
        for f in as_completed(futures):
            ext_rows.append(f.result())

    df_ext = pd.DataFrame(ext_rows)

    print("⏱ Merging…")
    df = df_fast.merge(df_ext, on="Ticker", how="left")

    print("⏱ Writing Excel output…")
    with pd.ExcelWriter(OUTPUT_XLSX, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=OUTPUT_SHEET)

    print("✅ DONE — Saved:", OUTPUT_XLSX)


# ============================================================

if __name__ == "__main__":
    try:
        main()
    except Exception:
        print("\n❌ FATAL ERROR\n")
        traceback.print_exc()
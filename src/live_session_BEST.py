from __future__ import annotations

import argparse
import json
import platform
import subprocess
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, date
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd
import requests
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from yahooquery import Ticker

# --------------------------------------------------------------------
# Paths & constants
# --------------------------------------------------------------------

ROOT = Path(__file__).resolve().parent.parent
EXCEL_DIR = ROOT / "excel"
DATASOURCE_XLSX = EXCEL_DIR / "DataSource.xlsx"
RAW_XLSX = EXCEL_DIR / "DataSource_Raw.xlsx"
CACHE_DIR = ROOT / "data"
CACHE_DIR.mkdir(exist_ok=True)
SLOW_CACHE_JSON = CACHE_DIR / "slow_cache.json"

RTDATA_SHEET = "DataRT"
RTDATA_TICKER_COL = "Ticker_Symbol"

RAW_SHEET = "RTData_Raw"
RAW_TABLE_NAME = "RTData_Raw"

COLUMNS = [
    "Ticker",
    "Close",
    "Open",
    "Last",
    "Low",
    "High",
    "P/E",
    "Change",
    "ChangePct",
    "Volume",
    "VolumeAverage",
    "Beta",
    "1YT",
    "Ddate",
    "EarningDate",
    "RepDiv",
]

FAST_COLS = [
    "Ticker",
    "Close",
    "Open",
    "Last",
    "Low",
    "High",
    "P/E",
    "Change",
    "ChangePct",
    "Volume",
    "VolumeAverage",
]

SLOW_COLS = [
    "Ticker",
    "Beta",
    "1YT",
    "Ddate",
    "EarningDate",
    "RepDiv",
]


# --------------------------------------------------------------------
# Utilities
# --------------------------------------------------------------------


def now_str() -> str:
    return datetime.now().strftime("%H:%M:%S")


def load_slow_cache() -> Dict[str, Dict[str, Any]]:
    if not SLOW_CACHE_JSON.exists():
        return {}
    try:
        with SLOW_CACHE_JSON.open("r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_slow_cache(cache: Dict[str, Dict[str, Any]]) -> None:
    try:
        with SLOW_CACHE_JSON.open("w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# --------------------------------------------------------------------
# Step 1: Read tickers from DataSource.xlsx / RTData table
# --------------------------------------------------------------------


def read_tickers() -> List[Optional[str]]:
    """
    Reads ONLY the Ticker_Symbol column from Excel TABLE RTData.
    Keeps empty tickers (Option A). Never reads outside the table.
    """

    from openpyxl import load_workbook
    from openpyxl.utils.cell import range_boundaries

    wb = load_workbook(DATASOURCE_XLSX, data_only=True)
    ws = wb[RTDATA_SHEET]

    # --- locate the RTData Excel table ---
    table = None
    for t in ws._tables.values():
        if t.name == "RTData":
            table = t
            break

    if table is None:
        raise ValueError("❌ Excel table RTData not found in sheet DataRT")

    # --- parse its boundaries ---
    min_col, min_row, max_col, max_row = range_boundaries(table.ref)

    # --- read the header row ---
    headers = [
        ws.cell(min_row, col).value for col in range(min_col, max_col + 1)
    ]
    if RTDATA_TICKER_COL not in headers:
        raise ValueError(
            f"❌ Column '{RTDATA_TICKER_COL}' not found in RTData table."
        )

    ticker_col_index = headers.index(RTDATA_TICKER_COL) + min_col

    # --- read tickers ---
    tickers: List[str] = []
    for r in range(min_row + 1, max_row + 1):
        v = ws.cell(r, ticker_col_index).value
        if v is None:
            v = ""  # keep empty tickers
        tickers.append(str(v).strip())

    return tickers


# --------------------------------------------------------------------
# Step 2: FAST fetch with yahooquery
# --------------------------------------------------------------------


def fetch_fast(tickers: List[Optional[str]]) -> pd.DataFrame:
    # Fast fields via yahooquery: price & summary_detail
    syms = [t for t in tickers if t]
    if not syms:
        return pd.DataFrame({col: [] for col in FAST_COLS})

    tq = Ticker(syms, asynchronous=True)

    raw_price = tq.price or {}
    raw_sd = tq.summary_detail or {}

    def get_dict(d: Any, key: str) -> Dict[str, Any]:
        v = d.get(key)
        return v if isinstance(v, dict) else {}

    price = {s: get_dict(raw_price, s) for s in syms}
    sd = {s: get_dict(raw_sd, s) for s in syms}

    rows: Dict[str, Dict[str, Any]] = {}

    for s in syms:
        p = price.get(s, {})
        d = sd.get(s, {})

        close = d.get("regularMarketPreviousClose") or p.get("regularMarketPreviousClose")
        open_ = d.get("regularMarketOpen") or p.get("regularMarketOpen")
        last = p.get("regularMarketPrice") or d.get("regularMarketPrice")
        low = d.get("regularMarketDayLow") or p.get("regularMarketDayLow")
        high = d.get("regularMarketDayHigh") or p.get("regularMarketDayHigh")
        pe = d.get("trailingPE") or p.get("trailingPE")

        change = d.get("regularMarketChange") or p.get("regularMarketChange")
        change_pct = d.get("regularMarketChangePercent") or p.get("regularMarketChangePercent")

        if change_pct is not None:
            try:
                if abs(change_pct) > 1:
                    change_pct = change_pct / 100.0
            except Exception:
                pass

        volume = d.get("regularMarketVolume") or p.get("regularMarketVolume")
        volavg = d.get("averageVolume") or p.get("averageVolume")

        rows[s] = {
            "Ticker": s,
            "Close": close,
            "Open": open_,
            "Last": last,
            "Low": low,
            "High": high,
            "P/E": pe,
            "Change": change,
            "ChangePct": change_pct,
            "Volume": volume,
            "VolumeAverage": volavg,
        }

    empty_fast = {col: None for col in FAST_COLS}
    empty_fast["Ticker"] = None

    result_rows = []
    for t in tickers:
        if t and t in rows:
            result_rows.append(rows[t])
        else:
            result_rows.append(empty_fast.copy())

    df = pd.DataFrame(result_rows)
    for col in FAST_COLS:
        if col not in df.columns:
            df[col] = None

    return df[FAST_COLS]


# --------------------------------------------------------------------
# Step 3: SLOW fetch via Yahoo JSON API + caching + threads
# --------------------------------------------------------------------


def _get_raw(block: Dict[str, Any], key: str) -> Optional[Any]:
    v = block.get(key)
    if isinstance(v, dict):
        return v.get("raw")
    return None


def _fetch_slow_one(symbol: str) -> Dict[str, Any]:
    # Fetch slow fields for a single ticker from Yahoo quoteSummary JSON.
    url = f"https://query1.finance.yahoo.com/v10/finance/quoteSummary/{symbol}"
    params = {
        "modules": "financialData,defaultKeyStatistics,calendarEvents",
    }

    try:
        r = requests.get(url, params=params, timeout=5)
        r.raise_for_status()
        data = r.json()
    except Exception:
        return {
            "Beta": None,
            "1YT": None,
            "Ddate": None,
            "EarningDate": None,
            "RepDiv": None,
        }

    try:
        result = (data.get("quoteSummary", {}).get("result") or [None])[0] or {}
    except Exception:
        result = {}

    fd = result.get("financialData", {}) or {}
    ks = result.get("defaultKeyStatistics", {}) or {}
    ce = result.get("calendarEvents", {}) or {}

    beta = _get_raw(ks, "beta")
    target = _get_raw(fd, "targetMeanPrice")

    # dividend yield as fraction, e.g. 0.034
    div_yield = _get_raw(fd, "dividendYield")

    # ex-dividend date
    ex_div_ts = _get_raw(fd, "exDividendDate")
    if ex_div_ts is None:
        ex_div_ts = _get_raw(ce, "exDividendDate")
    if isinstance(ex_div_ts, (int, float)):
        try:
            ddate = datetime.utcfromtimestamp(ex_div_ts).date()
        except Exception:
            ddate = None
    else:
        ddate = None

    # earnings date - usually list of timestamps in calendarEvents
    earnings_date = None
    earnings = ce.get("earnings", {}) or {}
    ed_list = earnings.get("earningsDate")
    if isinstance(ed_list, list) and ed_list:
        ed0 = ed_list[0]
        ed_raw = ed0.get("raw") if isinstance(ed0, dict) else None
        if isinstance(ed_raw, (int, float)):
            try:
                earnings_date = datetime.utcfromtimestamp(ed_raw).date()
            except Exception:
                earnings_date = None

    return {
        "Beta": beta,
        "1YT": target,
        "Ddate": ddate,
        "EarningDate": earnings_date,
        "RepDiv": div_yield,
    }

import re
from bs4 import BeautifulSoup

HEADERS = {"User-Agent": "Mozilla/5.0"}

def html_fetch_slow_fields(ticker: str):
    """
    Scrape: Beta, 1-year target price, ex-dividend date,
    earnings date, dividend yield.
    Returns: (beta, target, exdiv, earnings, divyield)
    """
    url = f"https://finance.yahoo.com/quote/{ticker}"
    try:
        html = requests.get(url, headers=HEADERS, timeout=10).text
    except:
        return None, None, None, None, None

    soup = BeautifulSoup(html, "lxml")

    def extract(label):
        m = soup.find("span", string=re.compile(label, re.I))
        if not m:
            return None
        v = m.find_next("span")
        return v.text.strip() if v else None

    beta = extract("Beta")
    tgt = extract("1y Target Est")
    exdiv = extract("Ex-Dividend Date")
    earnings = extract("Earnings Date")
    divy = extract("Dividend Yield")

    return beta, tgt, exdiv, earnings, divy
def fetch_slow(tickers: List[str]) -> pd.DataFrame:
    """
    Reliable slow-mode: strictly HTML scraping.
    No yahooquery for slow fields.
    """
    out = {
        "Ticker": [],
        "Beta": [],
        "1YT": [],
        "Ddate": [],
        "EarningDate": [],
        "RepDiv": [],
    }

    for t in tickers:
        if not t:
            # preserve empty tickers
            out["Ticker"].append("")
            out["Beta"].append(None)
            out["1YT"].append(None)
            out["Ddate"].append(None)
            out["EarningDate"].append(None)
            out["RepDiv"].append(None)
            continue

        beta, tgt, exdiv, earnings, divy = html_fetch_slow_fields(t)

        out["Ticker"].append(t)
        out["Beta"].append(beta)
        out["1YT"].append(tgt)
        out["Ddate"].append(exdiv)
        out["EarningDate"].append(earnings)
        out["RepDiv"].append(divy)

    return pd.DataFrame(out)

def write_raw_excel(df: pd.DataFrame):
    """
    Creates DataSource_Raw.xlsx with a clean RTData_Raw sheet + header + table formatting.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = RAW_SHEET

    # Write headers
    for col, name in enumerate(df.columns, start=1):
        c = ws.cell(row=1, column=col, value=name)
        c.font = c.font.copy(bold=True)

    # Write data rows
    for r, row in enumerate(df.itertuples(index=False), start=2):
        for c, value in enumerate(row, start=1):
            ws.cell(row=r, column=c, value=value)

    # Create Excel table
    from openpyxl.worksheet.table import Table, TableStyleInfo

    end_row = len(df) + 1
    end_col = len(df.columns)
    ref = f"A1:{get_column_letter(end_col)}{end_row}"

    table = Table(displayName=RAW_TABLE_NAME, ref=ref)
    style = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    table.tableStyleInfo = style
    ws.add_table(table)

    # Save workbook
    wb.save(RAW_XLSX)
def open_excel_file(path: Path):
    """
    Opens the Excel file on macOS using the 'open' command.
    Safe if file exists, does nothing otherwise.
    """
    try:
        if path.exists():
            subprocess.Popen(["open", str(path)])
        else:
            print(f"⚠ Cannot open — file not found: {path}")
    except Exception as e:
        print(f"⚠ Failed to open Excel file: {e}")
# --------------------------------------------------------------------
# Main
# --------------------------------------------------------------------


def main() -> None:
    parser = argparse.ArgumentParser(description="Live session DataSource → DataSource_Raw pipeline")
    parser.add_argument(
        "--slow",
        action="store_true",
        help="also refresh Beta, 1YT, Ddate, EarningDate, RepDiv via Yahoo JSON API (cached daily)",
    )
    args = parser.parse_args()

    mode = "SLOW" if args.slow else "FAST"
    print(f"🕒 START at {now_str()} — MODE: {mode}")

    t0 = time.time()
    tickers = read_tickers()
    t1 = time.time()
    print(f"⏱ Tickers read in {t1 - t0:.3f} sec")

    # FAST always runs
    tf0 = time.time()
    df_fast = fetch_fast(tickers)
    tf1 = time.time()
    print(f"⏱ Fast fetch in {tf1 - tf0:.3f} sec")

    df = df_fast.copy()

    if args.slow:
        ts0 = time.time()
        df_slow = fetch_slow(tickers)
        print("\nDEBUG -- SLOW df columns:", df_slow.columns.tolist())
        print("DEBUG -- first rows:\n", df_slow.head())
        ts1 = time.time()
        print(f"⏱ Slow fields fetched in {ts1 - ts0:.3f} sec")

        for col in SLOW_COLS:
            if col == "Ticker":
                continue
            df[col] = df_slow[col]

    # Ensure all columns exist
    for col in COLUMNS:
        if col not in df.columns:
            df[col] = None

    tw0 = time.time()
    write_raw_excel(df[COLUMNS])
    tw1 = time.time()
    print(f"⏱ Excel written in {tw1 - tw0:.3f} sec")

    total = time.time() - t0
    print(f"✅ Finished. Total time: {total:.3f} sec")
    print(f"📂 Output: {RAW_XLSX}")

    open_excel_file(RAW_XLSX)


if __name__ == "__main__":
    main()
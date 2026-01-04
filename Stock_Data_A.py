import pandas as pd
from yahooquery import Ticker
from datetime import datetime
import pytz
import os
import time
import traceback

# ============================================================
# SETTINGS
# ============================================================
WORKBOOK = "ASXData.xlsx"
INPUT_SHEET = "Tickers"
OUTPUT_SHEET = "LatestData"
TIMEZONE = pytz.timezone("Australia/Sydney")

CHUNK_SIZE = 20
MAX_RETRIES = 2
RETRY_DELAY = 4

# ============================================================
# VERIFY FILE
# ============================================================
if not os.path.exists(WORKBOOK):
    raise FileNotFoundError(f"Workbook '{WORKBOOK}' not found")

# ============================================================
# READ TICKERS
# ============================================================
tickers_df = pd.read_excel(WORKBOOK, sheet_name=INPUT_SHEET)
col_name = "Ticker" if "Ticker" in tickers_df.columns else "Tickers"

tickers = (
    tickers_df[col_name]
    .astype(str)
    .str.strip()
    .replace("", pd.NA)
    .dropna()
    .tolist()
)

if not tickers:
    raise ValueError("No tickers found")

print(f"✅ {len(tickers)} tickers loaded")

# ============================================================
# UTILITIES
# ============================================================
def chunked(seq, size):
    for i in range(0, len(seq), size):
        yield seq[i:i + size]


def fetch_with_retry(label, getter):
    """Retry YahooQuery call when it returns string or fails"""
    last_error = ""
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            data = getter()
            if isinstance(data, dict):
                return data, ""
            last_error = str(data)
            print(f"⚠️ {label}: attempt {attempt} returned string")
        except Exception as e:
            last_error = str(e)
            print(f"⚠️ {label}: attempt {attempt} failed → {e}")

        if attempt < MAX_RETRIES:
            time.sleep(RETRY_DELAY)

    return None, last_error


def normalize_module(data, module_name, ticker, diagnostics):
    """Guarantee exactly one row per ticker per module"""
    if isinstance(data, dict) and ticker in data and isinstance(data[ticker], dict):
        diagnostics["coverage"] += 1
        return data[ticker]

    diagnostics["missing"].append(module_name)
    return {}

# ============================================================
# FETCH + NORMALIZE (PRESERVE ALL TICKERS)
# ============================================================
rows = []

for chunk in chunked(tickers, CHUNK_SIZE):
    print(f"📡 Fetching chunk: {', '.join(chunk)}")
    t = Ticker(chunk)

    price_data, err_price = fetch_with_retry("price", lambda: t.price)
    summary_data, err_summary = fetch_with_retry("summary", lambda: t.summary_detail)
    calendar_data, err_calendar = fetch_with_retry("calendar", lambda: t.calendar_events)
    financial_data, err_financial = fetch_with_retry("financial", lambda: t.financial_data)

    for ticker in chunk:
        diag = {
            "missing": [],
            "errors": [],
            "coverage": 0
        }

        if err_price: diag["errors"].append("price")
        if err_summary: diag["errors"].append("summary")
        if err_calendar: diag["errors"].append("calendar")
        if err_financial: diag["errors"].append("financial")

        row = {"Ticker": ticker}

        row.update(normalize_module(price_data, "price", ticker, diag))
        row.update(normalize_module(summary_data, "summary", ticker, diag))
        row.update(normalize_module(calendar_data, "calendar", ticker, diag))
        row.update(normalize_module(financial_data, "financial", ticker, diag))

        # ---- Diagnostics ----
        row["DataCoverage"] = diag["coverage"]
        row["MissingModules"] = ",".join(diag["missing"])
        row["YahooError"] = ",".join(diag["errors"])

        if diag["coverage"] == 4:
            row["Status"] = "OK"
        elif diag["coverage"] > 0:
            row["Status"] = "PARTIAL"
        else:
            row["Status"] = "NO_DATA"

        rows.append(row)

    time.sleep(1)

df = pd.DataFrame(rows)

# ============================================================
# CLEAN HEADERS
# ============================================================
rename_map = {
    c: c.replace("regularMarket", "Market")
    for c in df.columns if c.startswith("regularMarket")
}
df.rename(columns=rename_map, inplace=True)

# ============================================================
# TIME CONVERSIONS
# ============================================================
def safe_to_datetime(x):
    try:
        if pd.api.types.is_number(x) or str(x).isdigit():
            return pd.to_datetime(float(x), unit="s", utc=True)
        return pd.to_datetime(x, utc=True, errors="coerce")
    except Exception:
        return pd.NaT

if "MarketTime" in df.columns:
    df["MarketTime"] = df["MarketTime"].apply(safe_to_datetime)
    df["MarketTime"] = df["MarketTime"].dt.tz_convert(TIMEZONE)
    df["MarketTime"] = df["MarketTime"].dt.strftime("%d-%b-%Y %H:%M")

for col in ["exDividendDate", "earningsDate"]:
    if col in df.columns:
        df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%d-%b-%Y")

df = df.fillna("")

# ============================================================
# WRITE TO EXCEL
# ============================================================
with pd.ExcelWriter(WORKBOOK, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    df.to_excel(writer, sheet_name=OUTPUT_SHEET, index=False)

# ============================================================
# SUMMARY
# ============================================================
print("\n=== SUMMARY ===")
print(f"Tickers processed : {len(tickers)}")
print(f"Rows written      : {len(df)}")
print(df["Status"].value_counts())
print(f"Workbook          : {os.path.abspath(WORKBOOK)}")
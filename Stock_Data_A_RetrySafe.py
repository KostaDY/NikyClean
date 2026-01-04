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

CHUNK_SIZE = 20          # Yahoo-safe batch size
MAX_RETRIES = 2          # retry when Yahoo returns strings
RETRY_DELAY = 4          # seconds between retries

# ============================================================
# VERIFY FILE
# ============================================================
if not os.path.exists(WORKBOOK):
    raise FileNotFoundError(
        f"Workbook '{WORKBOOK}' not found.\n"
        "Create it with a sheet 'Tickers' and a column 'Ticker' or 'Tickers'."
    )

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
    raise ValueError(f"No tickers found in sheet '{INPUT_SHEET}'.")

print(f"✅ {len(tickers)} tickers loaded")

# ============================================================
# UTILITIES
# ============================================================
def chunked(seq, size):
    for i in range(0, len(seq), size):
        yield seq[i:i + size]


def dict_to_df(data, name=""):
    """
    Robust YahooQuery → DataFrame converter
    Handles dicts, None, and Yahoo string error payloads
    """
    if data is None:
        print(f"⚠️ {name}: None received")
        return pd.DataFrame(columns=["Ticker"])

    if isinstance(data, str):
        print(f"⚠️ {name}: Yahoo returned error string → {data[:120]}")
        return pd.DataFrame(columns=["Ticker"])

    if isinstance(data, dict):
        try:
            df = pd.DataFrame.from_dict(data, orient="index").reset_index()
            df.rename(columns={"index": "Ticker"}, inplace=True)
            return df
        except Exception as e:
            print(f"⚠️ {name}: dict conversion failed → {e}")
            rows = []
            for k, v in data.items():
                if isinstance(v, dict):
                    rows.append({"Ticker": k, **v})
            return pd.DataFrame(rows)

    print(f"⚠️ {name}: unexpected type → {type(data).__name__}")
    return pd.DataFrame(columns=["Ticker"])


def fetch_with_retry(tickers, label, getter):
    """
    Retry wrapper for YahooQuery modules
    Retries when Yahoo returns string instead of dict
    """
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            data = getter()
            if isinstance(data, dict):
                return data
            print(f"⚠️ {label}: attempt {attempt} returned {type(data).__name__}")
        except Exception as e:
            print(f"⚠️ {label}: attempt {attempt} failed → {e}")

        if attempt < MAX_RETRIES:
            time.sleep(RETRY_DELAY)

    print(f"❌ {label}: giving up after {MAX_RETRIES} attempts")
    return None


# ============================================================
# FETCH DATA (SAFE, CHUNKED, RETRIED)
# ============================================================
price_frames = []
summary_frames = []
calendar_frames = []
financial_frames = []

for chunk in chunked(tickers, CHUNK_SIZE):
    print(f"📡 Fetching chunk: {', '.join(chunk)}")
    t = Ticker(chunk)

    price = fetch_with_retry(chunk, "price", lambda: t.price)
    summary = fetch_with_retry(chunk, "summary_detail", lambda: t.summary_detail)
    calendar = fetch_with_retry(chunk, "calendar_events", lambda: t.calendar_events)
    financial = fetch_with_retry(chunk, "financial_data", lambda: t.financial_data)

    price_frames.append(dict_to_df(price, "price"))
    summary_frames.append(dict_to_df(summary, "summary_detail"))
    calendar_frames.append(dict_to_df(calendar, "calendar_events"))
    financial_frames.append(dict_to_df(financial, "financial_data"))

    time.sleep(1)  # polite pause between chunks

# ============================================================
# MERGE DATASETS
# ============================================================
price_df = pd.concat(price_frames, ignore_index=True)
summary_df = pd.concat(summary_frames, ignore_index=True)
calendar_df = pd.concat(calendar_frames, ignore_index=True)
financial_df = pd.concat(financial_frames, ignore_index=True)

df = price_df.merge(summary_df, on="Ticker", how="left", suffixes=("", "_s"))
df = df.merge(financial_df, on="Ticker", how="left", suffixes=("", "_f"))
df = df.merge(calendar_df, on="Ticker", how="left", suffixes=("", "_c"))

# ============================================================
# CLEAN HEADERS
# ============================================================
rename_map = {
    c: c.replace("regularMarket", "Market")
    for c in df.columns if c.startswith("regularMarket")
}
df.rename(columns=rename_map, inplace=True)

# ============================================================
# SELECT RELEVANT FIELDS
# ============================================================
preferred_fields = [
    "Ticker",
    "MarketPrice",
    "MarketChange",
    "MarketChangePercent",
    "MarketDayHigh",
    "MarketDayLow",
    "MarketPreviousClose",
    "MarketVolume",
    "marketCap",
    "fiftyTwoWeekHigh",
    "fiftyTwoWeekLow",
    "targetMeanPrice",
    "dividendRate",
    "dividendYield",
    "exDividendDate",
    "earningsDate",
    "currency",
    "MarketTime"
]

df = df[[c for c in preferred_fields if c in df.columns]].copy()

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

# ============================================================
# FINAL CLEANUP
# ============================================================
df = (
    df.drop_duplicates(subset=["Ticker"])
      .sort_values("Ticker")
      .reset_index(drop=True)
      .fillna("")
)

# ============================================================
# WRITE TO EXCEL
# ============================================================
try:
    with pd.ExcelWriter(WORKBOOK, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name=OUTPUT_SHEET, index=False)
    print(f"💾 Updated '{OUTPUT_SHEET}' with {len(df)} records")
except Exception:
    print("❌ Excel writing failed:")
    traceback.print_exc()

# ============================================================
# SUMMARY
# ============================================================
count_target = (
    df["targetMeanPrice"].astype(str).replace("", pd.NA).dropna().shape[0]
    if "targetMeanPrice" in df.columns else 0
)

print("\n=== SUMMARY ===")
print(f"Tickers processed : {len(tickers)}")
print(f"Rows written      : {len(df)}")
print(f"Columns exported  : {len(df.columns)}")
print(f"Target prices     : {count_target}")
print(f"Workbook          : {os.path.abspath(WORKBOOK)}")
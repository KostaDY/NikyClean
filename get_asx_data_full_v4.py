import pandas as pd
from yahooquery import Ticker
from datetime import datetime
import pytz
import os

# === SETTINGS ===
workbook = "ASXData.xlsx"          # same workbook for input & output
input_sheet = "Tickers"            # sheet with column 'Ticker' or 'Tickers'
latest_sheet = "LatestData"        # sheet that always shows current snapshot
timezone = pytz.timezone("Australia/Sydney")

# === VERIFY FILE ===
if not os.path.exists(workbook):
    raise FileNotFoundError(
        f"Workbook '{workbook}' not found.\n"
        "Create it with a sheet 'Tickers' and a column 'Ticker' or 'Tickers'."
    )

# === READ TICKERS ===
tickers_df = pd.read_excel(workbook, sheet_name=input_sheet)
col_name = "Ticker" if "Ticker" in tickers_df.columns else "Tickers"
tickers = tickers_df[col_name].dropna().astype(str).tolist()
if not tickers:
    raise ValueError(f"No tickers found in sheet '{input_sheet}'.")
print(f"✅ {len(tickers)} tickers loaded:", ", ".join(tickers))

# === FETCH DATA ===
t = Ticker(tickers)
price_data = t.price
summary = t.summary_detail
calendar = t.calendar_events

# Normalize to DataFrames
def dict_to_df(data):
    if isinstance(data, dict):
        return pd.DataFrame.from_dict(data, orient="index").reset_index().rename(columns={"index": "Ticker"})
    else:
        return data.reset_index().rename(columns={"index": "Ticker"})

price_df = dict_to_df(price_data)
summary_df = dict_to_df(summary)
calendar_df = dict_to_df(calendar)

# Merge all datasets
df = price_df.merge(summary_df, on="Ticker", how="left", suffixes=("", "_s"))
df = df.merge(calendar_df, on="Ticker", how="left", suffixes=("", "_c"))

# === SELECT RELEVANT FIELDS ===
fields = [
    "Ticker",
    "regularMarketPrice",
    "regularMarketChange",
    "regularMarketChangePercent",
    "regularMarketDayHigh",
    "regularMarketDayLow",
    "regularMarketPreviousClose",
    "regularMarketVolume",
    "marketCap",
    "fiftyTwoWeekHigh",
    "fiftyTwoWeekLow",
    "targetMeanPrice",
    "dividendRate",
    "dividendYield",
    "exDividendDate",
    "earningsDate",
    "currency",
    "regularMarketTime"
]
df = df[[f for f in fields if f in df.columns]].copy()

# === TIME CONVERSION ===
def safe_to_datetime(x):
    try:
        if pd.api.types.is_number(x) or str(x).isdigit():
            return pd.to_datetime(float(x), unit="s", utc=True)
        return pd.to_datetime(x, utc=True, errors="coerce")
    except Exception:
        return pd.NaT

if "regularMarketTime" in df.columns:
    df["regularMarketTime"] = df["regularMarketTime"].apply(safe_to_datetime)
    df["regularMarketTime"] = df["regularMarketTime"].dt.tz_convert(timezone)
    df["regularMarketTime"] = df["regularMarketTime"].dt.strftime("%d-%b-%Y %H:%M")

for col in ["exDividendDate", "earningsDate"]:
    if col in df.columns:
        df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%d-%b-%Y")

# === TIMESTAMPED HISTORY SHEET ===
timestamp = datetime.now(timezone).strftime("%Y-%m-%d_%H-%M")
history_sheet = f"Data_{timestamp}"

# === WRITE BOTH SHEETS ===
with pd.ExcelWriter(workbook, engine="openpyxl", mode="a", if_sheet_exists="new") as writer:
    # 1. New timestamped sheet for archive
    df.to_excel(writer, sheet_name=history_sheet, index=False)
    # 2. LatestData sheet: overwrite each run
    df.to_excel(writer, sheet_name=latest_sheet, index=False)

print(f"💾 Added {len(df)} records to '{history_sheet}' and updated '{latest_sheet}' in {os.path.abspath(workbook)}")
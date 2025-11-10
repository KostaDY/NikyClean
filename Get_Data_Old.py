import pandas as pd
from yahooquery import Ticker
from datetime import datetime
import pytz
import os
import traceback

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

# === HELPER: robust dict-to-DataFrame ===
def dict_to_df(data, name=""):
    if not isinstance(data, dict):
        print(f"⚠️ {name}: Expected dict, got {type(data).__name__}. Returning empty DataFrame.")
        return pd.DataFrame(columns=["Ticker"])

    rows = []
    for k, v in data.items():
        # Each value should ideally be a dict of fields
        if isinstance(v, dict):
            rows.append({"Ticker": k, **v})
        else:
            # handle scalars (str, float, None, etc.)
            rows.append({"Ticker": k, "Value": v})
    if not rows:
        print(f"⚠️ {name}: No data rows found.")
        return pd.DataFrame(columns=["Ticker"])
    return pd.DataFrame(rows)

# === Normalize all datasets safely ===
price_df = dict_to_df(price_data, "price_data")
summary_df = dict_to_df(summary, "summary_detail")
calendar_df = dict_to_df(calendar, "calendar_events")

# === Merge all datasets ===
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
    "regularMarketTime",
    "Value",  # in case flat values were stored
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
    try:
        df["regularMarketTime"] = df["regularMarketTime"].apply(safe_to_datetime)
        df["regularMarketTime"] = df["regularMarketTime"].dt.tz_convert(timezone)
        df["regularMarketTime"] = df["regularMarketTime"].dt.strftime("%d-%b-%Y %H:%M")
    except Exception:
        print("⚠️ Time conversion failed for some entries.")

for col in ["exDividendDate", "earningsDate"]:
    if col in df.columns:
        df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%d-%b-%Y")

# === TIMESTAMPED HISTORY SHEET ===
timestamp = datetime.now(timezone).strftime("%Y-%m-%d_%H-%M")
history_sheet = f"Data_{timestamp}"

# === WRITE BOTH SHEETS ===
try:
    with pd.ExcelWriter(workbook, engine="openpyxl", mode="a", if_sheet_exists="new") as writer:
        # 1. New timestamped sheet for archive
        df.to_excel(writer, sheet_name=history_sheet, index=False)
        # 2. LatestData sheet: overwrite each run
        df.to_excel(writer, sheet_name=latest_sheet, index=False)
    print(f"💾 Added {len(df)} records to '{history_sheet}' and updated '{latest_sheet}' in {os.path.abspath(workbook)}")
except Exception as e:
    print("❌ Excel writing failed:")
    traceback.print_exc()
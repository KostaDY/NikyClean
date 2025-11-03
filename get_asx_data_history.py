import pandas as pd
from yahooquery import Ticker
from datetime import datetime
import pytz
import os

# === SETTINGS ===
workbook = "ASXData.xlsx"        # same workbook for input & output
input_sheet = "Tickers"          # sheet containing tickers in column 'Ticker' or 'Tickers'
timezone = pytz.timezone("Australia/Sydney")

# === VERIFY WORKBOOK EXISTS ===
if not os.path.exists(workbook):
    raise FileNotFoundError(
        f"❌ Workbook '{workbook}' not found.\n"
        "Create it with a sheet 'Tickers' and a column 'Ticker' or 'Tickers'."
    )

# === READ TICKERS FROM EXCEL ===
tickers_df = pd.read_excel(workbook, sheet_name=input_sheet)
col_name = "Ticker" if "Ticker" in tickers_df.columns else "Tickers"
tickers = tickers_df[col_name].dropna().astype(str).tolist()

if not tickers:
    raise ValueError(f"❌ No tickers found in sheet '{input_sheet}'.")

print(f"✅ {len(tickers)} tickers loaded:", ", ".join(tickers))

# === FETCH DATA FROM YAHOO FINANCE ===
t = Ticker(tickers)
data = t.price

# --- normalize whether dict or DataFrame ---
if isinstance(data, dict):
    df = pd.DataFrame.from_dict(data, orient="index")
else:
    df = data

df.reset_index(inplace=True)
df.rename(columns={"index": "symbol"}, inplace=True)

# === KEEP RELEVANT FIELDS ===
cols = ["symbol", "regularMarketPrice", "currency", "regularMarketTime"]
df = df[[c for c in cols if c in df.columns]].copy()

# === CONVERT TIME TO LOCAL (Sydney) ===
if "regularMarketTime" in df.columns:
    def safe_to_datetime(x):
        try:
            # numeric UNIX timestamp
            if pd.api.types.is_number(x) or str(x).isdigit():
                return pd.to_datetime(float(x), unit="s", utc=True)
            # formatted datetime string
            return pd.to_datetime(x, utc=True, errors="coerce")
        except Exception:
            return pd.NaT

    df["regularMarketTime"] = df["regularMarketTime"].apply(safe_to_datetime)
    df["regularMarketTime"] = df["regularMarketTime"].dt.tz_convert(timezone)
    df["regularMarketTime"] = df["regularMarketTime"].dt.strftime("%d-%b-%Y %H:%M")

# === RENAME FOR OUTPUT ===
df.rename(columns={
    "symbol": "Ticker",
    "regularMarketPrice": "Price",
    "currency": "Currency",
    "regularMarketTime": "MarketTime"
}, inplace=True)

# === DETERMINE TIMESTAMPED SHEET NAME ===
timestamp = datetime.now(timezone).strftime("%Y-%m-%d_%H-%M")
output_sheet = f"Data_{timestamp}"

# === WRITE RESULTS TO NEW SHEET IN SAME WORKBOOK ===
with pd.ExcelWriter(workbook, engine="openpyxl", mode="a", if_sheet_exists="new") as writer:
    df.to_excel(writer, sheet_name=output_sheet, index=False)

print(f"💾 Added {len(df)} records to sheet '{output_sheet}' in {os.path.abspath(workbook)}")
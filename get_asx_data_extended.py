import pandas as pd
from yahooquery import Ticker
from datetime import datetime
import pytz
import os

# === SETTINGS ===
workbook = "ASXData.xlsx"
input_sheet = "Tickers"
timezone = pytz.timezone("Australia/Sydney")

# === VERIFY FILE ===
if not os.path.exists(workbook):
    raise FileNotFoundError(f"Workbook '{workbook}' not found.")

tickers_df = pd.read_excel(workbook, sheet_name=input_sheet)
col_name = "Ticker" if "Ticker" in tickers_df.columns else "Tickers"
tickers = tickers_df[col_name].dropna().astype(str).tolist()

print(f"✅ {len(tickers)} tickers loaded:", ", ".join(tickers))

# === FETCH DATA ===
t = Ticker(tickers)
data = t.price
df = pd.DataFrame.from_dict(data, orient="index").reset_index().rename(columns={"index": "Ticker"})

# === FIELDS TO KEEP ===
fields = [
    "regularMarketPrice",
    "regularMarketChange",
    "regularMarketChangePercent",
    "regularMarketOpen",
    "regularMarketDayHigh",
    "regularMarketDayLow",
    "regularMarketPreviousClose",
    "regularMarketVolume",
    "marketCap",
    "fiftyTwoWeekHigh",
    "fiftyTwoWeekLow",
    "currency",
    "regularMarketTime"
]
df = df[["Ticker"] + [f for f in fields if f in df.columns]]

# === CONVERT TIME TO LOCAL (Sydney) ===
def safe_to_datetime(x):
    try:
        if pd.api.types.is_number(x) or str(x).isdigit():
            return pd.to_datetime(float(x), unit="s", utc=True)
        return pd.to_datetime(x, utc=True, errors="coerce")
    except Exception:
        return pd.NaT

df["regularMarketTime"] = df["regularMarketTime"].apply(safe_to_datetime)
df["regularMarketTime"] = df["regularMarketTime"].dt.tz_convert(timezone)
df["regularMarketTime"] = df["regularMarketTime"].dt.strftime("%d-%b-%Y %H:%M")

# === TIMESTAMPED OUTPUT SHEET ===
timestamp = datetime.now(timezone).strftime("%Y-%m-%d_%H-%M")
output_sheet = f"Data_{timestamp}"

# === WRITE RESULTS ===
with pd.ExcelWriter(workbook, engine="openpyxl", mode="a", if_sheet_exists="new") as writer:
    df.to_excel(writer, sheet_name=output_sheet, index=False)

print(f"💾 Added {len(df)} records to sheet '{output_sheet}' in {os.path.abspath(workbook)}")
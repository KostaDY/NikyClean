import pandas as pd
from yahooquery import Ticker
from datetime import datetime
import pytz
import os
import traceback

# === SETTINGS ===
workbook = "ASXData.xlsx"          # same workbook for input & output
input_sheet = "Tickers"            # sheet with column 'Ticker' or 'Tickers'
latest_sheet = "LatestData"        # only this sheet will be written
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
financial = t.financial_data   # <-- includes 1-Year Target Price and other metrics

# === Safe dict → DataFrame converter ===
def dict_to_df(data, name=""):
    if not isinstance(data, dict):
        print(f"⚠️ {name}: Expected dict, got {type(data).__name__}")
        return pd.DataFrame(columns=["Ticker"])
    try:
        df = pd.DataFrame.from_dict(data, orient="index").reset_index()
        df.rename(columns={"index": "Ticker"}, inplace=True)
        return df
    except Exception as e:
        print(f"⚠️ {name}: conversion failed → {e}")
        rows = []
        for k, v in data.items():
            if isinstance(v, dict):
                rows.append({"Ticker": k, **v})
        return pd.DataFrame(rows)

# === Convert all YahooQuery modules ===
price_df = dict_to_df(price_data, "price_data")
summary_df = dict_to_df(summary, "summary_detail")
calendar_df = dict_to_df(calendar, "calendar_events")
financial_df = dict_to_df(financial, "financial_data")

# === Merge all datasets ===
df = price_df.merge(summary_df, on="Ticker", how="left", suffixes=("", "_s"))
df = df.merge(financial_df, on="Ticker", how="left", suffixes=("", "_f"))
df = df.merge(calendar_df, on="Ticker", how="left", suffixes=("", "_c"))

# === Clean headers ===
rename_map = {c: c.replace("regularMarket", "Market") for c in df.columns if c.startswith("regularMarket")}
df.rename(columns=rename_map, inplace=True)

# === KEEP ONLY RELEVANT FIELDS (trim redundant noise) ===
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
    "targetMeanPrice",       # ✅ 1-Year Target
    "dividendRate",
    "dividendYield",
    "exDividendDate",
    "earningsDate",
    "currency",
    "MarketTime"
]
available = [f for f in preferred_fields if f in df.columns]
df = df[available].copy()

# === CONVERT TIME FIELDS ===
def safe_to_datetime(x):
    try:
        if pd.api.types.is_number(x) or str(x).isdigit():
            return pd.to_datetime(float(x), unit="s", utc=True)
        return pd.to_datetime(x, utc=True, errors="coerce")
    except Exception:
        return pd.NaT

if "MarketTime" in df.columns:
    df["MarketTime"] = df["MarketTime"].apply(safe_to_datetime)
    df["MarketTime"] = df["MarketTime"].dt.tz_convert(timezone)
    df["MarketTime"] = df["MarketTime"].dt.strftime("%d-%b-%Y %H:%M")

for col in ["exDividendDate", "earningsDate"]:
    if col in df.columns:
        df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%d-%b-%Y")

# === FINAL CLEANUP ===
df = df.drop_duplicates(subset=["Ticker"]).sort_values("Ticker").reset_index(drop=True)
df.columns = [str(c).strip() for c in df.columns]
df = df.fillna("")

# === WRITE TO EXCEL ===
try:
    with pd.ExcelWriter(workbook, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name=latest_sheet, index=False)
    print(f"💾 Updated '{latest_sheet}' with {len(df)} records and {len(df.columns)} columns in {os.path.abspath(workbook)}")
except Exception:
    print("❌ Excel writing failed:")
    traceback.print_exc()

# === SUMMARY ===
count_target = df["targetMeanPrice"].astype(str).replace("", pd.NA).dropna().shape[0] if "targetMeanPrice" in df.columns else 0
print("\n=== SUMMARY ===")
print(f"Tickers processed : {len(tickers)}")
print(f"Rows written      : {len(df)}")
print(f"Columns exported  : {len(df.columns)}")
print(f"Target price data : {count_target} tickers with 1-Year Target")
print(f"Output path       : {os.path.abspath(workbook)}")
import time
from datetime import datetime
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import numbers

# Paths
DATASOURCE = "excel/DataSource.xlsx"
RAWFILE = "excel/DataSource_Raw.xlsx"

# Constants
DATASHEET = "DataRT"
DATATABLE = "RTData"
RAW_SHEET = "RTData_Raw"


# ================================================================
# Helper: read tickers from DataSource.xlsx → RTData table
# ================================================================
def read_tickers():
    print("📥 Reading tickers from DataSource.xlsx → RTData…")

    wb = load_workbook(DATASOURCE, data_only=True)
    ws = wb[DATASHEET]

    # Identify table RTData
    table = None
    for t in ws._tables.values():
        if t.name == DATATABLE:
            table = t
            break

    if table is None:
        raise RuntimeError("❌ Could not find table RTData in DataSource.xlsx")

    ref = table.ref  # Example: "A2:AF96"
    cells = ws[ref]   # <-- returns a 2D tuple ((row1cells),(row2cells),...)

    header_row = cells[0]               # first row
    data_rows = cells[1:]               # all remaining rows

    # Extract column names
    headers = [c.value for c in header_row]

    # Locate Ticker_Symbol
    try:
        ticker_idx = headers.index("Ticker_Symbol")
    except ValueError:
        raise RuntimeError("❌ Column 'Ticker_Symbol' not found in RTData table.")

    # Extract values from that column
    tickers = [row[ticker_idx].value for row in data_rows]

    print(f"✔ Loaded {len(tickers)} tickers (empty tickers kept).")
    return tickers


# ================================================================
# Download market data (placeholder synthetic data generator)
# REPLACE THIS WITH YOUR REAL FETCH FUNCTION
# ================================================================
def fetch_live_data(tickers):
    """
    Produces a DataFrame with correct columns.
    Replace internals with real API logic.
    """
    rows = []
    for t in tickers:
        if not t:
            rows.append({c: "" for c in [
                "Ticker","Close","Open","Last","Low","High",
                "PE","Change","ChangePct","Volume","VolumeAverage",
                "Beta","OneYearTarget","DividendDate","EarningsDate","Dividend"
            ]})
            continue

        # Demo dummy values:
        rows.append({
            "Ticker": t,
            "Close": 100.55,
            "Open": 101.11,
            "Last": 100.78,
            "Low": 98.44,
            "High": 102.33,
            "PE": 18.4,
            "Change": -1.22,
            "ChangePct": 0.034,      # 3.4%
            "Volume": 4_532_111,
            "VolumeAverage": 2_931_000,
            "Beta": 1.18,
            "OneYearTarget": 130.50,
            "DividendDate": "2025-11-06",
            "EarningsDate": "2025-12-10",
            "Dividend": 0.012        # 1.2%
        })

    return pd.DataFrame(rows)


# ================================================================
# Create fresh new workbook (safe, no corruption)
# ================================================================
def create_clean_raw_workbook():
    wb = Workbook()
    ws = wb.active
    ws.title = RAW_SHEET

    # This is a *clean* new file, created every run.
    wb.save(RAWFILE)


# ================================================================
# Write formatted DataFrame to workbook
# ================================================================
def write_raw_data(df):
    print("📝 Writing RTData_Raw…")

    # Always create a clean workbook first
    create_clean_raw_workbook()

    wb = load_workbook(RAWFILE)
    ws = wb[RAW_SHEET]

    # Write headers + rows
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    # Apply Excel formats
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row_dict = {ws.cell(row=1, column=i).value: cell for i, cell in enumerate(row, start=1)}

        # Percent columns
        if "ChangePct" in row_dict:
            row_dict["ChangePct"].number_format = "0.00%"

        if "Dividend" in row_dict:
            row_dict["Dividend"].number_format = "0.00%"

        # Dates
        for d in ["DividendDate", "EarningsDate"]:
            if d in row_dict:
                c = row_dict[d]
                if isinstance(c.value, str) and "-" in c.value:
                    try:
                        c.value = datetime.fromisoformat(c.value)
                        c.number_format = "mmm d, yyyy"
                    except Exception:
                        pass

        # Integers (Volume, VolumeAverage)
        for f in ["Volume", "VolumeAverage"]:
            if f in row_dict:
                row_dict[f].number_format = "#,##0"

        # Floats
        for f in ["Close", "Open", "Last", "Low", "High", "PE", "Beta",
                  "OneYearTarget", "Change"]:
            if f in row_dict:
                row_dict[f].number_format = "0.00"

    wb.save(RAWFILE)


# ================================================================
# Main pipeline with performance timestamps
# ================================================================
def main():
    t0 = time.time()
    print("🕒 Starting pipeline at", datetime.now().strftime("%H:%M:%S"))

    # Step 1: read tickers
    t1 = time.time()
    tickers = read_tickers()
    print(f"⏱ Tickers loaded in {t1 - t0:.3f} sec")

    # Step 2: download/fetch data
    t2 = time.time()
    df = fetch_live_data(tickers)
    print(f"⏱ Data fetched in {t2 - t1:.3f} sec")

    # Step 3: write Excel
    t3 = time.time()
    write_raw_data(df)
    print(f"⏱ Excel written in {t3 - t2:.3f} sec")

    print(f"✔ Finished. Total time: {t3 - t0:.3f} sec")


# ================================================================
if __name__ == "__main__":
    main()
import pandas as pd
from pathlib import Path

def read_medy(path, sheet_name="DataRT"):
    """
    Read quasi-real-time values from DataSource.xlsx/DataRT.

    The real header is on Excel row 2 (pandas header=1).
    """

    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"DataSource not found: {path}")

    # Critical: row 2 contains actual column headers
    df = pd.read_excel(
        p,
        sheet_name=sheet_name,
        header=1
    )

    # Clean column names (trim whitespace)
    df.columns = [str(c).strip() for c in df.columns]

    # Relevant columns for Live Session
    rename_map = {
        "Ticker_Symbol": "Ticker",
        "M_Price": "Last",
        "Last_trade_time": "Time",
        "Change%": "ChangePct",
    }

    # Apply renaming only for existing columns
    df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns},
              inplace=True)

    # Verify required columns
    if "Ticker" not in df.columns:
        raise KeyError("Column 'Ticker_Symbol' missing after header=1")
    if "Last" not in df.columns:
        raise KeyError("Column 'M-Price' missing after header=1")

    # Convert Last to numeric
    df["Last"] = pd.to_numeric(df["Last"], errors="coerce")

    # Keep only rows with valid ticker and price
    df = df[df["Ticker"].notna() & df["Last"].notna()].copy()

    # Reset index for clean iteration
    df.reset_index(drop=True, inplace=True)

    return df
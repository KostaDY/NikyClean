import time
import pandas as pd
from pathlib import Path
from medy_reader import read_medy

# --------------------
# Paths
# --------------------
ROOT = Path(__file__).resolve().parents[1]
DATASOURCE = ROOT / "excel" / "DataSource.xlsx"
LIVE_DIR = ROOT / "live"
LIVE_DIR.mkdir(exist_ok=True)
LIVE_FILE = LIVE_DIR / "live_values.csv"

# --------------------
# Helpers
# --------------------
def export_snapshot(snapshot: dict):
    """Write live_values.csv from snapshot dictionary"""
    rows = []
    for key, (value, vtype) in snapshot.items():
        rows.append({"key": key, "value": value, "type": vtype})

    df = pd.DataFrame(rows)
    df.to_csv(LIVE_FILE, index=False)


# --------------------
# Build snapshot from DataSource
# --------------------
def build_snapshot():
    """Return dict of (key -> (value, type))"""
    df = read_medy(DATASOURCE)

    snapshot = {}

    for _, row in df.iterrows():
        ticker = row["Ticker"]

        # Real-time last price
        if "Last" in row:
            snapshot[f"{ticker}.price"] = (row["Last"], "number")

        # Change %
        if "ChangePct" in row:
            snapshot[f"{ticker}.change_pct"] = (row["ChangePct"], "number")

        # Last trade time
        if "Time" in row:
            snapshot[f"{ticker}.time"] = (str(row["Time"]), "text")

        # Currency
        if "Currency" in row:
            snapshot[f"{ticker}.currency"] = (row["Currency"], "text")

    return snapshot


# --------------------
# Live loop
# --------------------
def live_loop(interval=5):
    print("Starting live session. Press Ctrl+C to stop.")
    while True:
        try:
            snapshot = build_snapshot()
            export_snapshot(snapshot)
            print(f"Updated {len(snapshot)} live keys.")
            time.sleep(interval)
        except KeyboardInterrupt:
            print("Live session stopped.")
            break
        except Exception as e:
            print("ERROR:", e)
            time.sleep(interval)


# --------------------
# CLI entry
# --------------------
if __name__ == "__main__":
    live_loop(interval=3)
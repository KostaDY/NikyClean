#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
PHASE 1 – 5-STATE VERSION

Build 400-digit sequence using 5 states:

1 : <= -2%
2 : (-2%, -0.5%]
3 : (-0.5%, 0.5%)
4 : [0.5%, 2%)
5 : >= 2%

Outputs:
- digits_400_5_history.csv
- digits_400_5.txt
- digits_400_5_state.csv
"""

from datetime import datetime, UTC
from pathlib import Path
import pandas as pd

from markov_core import read_tickers_from_excel, download_closes, XLSX

# ============================================================
# SETTINGS
# ============================================================

N_DIGITS = 400
DOWNLOAD_PERIOD = "900d"

OUT_HISTORY = Path("digits_400_5_history.csv")
OUT_TXT = Path("digits_400_5.txt")
OUT_STATE = Path("digits_400_5_state.csv")


# ============================================================
# 5-STATE QUANTIZATION
# ============================================================

def quantize_5(pct: float) -> str:
    if pct <= -2.0:
        return "1"
    elif pct <= -0.5:
        return "2"
    elif pct < 0.5:
        return "3"
    elif pct < 2.0:
        return "4"
    else:
        return "5"


# ============================================================
# BUILD DIGITS
# ============================================================

def build_digits_5(closes: pd.Series, n_digits: int = 400):

    need = n_digits + 1
    closes = closes.dropna()

    if len(closes) < need:
        raise RuntimeError(f"Need {need} closes, got {len(closes)}")

    closes_tail = closes.tail(need)

    df = pd.DataFrame({"Close": closes_tail})
    df["PctChange"] = df["Close"].pct_change() * 100.0
    df = df.iloc[1:].copy()

    df["Digit"] = df["PctChange"].apply(lambda x: quantize_5(float(x)))
    df["Date"] = pd.to_datetime(df.index).date.astype(str)

    digits_chrono = df["Digit"].tolist()
    digits_youngest_first = "".join(digits_chrono[::-1])

    return digits_youngest_first, df[["Date", "Close", "PctChange", "Digit"]]


# ============================================================
# MAIN
# ============================================================

def main():

    tickers = read_tickers_from_excel(XLSX)

    all_rows = []
    txt_lines = [f"GeneratedUTC,{datetime.now(UTC).isoformat()}"]
    state_rows = []

    for t in tickers:

        closes = download_closes(t, period=DOWNLOAD_PERIOD)
        digits_youngest_first, df = build_digits_5(closes, N_DIGITS)

        df2 = df.copy()
        df2.insert(0, "Ticker", t)
        all_rows.append(df2)

        txt_lines.append(f"{t},{digits_youngest_first}")

        last_date = df["Date"].iloc[-1]

        state_rows.append({
            "Ticker": t,
            "LastDigitDate": last_date,
            "Digits": digits_youngest_first
        })

    hist = pd.concat(all_rows, ignore_index=True)
    hist.to_csv(OUT_HISTORY, index=False)

    OUT_TXT.write_text("\n".join(txt_lines) + "\n", encoding="utf-8")
    pd.DataFrame(state_rows).to_csv(OUT_STATE, index=False)

    print("5-state 400-digit build completed.")


if __name__ == "__main__":
    main()
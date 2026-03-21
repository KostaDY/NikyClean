#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

from datetime import datetime, UTC
from pathlib import Path

import pandas as pd

from markov_core import read_tickers_from_excel, download_closes, quantize, XLSX

STATE = Path("digits_400_state.csv")
OUT_TXT = Path("digits_400.txt")

N_DIGITS = 400
LOOKBACK_PERIOD = "20d"  # enough to cover weekends/holidays most of the time


def get_latest_digit_and_date(ticker: str):
    closes = download_closes(ticker, period=LOOKBACK_PERIOD)
    if len(closes) < 2:
        raise RuntimeError(f"{ticker}: insufficient closes for latest digit")

    last_close = float(closes.iloc[-1])
    prev_close = float(closes.iloc[-2])

    pct = (last_close / prev_close - 1.0) * 100.0
    digit = quantize(pct)

    last_dt = pd.to_datetime(closes.index[-1]).date().isoformat()
    return digit, last_dt


def main():
    tickers = read_tickers_from_excel(XLSX)

    if not STATE.exists():
        raise RuntimeError("digits_400_state.csv not found. Run phase1 first.")

    df_state = pd.read_csv(STATE, dtype={"Ticker": str, "LastDigitDate": str, "Digits": str})
    df_state = df_state.set_index("Ticker")

    updated = 0
    errors = []

    for t in tickers:
        if t not in df_state.index:
            errors.append(f"{t}: missing from state (run phase1 again)")
            continue

        digits = str(df_state.at[t, "Digits"])
        last_date = str(df_state.at[t, "LastDigitDate"])

        if len(digits) != N_DIGITS or any(ch not in "123" for ch in digits):
            errors.append(f"{t}: bad Digits in state (len={len(digits)})")
            continue

        try:
            dig, dig_date = get_latest_digit_and_date(t)
        except Exception as e:
            errors.append(f"{t}: update failed ({e})")
            continue

        # Update only once per trading day
        if dig_date != last_date:
            digits = dig + digits
            digits = digits[:N_DIGITS]  # keep newest 400 digits
            df_state.at[t, "Digits"] = digits
            df_state.at[t, "LastDigitDate"] = dig_date
            updated += 1

    # Write updated state
    df_state.reset_index().to_csv(STATE, index=False)

    # Write txt snapshot
    lines = [f"GeneratedUTC,{datetime.now(UTC).isoformat()}"]
    for t in tickers:
        if t in df_state.index:
            lines.append(f"{t},{df_state.at[t,'Digits']}")
    OUT_TXT.write_text("\n".join(lines) + "\n", encoding="utf-8")

    print(f"OK: updated {updated} tickers. Wrote {STATE} and {OUT_TXT}.")
    if errors:
        print("Warnings:")
        for e in errors[:30]:
            print(" -", e)
        if len(errors) > 30:
            print(f" - ... {len(errors)-30} more")


if __name__ == "__main__":
    main()
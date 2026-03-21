#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

from datetime import datetime, UTC
from pathlib import Path

import pandas as pd

from markov_core import read_tickers_from_excel, download_closes, build_digits_from_closes, XLSX

OUT_HISTORY = Path("digits_400_history.csv")
OUT_TXT = Path("digits_400.txt")
OUT_STATE = Path("digits_400_state.csv")

N_DIGITS = 400
DOWNLOAD_PERIOD = "900d"  # safe to get >=401 trading closes most of the time


def main():
    tickers = read_tickers_from_excel(XLSX)

    all_rows = []
    txt_lines = [f"GeneratedUTC,{datetime.now(UTC).isoformat()}"]
    state_rows = []

    for t in tickers:
        closes = download_closes(t, period=DOWNLOAD_PERIOD)
        digits_youngest_first, df = build_digits_from_closes(closes, n_digits=N_DIGITS)

        # history rows (chronological)
        df2 = df.copy()
        df2.insert(0, "Ticker", t)
        all_rows.append(df2)

        # txt
        txt_lines.append(f"{t},{digits_youngest_first}")

        # compact state: last close date used for the youngest digit
        last_date = df["Date"].iloc[-1]  # last pct-change date (youngest digit date)
        state_rows.append({"Ticker": t, "LastDigitDate": last_date, "Digits": digits_youngest_first})

    hist = pd.concat(all_rows, ignore_index=True)
    hist.to_csv(OUT_HISTORY, index=False)

    OUT_TXT.write_text("\n".join(txt_lines) + "\n", encoding="utf-8")
    pd.DataFrame(state_rows).to_csv(OUT_STATE, index=False)

    print(f"OK: wrote {OUT_HISTORY}, {OUT_TXT}, {OUT_STATE}")


if __name__ == "__main__":
    main()
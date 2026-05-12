#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from pathlib import Path
import re
import pandas as pd
import xlwings as xw


CSV_PATH = Path("/Users/kostayanev/NikyClean/TradeArchive.csv")
OUT_XLSX = Path("/Users/kostayanev/NikyClean/TradeTable.xlsx")

TRADE_SHEET = "TradeTable"
MARKET_SHEET = "MARKET"


def parse_market_price(value):
    if value is None:
        return None

    if isinstance(value, (int, float)):
        return float(value)

    s = str(value).strip()
    s = re.sub(r"[^\d,.\-]", "", s)

    if not s:
        return None

    if "," in s and "." in s:
        s = s.replace(",", "")
    elif "," in s and "." not in s:
        s = s.replace(",", ".")

    try:
        return float(s)
    except ValueError:
        return None


def load_trades(csv_path: Path) -> pd.DataFrame:
    df = pd.read_csv(
        csv_path,
        header=None,
        names=["Date", "Ticker", "Number", "Price"]
    )

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Ticker"] = df["Ticker"].astype(str).str.strip()
    df["Number"] = pd.to_numeric(df["Number"], errors="coerce")
    df["Price"] = pd.to_numeric(df["Price"], errors="coerce")

    df = df.dropna(subset=["Date", "Ticker", "Number", "Price"])
    df = df[df["Ticker"] != ""]
    df = df[df["Number"] != 0]
    df = df[df["Price"] != 0]

    return df.reset_index(drop=True)


def apply_sales_for_ticker(tdf: pd.DataFrame):
    lots = []
    sales = []

    for _, row in tdf.iterrows():
        if row["Price"] < 0:
            lots.append({
                "date": row["Date"],
                "number": abs(float(row["Number"])),
                "price": float(row["Price"])
            })
        else:
            sales.append({
                "number": abs(float(row["Number"])),
                "price": float(row["Price"])
            })

    for sale in sales:
        sale_qty = sale["number"]
        sale_price_abs = abs(sale["price"])

        while sale_qty > 0 and lots:

            lower_or_equal = [
                i for i, lot in enumerate(lots)
                if abs(lot["price"]) <= sale_price_abs
            ]

            if lower_or_equal:
                best_i = max(
                    lower_or_equal,
                    key=lambda i: abs(lots[i]["price"])
                )
            else:
                best_i = min(
                    range(len(lots)),
                    key=lambda i: abs(lots[i]["price"])
                )

            lot = lots[best_i]
            matched_qty = min(sale_qty, lot["number"])

            lot["number"] -= matched_qty
            sale_qty -= matched_qty

            if lot["number"] <= 1e-12:
                lots.pop(best_i)

        # If sales exceed total purchases, the rest is simply ignored.
        # No warning, because the requested output is remaining purchase lots.

    return lots


def build_trade_table(df: pd.DataFrame) -> dict:
    result = {}

    for ticker, tdf in df.groupby("Ticker", sort=True):
        result[ticker] = apply_sales_for_ticker(tdf)

    return result


def read_market_prices(wb):
    ws = wb.sheets[MARKET_SHEET]

    last_row = ws.range("A" + str(ws.cells.last_cell.row)).end("up").row

    data = ws.range(f"A2:C{last_row}").value

    market = {}

    if not isinstance(data, list):
        return market

    for row in data:
        if not row:
            continue

        ticker = str(row[0]).strip() if row[0] is not None else ""
        price = parse_market_price(row[2]) if len(row) >= 3 else None

        if ticker:
            market[ticker] = price

    return market


def pad_rows(rows, width):
    return [
        row + [""] * (width - len(row))
        for row in rows
    ]


def write_trade_table(result: dict, market: dict, wb):
    sheet_names = [s.name for s in wb.sheets]

    if TRADE_SHEET in sheet_names:
        ws = wb.sheets[TRADE_SHEET]
        ws.clear()
    else:
        ws = wb.sheets.add(TRADE_SHEET, before=wb.sheets[0])

    max_lots = max((len(lots) for lots in result.values()), default=0)

    headers = ["Ticker", "Mkt"]

    for i in range(1, max_lots + 1):
        headers += [f"D{i}", f"N{i}", f"P{i}"]

    rows = [headers]

    for ticker, lots in result.items():
        row = [ticker, market.get(ticker)]

        for lot in lots:
            row += [
                lot["date"].date(),
                lot["number"],
                lot["price"]
            ]

        rows.append(row)

    rows = pad_rows(rows, len(headers))

    ws.range("A1").value = rows

    last_row = len(rows)
    last_col = len(headers)

    ws.range((1, 1), (1, last_col)).font.bold = True
    ws.range((1, 1), (1, last_col)).api.HorizontalAlignment = -4108

    ws.range("A:A").column_width = 13
    ws.range("B:B").column_width = 12

    for col in range(3, last_col + 1):
        ws.range((1, col)).column_width = 11

    for lot_idx in range(max_lots):
        d_col = 3 + lot_idx * 3
        n_col = d_col + 1
        p_col = d_col + 2

        ws.range((2, d_col), (last_row, d_col)).number_format = "yyyy-mm-dd"
        ws.range((2, n_col), (last_row, n_col)).number_format = "0"
        ws.range((2, p_col), (last_row, p_col)).number_format = "0.00"

    ws.range((2, 2), (last_row, 2)).number_format = "0.00"

    # Green formatting
    for r in range(2, last_row + 1):
        mkt = ws.range((r, 2)).value

        if not isinstance(mkt, (int, float)):
            continue

        for lot_idx in range(max_lots):
            d_col = 3 + lot_idx * 3
            p_col = d_col + 2

            buy_price = ws.range((r, p_col)).value

            if isinstance(buy_price, (int, float)) and abs(buy_price) < mkt:
                rng = ws.range((r, d_col), (r, p_col))
                rng.color = (198, 239, 206)
                rng.font.bold = True
                rng.font.color = (0, 97, 0)

    ws.activate()
    ws.range("C2").select()
    wb.app.api.ActiveWindow.FreezePanes = True

    ws.autofit()


def main():
    df = load_trades(CSV_PATH)
    result = build_trade_table(df)

    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        wb = app.books.open(str(OUT_XLSX))

        market = read_market_prices(wb)
        write_trade_table(result, market, wb)

        wb.save()
        wb.close()

    finally:
        app.quit()

    print(f"✅ Updated: {OUT_XLSX}")
    print("MARKET sheet was not modified.")
    print("Market prices were read from MARKET column C.")
    print("Sales matching is price-based, not chronology-based.")


if __name__ == "__main__":
    main()
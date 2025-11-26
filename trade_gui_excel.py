#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
TRADES – Excel-driven GUI for tabular data in TRADES.xlsx (Excel 365 macOS)

Workbook: TRADES.xlsx (keep it CLOSED while running this script)
Sheets:
  - Transactions: tabular data starting at A1 with headers:
        A: Date
        B: Prefix
        C: Ticker
        D: Number
        E: Price
        F: Act
        G+: formula columns (Amount, Day, Day_Buys, ..., Sale_Aggr.)
    We treat the contiguous block from row 1 down where A–F have data as the "table".

  - Stock: tickers in A2:A... (non-empty)
  - Log (optional): log entries [Timestamp, Action, Details]

Operations:
  - Add: append row under last data row in Transactions:
         * writes A–F (Date, Prefix, Ticker, Number, Price, Act='add')
         * copies formulas in G..last_col from previous row
  - Delete: delete first row matching (Date, Prefix, Ticker, Number, Price)
            by deleting that entire sheet row
  - Open WB: show TRADES.xlsx in Excel
  - Exit: save & close workbook + quit Excel instance

All edits are done via xlwings (Excel engine), no openpyxl, no XML corruption.
"""

import os
from datetime import datetime, date
import tkinter as tk
from tkinter import ttk, messagebox
import subprocess

import xlwings as xw  # pip install xlwings

# -------------------------------------------------
# CONFIG
# -------------------------------------------------

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, "TRADES.xlsx")

TRANS_SHEET = "Transactions"
STOCK_SHEET = "Stock"
LOG_SHEET = "Log"

# Globals for Excel instance
_app = None
_wb = None


# -------------------------------------------------
# EXCEL CONNECTION HELPERS
# -------------------------------------------------

def debug(msg: str):
    print("[DEBUG]", msg)


def ensure_workbook():
    """
    Ensure we have an xlwings App and the TRADES.xlsx workbook open.
    Returns (app, wb).
    """
    global _app, _wb

    if _app is None or _wb is None:
        debug(f"Opening Excel and workbook: {EXCEL_FILE}")
        _app = xw.App(visible=False, add_book=False)
        _wb = _app.books.open(EXCEL_FILE)

    return _app, _wb


def get_ws(ws_name):
    _, wb = ensure_workbook()
    try:
        return wb.sheets[ws_name]
    except Exception as e:
        raise KeyError(f"Sheet '{ws_name}' not found in TRADES.xlsx ({e})")


def excel_save():
    """Save TRADES.xlsx."""
    _, wb = ensure_workbook()
    wb.save(EXCEL_FILE)
    debug("Workbook saved.")


def excel_quit():
    """Save, close workbook and quit Excel instance."""
    global _app, _wb
    if _wb is not None:
        try:
            _wb.save(EXCEL_FILE)
            _wb.close()
        except Exception as e:
            debug(f"Error closing workbook: {e}")
    if _app is not None:
        try:
            _app.quit()
        except Exception as e:
            debug(f"Error quitting Excel app: {e}")
    _app = None
    _wb = None


# -------------------------------------------------
# LOGGING
# -------------------------------------------------

def append_log_entry(action: str, details: str):
    """
    Append an entry [Timestamp, Action, Details] to Log sheet if it exists.
    """
    try:
        ws_log = get_ws(LOG_SHEET)
    except KeyError:
        return
    except Exception as e:
        debug(f"append_log_entry: cannot access Log sheet: {e}")
        return

    last_row = ws_log.cells.last_cell.row
    r = last_row
    while r >= 1 and ws_log.range(r, 1).value in (None, ""):
        r -= 1
    new_row = r + 1

    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws_log.range(new_row, 1).value = ts
    ws_log.range(new_row, 2).value = action
    ws_log.range(new_row, 3).value = details

    excel_save()


# -------------------------------------------------
# TABLE (RANGE) DIMENSIONS
# -------------------------------------------------

def get_transactions_table_dims():
    """
    Determine the "table" on Transactions sheet as:

    - Header row: 1
    - Data rows: 2..last_data_row
      where last_data_row is last row where any of A..F has data.
    - Last column: last used column in row 1 (header row).

    Returns (ws, first_data_row, last_data_row, last_col).
    If no data rows, last_data_row == 1.
    """
    ws = get_ws(TRANS_SHEET)

    used = ws.used_range
    last_used_row = used.last_cell.row
    last_used_col = used.last_cell.column

    # Header row is 1; last column is where header row has non-empty
    c = last_used_col
    while c >= 1 and ws.range(1, c).value in (None, ""):
        c -= 1
    if c < 6:
        # At least A..F should exist
        last_col = 6
    else:
        last_col = c

    # Find last data row based on A..F (columns 1..6)
    last_data_row = 1
    for r in range(2, last_used_row + 1):
        values = ws.range((r, 1), (r, 6)).value
        if not isinstance(values, list):
            values = [values]
        if any(v not in (None, "") for v in values):
            last_data_row = r

    first_data_row = 2
    return ws, first_data_row, last_data_row, last_col


# -------------------------------------------------
# LOAD TICKERS FROM STOCK
# -------------------------------------------------

def load_tickers_from_stock():
    """
    Return list of tickers from Stock!A2:A... (non-empty, uppercase, de-duplicated).
    """
    try:
        ws_stock = get_ws(STOCK_SHEET)
    except KeyError:
        debug("Stock sheet missing.")
        return []
    except Exception as e:
        debug(f"load_tickers_from_stock error: {e}")
        return []

    last_row = ws_stock.cells.last_cell.row
    r = last_row
    while r >= 2 and ws_stock.range(r, 1).value in (None, ""):
        r -= 1
    if r < 2:
        return []

    values = ws_stock.range(f"A2:A{r}").value
    if not isinstance(values, list):
        values = [values]

    tickers = []
    for v in values:
        if v is not None and str(v).strip() != "":
            tickers.append(str(v).strip().upper())

    seen = set()
    out = []
    for t in tickers:
        if t not in seen:
            seen.add(t)
            out.append(t)

    debug(f"load_tickers_from_stock: {len(out)} tickers loaded.")
    return out


# -------------------------------------------------
# VALIDATION
# -------------------------------------------------

def validate_date(value: str):
    v = value.strip()
    if not v:
        return False, "Date required."
    try:
        d = datetime.strptime(v, "%m/%d/%y").date()
        return True, d
    except ValueError:
        return False, "Use mm/dd/yy."


def validate_prefix(value: str):
    v = value.strip()
    if not v:
        return False, "Prefix required."
    if not v.isdigit():
        return False, "Prefix must be integer."
    i = int(v)
    if not (1 <= i <= 9):
        return False, "Prefix 1–9 only."
    return True, i


def validate_ticker(value: str, allowed):
    v = value.strip().upper()
    if not v:
        return False, "Ticker required."
    if allowed and v not in allowed:
        return False, f"Ticker '{v}' not in Stock."
    return True, v


def validate_number(value: str):
    v = value.strip()
    if not v:
        return False, "Number required."
    if not v.isdigit():
        return False, "Number must be integer."
    n = int(v)
    if n <= 0:
        return False, "> 0 required."
    return True, n


def validate_price(value: str):
    v = value.strip()
    if not v:
        return False, "Price required."
    try:
        p = float(v)
        return True, p
    except ValueError:
        return False, "Price must be numeric."


# -------------------------------------------------
# TABLE OPERATIONS VIA RANGES (xlwings)
# -------------------------------------------------

def add_transaction(excel_date, prefix_i, ticker_s, number_i, price_f):
    """
    Add a row to Transactions "table":

    - Compute last data row from A..F
    - New row = last_data_row + 1
    - Write A..F = [date, prefix, ticker, number, price, 'add']
    - Copy formulas in G..last_col from previous data row (if any)
    """
    ws, first_data_row, last_data_row, last_col = get_transactions_table_dims()

    if last_data_row < first_data_row:
        # No data rows yet -> new row is 2
        new_row = first_data_row
        prev_row = None
    else:
        new_row = last_data_row + 1
        prev_row = last_data_row

    # 1) Write data columns A..F
    row_values = [excel_date, int(prefix_i), ticker_s, int(number_i), float(price_f), "add"]
    ws.range((new_row, 1), (new_row, 6)).value = row_values

    # 2) Copy formulas in G..last_col from previous row, if there is one
    if prev_row is not None and last_col > 6:
        src = ws.range((prev_row, 7), (prev_row, last_col))
        dst = ws.range((new_row, 7), (new_row, last_col))
        try:
            # Copy formulas if they exist; fall back to values
            formulas = src.formula
            dst.formula = formulas
        except Exception:
            dst.value = src.value

    excel_save()
    debug(f"add_transaction: row={new_row}, values={row_values}, last_col={last_col}")
    append_log_entry("add", f"Row {new_row}: {row_values}")


def delete_transaction(excel_date, prefix_i, ticker_s, number_i, price_f):
    """
    Delete first row in Transactions where A..E match:
       (Date, Prefix, Ticker, Number, Price)

    Deletes the entire sheet row so formulas and other columns shift up.
    """
    ws, first_data_row, last_data_row, last_col = get_transactions_table_dims()

    if last_data_row < first_data_row:
        messagebox.showwarning("Delete", "No transactions to delete.")
        return False

    target_row = None
    for r in range(first_data_row, last_data_row + 1):
        vals = ws.range((r, 1), (r, 5)).value  # Date, Prefix, Ticker, Number, Price
        if not isinstance(vals, list):
            vals = [vals]

        date_val = vals[0]
        pref_val = vals[1] if len(vals) > 1 else None
        tick_val = vals[2] if len(vals) > 2 else None
        num_val = vals[3] if len(vals) > 3 else None
        price_val = vals[4] if len(vals) > 4 else None

        # Normalize date
        if isinstance(date_val, datetime):
            date_val = date_val.date()

        tick_norm = (tick_val or "").strip().upper()
        try:
            price_norm = float(price_val)
        except Exception:
            price_norm = price_val

        if (date_val == excel_date and
            pref_val == prefix_i and
            tick_norm == ticker_s and
            num_val == number_i and
            price_norm == price_f):
            target_row = r
            break

    if target_row is None:
        msg = f"No match for ({excel_date}, {prefix_i}, {ticker_s}, {number_i}, {price_f})"
        debug(f"delete_transaction: {msg}")
        append_log_entry("delete_fail", msg)
        messagebox.showwarning("Delete", "No matching transaction found.")
        return False

    debug(f"delete_transaction: deleting sheet row {target_row}")
    ws.range(target_row, 1).api.EntireRow.Delete()
    excel_save()
    append_log_entry(
        "delete",
        f"Deleted row {target_row} for ({excel_date}, {prefix_i}, {ticker_s}, {number_i}, {price_f})"
    )
    messagebox.showinfo("Delete", "Transaction deleted.")
    return True


# -------------------------------------------------
# GUI
# -------------------------------------------------

def build_gui():
    root = tk.Tk()
    root.title("TRADES – Excel-driven Transactions Manager")

    frm = ttk.Frame(root, padding=10)
    frm.grid(row=0, column=0, sticky="nsew")

    # Load tickers once at start (button can reload them)
    ticker_list = load_tickers_from_stock()

    # Variables
    date_var = tk.StringVar(value=date.today().strftime("%m/%d/%y"))
    prefix_var = tk.StringVar(value="1")
    ticker_var = tk.StringVar()
    number_var = tk.StringVar()
    price_var = tk.StringVar()

    date_err = tk.StringVar(value="")
    prefix_err = tk.StringVar(value="")
    ticker_err = tk.StringVar(value="")
    number_err = tk.StringVar(value="")
    price_err = tk.StringVar(value="")

    # --- Layout helpers
    def err_label(row):
        lbl = tk.Label(frm, text="", fg="red", anchor="w")
        lbl.grid(row=row, column=1, columnspan=2, sticky="w")
        return lbl

    # Date
    ttk.Label(frm, text="Date (mm/dd/yy):").grid(row=0, column=0, sticky="w")
    date_entry = ttk.Entry(frm, textvariable=date_var, width=15)
    date_entry.grid(row=0, column=1, sticky="w")
    date_err_lbl = err_label(1)

    # Prefix
    ttk.Label(frm, text="Prefix (1–9):").grid(row=2, column=0, sticky="w")
    prefix_cb = ttk.Combobox(
        frm, textvariable=prefix_var,
        values=[str(i) for i in range(1, 10)],
        width=5, state="readonly"
    )
    prefix_cb.grid(row=2, column=1, sticky="w")
    prefix_err_lbl = err_label(3)

    # Ticker
    ttk.Label(frm, text="Ticker:").grid(row=4, column=0, sticky="w")
    ticker_cb = ttk.Combobox(
        frm, textvariable=ticker_var,
        values=ticker_list,
        width=15, state="readonly"
    )
    ticker_cb.grid(row=4, column=1, sticky="w")
    ticker_err_lbl = err_label(5)

    # Number
    ttk.Label(frm, text="Number:").grid(row=6, column=0, sticky="w")
    number_entry = ttk.Entry(frm, textvariable=number_var, width=15)
    number_entry.grid(row=6, column=1, sticky="w")
    number_err_lbl = err_label(7)

    # Price
    ttk.Label(frm, text="Price:").grid(row=8, column=0, sticky="w")
    price_entry = ttk.Entry(frm, textvariable=price_var, width=15)
    price_entry.grid(row=8, column=1, sticky="w")
    price_err_lbl = err_label(9)

    # Validation on focus out
    def on_date_out(_=None):
        ok, res = validate_date(date_var.get())
        date_err.set("" if ok else res)

    def on_prefix_out(_=None):
        ok, res = validate_prefix(prefix_var.get())
        prefix_err.set("" if ok else res)

    def on_ticker_out(_=None):
        ok, res = validate_ticker(ticker_var.get(), ticker_list)
        ticker_err.set("" if ok else res)

    def on_number_out(_=None):
        ok, res = validate_number(number_var.get())
        number_err.set("" if ok else res)

    def on_price_out(_=None):
        ok, res = validate_price(price_var.get())
        price_err.set("" if ok else res)

    date_entry.bind("<FocusOut>", on_date_out)
    prefix_cb.bind("<FocusOut>", on_prefix_out)
    ticker_cb.bind("<FocusOut>", on_ticker_out)
    number_entry.bind("<FocusOut>", on_number_out)
    price_entry.bind("<FocusOut>", on_price_out)

    # Bind error vars to labels
    date_err.trace_add("write", lambda *_: date_err_lbl.config(text=date_err.get()))
    prefix_err.trace_add("write", lambda *_: prefix_err_lbl.config(text=prefix_err.get()))
    ticker_err.trace_add("write", lambda *_: ticker_err_lbl.config(text=ticker_err.get()))
    number_err.trace_add("write", lambda *_: number_err_lbl.config(text=number_err.get()))
    price_err.trace_add("write", lambda *_: price_err_lbl.config(text=price_err.get()))

    def clear_errors():
        date_err.set("")
        prefix_err.set("")
        ticker_err.set("")
        number_err.set("")
        price_err.set("")

    def clear_fields():
        date_var.set(date.today().strftime("%m/%d/%y"))
        prefix_var.set("1")
        number_var.set("")
        price_var.set("")

    # --- Actions

    def reload_tickers():
        nonlocal ticker_list
        ticker_list = load_tickers_from_stock()
        ticker_cb["values"] = ticker_list
        if ticker_list:
            ticker_var.set(ticker_list[0])
            ticker_err.set("")
        else:
            ticker_var.set("")
            ticker_err.set("No tickers in Stock.")

    ttk.Button(frm, text="Reload tickers", command=reload_tickers)\
        .grid(row=4, column=2, padx=5)

    def do_add():
        clear_errors()
        ok_d, res_d = validate_date(date_var.get())
        ok_pfx, res_pfx = validate_prefix(prefix_var.get())
        ok_t, res_t = validate_ticker(ticker_var.get(), ticker_list)
        ok_n, res_n = validate_number(number_var.get())
        ok_pr, res_pr = validate_price(price_var.get())

        if not ok_d:
            date_err.set(res_d)
        if not ok_pfx:
            prefix_err.set(res_pfx)
        if not ok_t:
            ticker_err.set(res_t)
        if not ok_n:
            number_err.set(res_n)
        if not ok_pr:
            price_err.set(res_pr)

        if not (ok_d and ok_pfx and ok_t and ok_n and ok_pr):
            messagebox.showerror("Validation", "Please correct highlighted fields.")
            return

        excel_date, prefix_i, ticker_s, number_i, price_f = \
            res_d, res_pfx, res_t, res_n, res_pr

        # Confirm purchase/sale
        if price_f < 0:
            if not messagebox.askyesno("Confirm", f"Price={price_f}. Confirm PURCHASE?"):
                return
        elif price_f > 0:
            if not messagebox.askyesno("Confirm", f"Price={price_f}. Confirm SALE?"):
                return

        try:
            add_transaction(excel_date, prefix_i, ticker_s, number_i, price_f)
            messagebox.showinfo("Add", "Transaction added.")
            clear_fields()
        except Exception as e:
            debug(f"do_add: error: {e}")
            messagebox.showerror("Error", f"Add failed:\n{e}")

    def do_delete():
        clear_errors()
        ok_d, res_d = validate_date(date_var.get())
        ok_pfx, res_pfx = validate_prefix(prefix_var.get())
        ok_t, res_t = validate_ticker(ticker_var.get(), ticker_list)
        ok_n, res_n = validate_number(number_var.get())
        ok_pr, res_pr = validate_price(price_var.get())

        if not ok_d:
            date_err.set(res_d)
        if not ok_pfx:
            prefix_err.set(res_pfx)
        if not ok_t:
            ticker_err.set(res_t)
        if not ok_n:
            number_err.set(res_n)
        if not ok_pr:
            price_err.set(res_pr)

        if not (ok_d and ok_pfx and ok_t and ok_n and ok_pr):
            messagebox.showerror("Validation", "Please correct highlighted fields.")
            return

        excel_date, prefix_i, ticker_s, number_i, price_f = \
            res_d, res_pfx, res_t, res_n, res_pr

        try:
            delete_transaction(excel_date, prefix_i, ticker_s, number_i, price_f)
            clear_fields()
        except Exception as e:
            debug(f"do_delete: error: {e}")
            messagebox.showerror("Error", f"Delete failed:\n{e}")

    def do_open_wb():
        append_log_entry("open", "Workbook opened from GUI.")
        try:
            _, wb = ensure_workbook()
            wb.app.visible = True
            wb.activate()
        except Exception as e:
            debug(f"do_open_wb: {e}")
            try:
                if os.name == "posix":
                    subprocess.call(["open", EXCEL_FILE])
                else:
                    os.startfile(EXCEL_FILE)
            except Exception as e2:
                messagebox.showerror("Error", f"Cannot open workbook:\n{e2}")

    def do_exit():
        append_log_entry("exit", "GUI session closed.")
        excel_quit()
        root.destroy()

    # Button bar
    btn = ttk.Frame(frm)
    btn.grid(row=10, column=0, columnspan=3, pady=10, sticky="w")

    ttk.Button(btn, text="Add",    command=do_add).grid(row=0, column=0, padx=5)
    ttk.Button(btn, text="Delete", command=do_delete).grid(row=0, column=1, padx=5)
    ttk.Button(btn, text="Open WB",command=do_open_wb).grid(row=0, column=2, padx=5)
    ttk.Button(btn, text="Exit",   command=do_exit).grid(row=0, column=3, padx=5)

    return root


# -------------------------------------------------
# MAIN
# -------------------------------------------------

def main():
    debug(f"Using TRADES.xlsx at: {EXCEL_FILE}")
    try:
        ensure_workbook()
        # simple sanity: transactions sheet exists and headers A1..F1 non-empty
        ws = get_ws(TRANS_SHEET)
        headers = ws.range((1, 1), (1, 6)).value
        if not isinstance(headers, list):
            headers = [headers]
        if any(h in (None, "") for h in headers):
            debug(f"Warning: some headers empty in A1:F1 -> {headers}")
    except Exception as e:
        debug(f"Startup error: {e}")
        messagebox.showerror("Startup error", str(e))
        excel_quit()
        return

    root = build_gui()
    root.mainloop()
    excel_quit()


if __name__ == "__main__":
    main()
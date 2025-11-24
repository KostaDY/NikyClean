#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import csv
from datetime import datetime, date
import openpyxl
import tkinter as tk
from tkinter import ttk, messagebox
import subprocess

# ======================================
# CONFIGURATION
# ======================================

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, "TRADES.xlsx")
LOG_CSV_FILE = os.path.join(BASE_DIR, "TRADES_log.csv")

TRANS_HEADERS = ["Date", "Prefix", "Ticker", "Number", "Price", "Act"]
LOG_HEADERS = ["Timestamp", "Action", "Details"]

TICKER_LIST = []
TICKER_DISPLAY = []
last_deleted_entry = None

wb = None
ws_trans = None
ws_log = None


# ======================================
# INITIALIZATION
# ======================================

def init_log_csv():
    if not os.path.exists(LOG_CSV_FILE):
        with open(LOG_CSV_FILE, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(LOG_HEADERS)


def init_workbook():
    global wb, ws_trans, ws_log

    if not os.path.exists(EXCEL_FILE):
        wb2 = openpyxl.Workbook()

        ws_t = wb2.active
        ws_t.title = "Transactions"
        for col, h in enumerate(TRANS_HEADERS, start=1):
            ws_t.cell(row=1, column=col).value = h

        ws_l = wb2.create_sheet("Log")
        for col, h in enumerate(LOG_HEADERS, start=1):
            ws_l.cell(row=1, column=col).value = h

        ws_s = wb2.create_sheet("Stock")
        ws_s.cell(1, 1, "Ticker")
        ws_s.cell(1, 2, "Description")

        wb2.save(EXCEL_FILE)

    wb = openpyxl.load_workbook(EXCEL_FILE)

    ws_trans = wb["Transactions"]
    ws_log = wb["Log"]

    wb.save(EXCEL_FILE)


def rebuild_log_from_csv():
    global ws_log, wb

    while ws_log.max_row > 1:
        ws_log.delete_rows(2)

    if os.path.exists(LOG_CSV_FILE):
        with open(LOG_CSV_FILE, "r", encoding="utf-8") as f:
            rdr = csv.reader(f)
            next(rdr, None)
            for row in rdr:
                nr = ws_log.max_row + 1
                for c, v in enumerate(row, start=1):
                    ws_log.cell(nr, c).value = v

    wb.save(EXCEL_FILE)


# ======================================
# LOGGING
# ======================================

def append_log(action, details):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    row = [ts, action, details]

    with open(LOG_CSV_FILE, "a", newline="", encoding="utf-8") as f:
        csv.writer(f).writerow(row)

    nr = ws_log.max_row + 1
    for c, v in enumerate(row, start=1):
        ws_log.cell(nr, c).value = v

    wb.save(EXCEL_FILE)


# ======================================
# LOAD TICKERS FROM STOCK SHEET
# ======================================

def load_tickers_from_workbook():
    tickers = []
    display = []

    try:
        wb2 = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    except:
        return [], []

    if "Stock" not in wb2.sheetnames:
        return [], []

    ws_s = wb2["Stock"]

    # CASE 1 — Named range exists
    if "Tickers" in wb2.defined_names:
        dn = wb2.defined_names["Tickers"]
        for sheetname, cell_ref in dn.destinations:
            ws2 = wb2[sheetname]
            rng = ws2[cell_ref]
            for row in rng:
                cell = row[0]
                if cell.value:
                    t = str(cell.value).strip().upper()
                    desc = ws2.cell(row=cell.row, column=cell.col_idx+1).value
                    desc = str(desc).strip() if desc else ""
                    tickers.append(t)
                    display.append(f"{t} – {desc}" if desc else t)
    else:
        # CASE 2 — Column A+B
        for r in ws_s.iter_rows(min_row=2, max_col=2):
            if r[0].value:
                t = str(r[0].value).strip().upper()
                desc = r[1].value
                desc = str(desc).strip() if desc else ""
                tickers.append(t)
                display.append(f"{t} – {desc}" if desc else t)

    final = list(dict.fromkeys(zip(tickers, display)))
    return [f[0] for f in final], [f[1] for f in final]


# ======================================
# VALIDATION
# ======================================

def validate_entry(datestr, prefixstr, ticker, numberstr, pricestr, act):
    # date
    datetime.strptime(datestr, "%m/%d/%y")

    # prefix
    if not prefixstr.isdigit():
        raise ValueError("Prefix must be integer.")
    px = int(prefixstr)
    if not (1 <= px <= 9):
        raise ValueError("Prefix must be between 1 and 9.")

    # ticker
    tick = ticker.strip().upper()
    if TICKER_LIST and tick not in TICKER_LIST:
        raise ValueError(f"Ticker must be one of: {', '.join(TICKER_LIST)}")

    # number
    if not numberstr.isdigit():
        raise ValueError("Number must be positive integer.")
    num = int(numberstr)
    if num <= 0:
        raise ValueError("Number must be > 0.")

    # price
    try:
        price = float(pricestr)
    except:
        raise ValueError("Price must be numeric.")

    act_clean = act.lower().strip() or "add"
    if act_clean not in ["add", "delete", "open", "exit"]:
        raise ValueError("Act must be add/delete/open/exit.")

    # Confirmation ONLY for ADD
    if act_clean == "add":
        if price < 0:
            if not messagebox.askyesno("Confirm Purchase",
                                       f"Price = {price}. Confirm PURCHASE?"):
                raise ValueError("Cancelled.")
        elif price > 0:
            if not messagebox.askyesno("Confirm Sale",
                                       f"Price = {price}. Confirm SALE?"):
                raise ValueError("Cancelled.")

    excel_date = datetime.strptime(datestr, "%m/%d/%y").date()
    return excel_date, px, tick, num, price, act_clean


# ======================================
# APPEND TRANSACTION
# ======================================

def append_transaction(entry):
    nr = ws_trans.max_row + 1
    for c, v in enumerate(entry, start=1):
        ws_trans.cell(nr, c).value = v
    wb.save(EXCEL_FILE)
    append_log("add", f"Added: {entry}")


# ======================================
# SAFE DELETE
# ======================================

def delete_matching_transaction(entry):
    global last_deleted_entry

    date_s, prefix_i, ticker_s, number_i, price_f, _ = entry
    max_row = ws_trans.max_row

    target = None
    for r in range(2, max_row + 1):
        vals = [
            ws_trans.cell(r, 1).value,
            ws_trans.cell(r, 2).value,
            ws_trans.cell(r, 3).value,
            ws_trans.cell(r, 4).value,
            ws_trans.cell(r, 5).value
        ]
        if vals == [date_s, prefix_i, ticker_s, number_i, price_f]:
            target = r
            break

    if not target:
        messagebox.showwarning("Not found", "No matching transaction found.")
        append_log("delete_fail", f"No match for {entry}")
        return False

    saved_row = [ws_trans.cell(target, c).value for c in range(1, 7)]
    last_deleted_entry = (saved_row, target)

    last_used = max_row
    while last_used > 1:
        if any(ws_trans.cell(last_used, c).value not in ("", None)
               for c in range(1, 7)):
            break
        last_used -= 1

    for r in range(target, last_used):
        for c in range(1, 7):
            ws_trans.cell(r, c).value = ws_trans.cell(r + 1, c).value

    for c in range(1, 7):
        ws_trans.cell(last_used, c).value = None

    wb.save(EXCEL_FILE)
    append_log("delete", f"Deleted: {saved_row}")
    messagebox.showinfo("Deleted", "Transaction deleted.")
    return True


# ======================================
# UNDO DELETE
# ======================================

def undo_last_delete():
    global last_deleted_entry

    if not last_deleted_entry:
        messagebox.showinfo("Undo", "Nothing to undo.")
        return

    values, _ = last_deleted_entry
    append_transaction(values)
    append_log("undo_delete", f"Restored: {values}")
    last_deleted_entry = None
    messagebox.showinfo("Undo", "Last deleted transaction restored.")


# ======================================
# OPEN WORKBOOK
# ======================================

def open_workbook():
    append_log("open", "Opened workbook.")
    if os.name == "posix":
        subprocess.call(["open", EXCEL_FILE])
    else:
        os.startfile(EXCEL_FILE)


# ======================================
# GUI
# ======================================

def build_gui():
    global TICKER_LIST, TICKER_DISPLAY

    root = tk.Tk()
    root.title("TRADES – Transactions Manager")

    frm = ttk.Frame(root, padding=10)
    frm.grid(row=0, column=0, sticky="nsew")

    date_var = tk.StringVar(value=date.today().strftime("%m/%d/%y"))
    prefix_var = tk.StringVar(value="1")
    ticker_var = tk.StringVar()
    number_var = tk.StringVar()
    price_var = tk.StringVar()
    act_var = tk.StringVar(value="add")

    TICKER_LIST, TICKER_DISPLAY = load_tickers_from_workbook()
    if TICKER_LIST:
        ticker_var.set(TICKER_LIST[0])

    # --- Date
    ttk.Label(frm, text="Date (mm/dd/yy):").grid(row=0, column=0, sticky="w")
    ttk.Entry(frm, textvariable=date_var, width=15).grid(row=0, column=1)

    # --- Prefix
    ttk.Label(frm, text="Prefix (1–9):").grid(row=1, column=0, sticky="w")
    ttk.Combobox(frm, textvariable=prefix_var,
                 values=[str(i) for i in range(1, 10)],
                 width=5, state="readonly").grid(row=1, column=1)

    # --- Ticker
    ttk.Label(frm, text="Ticker:").grid(row=2, column=0, sticky="w")

    ticker_cb = ttk.Combobox(
        frm, textvariable=ticker_var, width=30,
        values=TICKER_DISPLAY, state="readonly"
    )
    ticker_cb.grid(row=2, column=1)

    def on_ticker_select(e):
        val = ticker_var.get()
        ticker_var.set(val.split("–")[0].strip())

    ticker_cb.bind("<<ComboboxSelected>>", on_ticker_select)

    def reload_tickers():
        global TICKER_LIST, TICKER_DISPLAY
        TICKER_LIST, TICKER_DISPLAY = load_tickers_from_workbook()
        ticker_cb["values"] = TICKER_DISPLAY
        if TICKER_LIST:
            ticker_var.set(TICKER_LIST[0])

    ttk.Button(frm, text="Reload Tickers", command=reload_tickers)\
        .grid(row=2, column=2, padx=5)

    # --- Number
    ttk.Label(frm, text="Number:").grid(row=3, column=0, sticky="w")
    ttk.Entry(frm, textvariable=number_var, width=15).grid(row=3, column=1)

    # --- Price
    ttk.Label(frm, text="Price:").grid(row=4, column=0, sticky="w")
    ttk.Entry(frm, textvariable=price_var, width=15).grid(row=4, column=1)

    # --- Act
    ttk.Label(frm, text="Act:").grid(row=5, column=0, sticky="w")
    ttk.Combobox(frm, textvariable=act_var,
                 values=["add", "delete", "open", "exit"],
                 width=10, state="readonly").grid(row=5, column=1)

    # Clear input (after add/delete)
    def clear_fields():
        date_var.set(date.today().strftime("%m/%d/%y"))
        prefix_var.set("1")
        number_var.set("")
        price_var.set("")
        act_var.set("add")

    # Run main action
    def run_action():
        try:
            entry = validate_entry(
                date_var.get(),
                prefix_var.get(),
                ticker_var.get(),
                number_var.get(),
                price_var.get(),
                act_var.get()
            )
        except ValueError as e:
            messagebox.showerror("Validation", str(e))
            return

        act = entry[-1]

        if act == "open":
            open_workbook()
            return

        if act == "exit":
            append_log("exit", "GUI session closed.")
            root.destroy()
            return

        if act == "delete":
            if delete_matching_transaction(entry):
                clear_fields()
            return

        if act == "add":
            append_transaction(entry)
            messagebox.showinfo("Added", "Transaction added.")
            clear_fields()
            return

    # ============ BUTTON BAR ============

    btn = ttk.Frame(frm)
    btn.grid(row=6, column=0, columnspan=3, pady=10)

    ttk.Button(btn, text="Add",
               command=lambda: [act_var.set("add"), run_action()]
               ).grid(row=0, column=0, padx=5)

    ttk.Button(btn, text="Delete",
               command=lambda: [act_var.set("delete"), run_action()]
               ).grid(row=0, column=1, padx=5)

    ttk.Button(btn, text="Open WB",
               command=open_workbook
               ).grid(row=0, column=2, padx=5)

    ttk.Button(btn, text="Undo Delete",
               command=undo_last_delete
               ).grid(row=0, column=3, padx=5)

    ttk.Button(btn, text="Exit",
               command=lambda: [
                   append_log("exit", "GUI session closed."),
                   root.destroy()
               ]
               ).grid(row=0, column=4, padx=5)

    return root


# ======================================
# MAIN
# ======================================

def main():
    init_log_csv()
    init_workbook()
    rebuild_log_from_csv()
    append_log("start", "GUI session started.")

    root = build_gui()
    root.mainloop()
    wb.save(EXCEL_FILE)


if __name__ == "__main__":
    main()
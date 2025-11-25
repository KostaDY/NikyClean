#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import tkinter as tk
from tkinter import ttk, messagebox
from openpyxl import load_workbook

from tableloader import safe_update_table, atomic_save, get_table_info, excel_bounds
from csv_exporter import export_tabtrans_to_csv

# ============================================================
# CONFIGURATION
# ============================================================

BASE_DIR      = "/Users/kostayanev/NikyClean"
TRADES_XLSX   = os.path.join(BASE_DIR, "TRADES.xlsx")
CSV_PATH      = os.path.join(BASE_DIR, "Transactions.csv")

SHEET_NAME    = "Transactions"
TABLE_NAME    = "TabTrans"

# Data-entry columns (only these are imported from CSV)
DATA_COLUMNS  = ["Date", "Prefix", "Ticker", "Number", "Price"]


# ============================================================
# EXCEL HELPERS
# ============================================================

def open_workbook(data_only=False):
    """Open TRADES.xlsx."""
    return load_workbook(TRADES_XLSX, data_only=data_only)


def load_table_data():
    """
    Load TabTrans values for the Treeview.
    Returns:
        col_names : list of header names
        rows      : list of (excel_row, [values...])
        header_row, min_col, max_col, first_data_row
    """
    wb = open_workbook(data_only=True)
    ws = wb[SHEET_NAME]
    header_row, min_col, max_col, col_names = get_table_info(ws, TABLE_NAME)
    (min_row, _), (max_row, _) = excel_bounds(ws.tables[TABLE_NAME].ref)

    rows = []
    for r in range(header_row + 1, max_row + 1):
        row_vals = [ws.cell(row=r, column=c).value for c in range(min_col, max_col + 1)]
        rows.append((r, row_vals))  # include Excel row number
    return col_names, rows, header_row, min_col, max_col, header_row + 1


def append_row_to_table(values_dict):
    """
    Append a new row to TabTrans, updating only DATA_COLUMNS.
    values_dict: {"Date": ..., "Prefix": ..., ...}
    """
    wb = open_workbook(data_only=False)
    ws = wb[SHEET_NAME]
    table = ws.tables[TABLE_NAME]

    (min_row, min_col), (max_row, max_col) = excel_bounds(table.ref)
    header_row = min_row
    next_row = max_row + 1

    # map headers to offsets
    _, _, _, col_names = get_table_info(ws, TABLE_NAME)
    col_index = {name: i for i, name in enumerate(col_names)}

    # write data columns
    for col_name in DATA_COLUMNS:
        excel_col = min_col + col_index[col_name]
        ws.cell(row=next_row, column=excel_col).value = values_dict.get(col_name)

    # resize table to include new row
    from openpyxl.utils import get_column_letter
    start_cell = f"{get_column_letter(min_col)}{header_row}"
    end_cell   = f"{get_column_letter(max_col)}{next_row}"
    table.ref = f"{start_cell}:{end_cell}"

    wb.save(TRADES_XLSX)


def update_row_in_table(excel_row, values_dict):
    """
    Update an existing row (excel_row) in TabTrans for DATA_COLUMNS only.
    """
    wb = open_workbook(data_only=False)
    ws = wb[SHEET_NAME]
    table = ws.tables[TABLE_NAME]

    (min_row, min_col), (max_row, max_col) = excel_bounds(table.ref)
    _, _, _, col_names = get_table_info(ws, TABLE_NAME)
    col_index = {name: i for i, name in enumerate(col_names)}

    if excel_row <= min_row or excel_row > max_row:
        raise ValueError("Row out of table bounds")

    for col_name in DATA_COLUMNS:
        excel_col = min_col + col_index[col_name]
        ws.cell(row=excel_row, column=excel_col).value = values_dict.get(col_name)

    wb.save(TRADES_XLSX)


def delete_row_from_table(excel_row):
    """
    Delete a row from TabTrans and shrink the table.
    """
    wb = open_workbook(data_only=False)
    ws = wb[SHEET_NAME]
    table = ws.tables[TABLE_NAME]

    (min_row, min_col), (max_row, max_col) = excel_bounds(table.ref)
    if excel_row <= min_row or excel_row > max_row:
        raise ValueError("Row out of table bounds")

    ws.delete_rows(excel_row, 1)

    # shrink table range
    from openpyxl.utils import get_column_letter
    new_max_row = max_row - 1
    start_cell = f"{get_column_letter(min_col)}{min_row}"
    end_cell   = f"{get_column_letter(max_col)}{new_max_row}"
    table.ref = f"{start_cell}:{end_cell}"

    wb.save(TRADES_XLSX)


# ============================================================
# GUI CLASS
# ============================================================

class TradesGUI:
    def __init__(self, root):
        self.root = root
        root.title("Trades Entry")

        self.selected_excel_row = None   # Excel row index for current selection
        self.col_names = []              # Table headers
        self.row_map = {}                # tree item id -> excel_row

        main = ttk.Frame(root, padding=10)
        main.grid(sticky="nsew")
        root.rowconfigure(0, weight=1)
        root.columnconfigure(0, weight=1)

        # ---------- Treeview with TabTrans ----------
        tree_frame = ttk.Frame(main)
        tree_frame.grid(row=0, column=0, columnspan=5, sticky="nsew")

        self.tree = ttk.Treeview(tree_frame, show="headings")
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        self.tree.bind("<Double-1>", self.on_open_clicked)  # C: double-click opens

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        # ---------- Buttons ----------
        btn_frame = ttk.Frame(main, padding=(0, 5))
        btn_frame.grid(row=1, column=0, columnspan=5, sticky="w")

        ttk.Button(btn_frame, text="Open",   command=self.on_open_clicked).grid(row=0, column=0, padx=2)
        ttk.Button(btn_frame, text="Add New",command=self.on_add_new_clicked).grid(row=0, column=1, padx=2)
        ttk.Button(btn_frame, text="Delete", command=self.on_delete_clicked).grid(row=0, column=2, padx=2)
        ttk.Button(btn_frame, text="Refresh",command=self.refresh_tree).grid(row=0, column=3, padx=2)
        ttk.Button(btn_frame, text="Exit",   command=self.on_exit_clicked).grid(row=0, column=4, padx=2)

        # ---------- Entry / Edit form ----------
        form = ttk.LabelFrame(main, text="Entry / Edit", padding=10)
        form.grid(row=2, column=0, columnspan=5, sticky="ew", pady=(5, 0))

        ttk.Label(form, text="Date:").grid(  row=0, column=0, sticky="e")
        ttk.Label(form, text="Prefix:").grid(row=1, column=0, sticky="e")
        ttk.Label(form, text="Ticker:").grid(row=2, column=0, sticky="e")
        ttk.Label(form, text="Number:").grid(row=3, column=0, sticky="e")
        ttk.Label(form, text="Price:").grid( row=4, column=0, sticky="e")

        self.e_date   = ttk.Entry(form, width=12)
        self.e_prefix = ttk.Entry(form, width=6)
        self.e_ticker = ttk.Entry(form, width=10)
        self.e_number = ttk.Entry(form, width=10)
        self.e_price  = ttk.Entry(form, width=10)

        self.e_date.grid(  row=0, column=1, padx=2, pady=2)
        self.e_prefix.grid(row=1, column=1, padx=2, pady=2)
        self.e_ticker.grid(row=2, column=1, padx=2, pady=2)
        self.e_number.grid(row=3, column=1, padx=2, pady=2)
        self.e_price.grid( row=4, column=1, padx=2, pady=2)

        ttk.Button(form, text="Save (Add / Update)", command=self.on_save_clicked).grid(
            row=5, column=0, columnspan=2, pady=(5, 0)
        )

        main.rowconfigure(0, weight=1)

        # Initial load of table
        self.refresh_tree()

    # --------------------------------------------------------
    # Treeview helpers
    # --------------------------------------------------------
    def refresh_tree(self):
        """Reload the Treeview from Excel."""
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.row_map.clear()
        self.selected_excel_row = None

        try:
            col_names, rows, header_row, min_col, max_col, first_data_row = load_table_data()
        except Exception as e:
            messagebox.showerror("Error", f"Cannot load Excel table:\n{e}")
            return

        self.col_names = col_names
        self.tree["columns"] = col_names
        for name in col_names:
            self.tree.heading(name, text=name)
            self.tree.column(name, width=90, anchor="center")

        for excel_row, vals in rows:
            item_id = self.tree.insert("", "end", values=vals)
            self.row_map[item_id] = excel_row

    def on_tree_select(self, event=None):
        sel = self.tree.selection()
        if not sel:
            self.selected_excel_row = None
            return
        item_id = sel[0]
        self.selected_excel_row = self.row_map.get(item_id)

    # --------------------------------------------------------
    # Button handlers
    # --------------------------------------------------------
    def on_open_clicked(self, event=None):
        """Open selected row into the entry form."""
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Info", "Please select a row in the table.")
            return
        item_id = sel[0]
        excel_row = self.row_map.get(item_id)
        if excel_row is None:
            return

        vals = self.tree.item(item_id, "values")
        col_idx = {name: i for i, name in enumerate(self.col_names)}

        def get_val(col):
            idx = col_idx.get(col)
            return vals[idx] if idx is not None and idx < len(vals) else ""

        # populate form
        self.e_date.delete(0, tk.END)
        self.e_prefix.delete(0, tk.END)
        self.e_ticker.delete(0, tk.END)
        self.e_number.delete(0, tk.END)
        self.e_price.delete(0, tk.END)

        self.e_date.insert(0,   get_val("Date"))
        self.e_prefix.insert(0, get_val("Prefix"))
        self.e_ticker.insert(0, get_val("Ticker"))
        self.e_number.insert(0, get_val("Number"))
        self.e_price.insert(0,  get_val("Price"))

        self.selected_excel_row = excel_row

    def on_add_new_clicked(self):
        """Prepare form for a new row."""
        self.selected_excel_row = None
        self.e_date.delete(0, tk.END)
        self.e_prefix.delete(0, tk.END)
        self.e_ticker.delete(0, tk.END)
        self.e_number.delete(0, tk.END)
        self.e_price.delete(0, tk.END)

    def on_delete_clicked(self):
        """Delete selected row (with confirmation)."""
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Info", "Please select a row to delete.")
            return
        item_id = sel[0]
        excel_row = self.row_map.get(item_id)
        if excel_row is None:
            return

        ans = messagebox.askyesno("Confirm Delete", "Delete selected transaction?")
        if not ans:
            return

        try:
            delete_row_from_table(excel_row)
            self.refresh_tree()
        except Exception as e:
            messagebox.showerror("Error", f"Could not delete row:\n{e}")

    def on_save_clicked(self):
        """Add or update a transaction (with basic validation)."""
        date   = self.e_date.get().strip()
        prefix = self.e_prefix.get().strip()
        ticker = self.e_ticker.get().strip()
        number = self.e_number.get().strip()
        price  = self.e_price.get().strip()

        # basic validation
        if not ticker:
            messagebox.showerror("Validation", "Ticker is required.")
            return
        if not date:
            messagebox.showerror("Validation", "Date is required.")
            return

        try:
            num_val = float(number)
        except Exception:
            messagebox.showerror("Validation", "Number must be numeric.")
            return
        try:
            price_val = float(price)
        except Exception:
            messagebox.showerror("Validation", "Price must be numeric.")
            return

        values = {
            "Date":   date,
            "Prefix": prefix,
            "Ticker": ticker,
            "Number": num_val,
            "Price":  price_val,
        }

        try:
            if self.selected_excel_row is None:
                append_row_to_table(values)
            else:
                update_row_in_table(self.selected_excel_row, values)
            self.refresh_tree()
            messagebox.showinfo("OK", "Transaction saved.")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save transaction:\n{e}")

    def on_exit_clicked(self):
        """Export entire TabTrans to CSV, then close GUI."""
        try:
            export_tabtrans_to_csv()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export CSV:\n{e}")
        self.root.destroy()


# ============================================================
# INITIALIZE FROM CSV, THEN RUN GUI
# ============================================================

def initialize_tabtrans_from_csv():
    """
    At GUI start:
    - If Transactions.csv exists: import its data columns into TabTrans.
    - Formula columns remain untouched.
    """
    if not os.path.exists(CSV_PATH):
        print("⚠ Transactions.csv does not exist — starting from current Excel.")
        return
    try:
        wb = safe_update_table(CSV_PATH, TRADES_XLSX, SHEET_NAME, TABLE_NAME, DATA_COLUMNS)
        atomic_save(wb, TRADES_XLSX)
        print("✔ TRADES.xlsx updated from CSV (data columns only).")
    except Exception as e:
        print("✖ Failed to update Excel from CSV:", e)


def main():
    initialize_tabtrans_from_csv()
    root = tk.Tk()
    TradesGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
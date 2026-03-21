import pandas as pd
from pathlib import Path

INPUT_XLSX  = "trades.xlsx"
INPUT_SHEET = "Trades"
OUTPUT_XLSX = "PriceFIFO_Result.xlsx"

df = pd.read_excel(INPUT_XLSX, sheet_name=INPUT_SHEET)
df = df.iloc[:, :3]
df.columns = ["Date", "Number", "Price"]

df["Date"]   = pd.to_datetime(df["Date"], errors="coerce")
df["Number"] = df["Number"].astype(int)
df["Price"]  = df["Price"].astype(float)

# purchases: negative price; sales: positive price
purchases = df[df["Price"] < 0].copy()
sales     = df[df["Price"] > 0].copy()

# helper: positive purchase price for matching
purchases["_P"] = -purchases["Price"]

# deterministic sorting (FIFO within equal matching key)
purchases.sort_values(["_P", "Date"], inplace=True)         # ascending purchase price
sales.sort_values(["Price", "Date"], inplace=True)          # ascending sale price

# tracking
purchases["_RemainingQty"] = purchases["Number"]
purchases["_Trace_OK"]     = ""
purchases["_Trace_VIOL"]   = ""

sales["_UnsatisfiedQty"]   = 0
sales["_Matched_OK_Qty"]   = 0
sales["_Matched_VIOL_Qty"] = 0
sales["_Matched_OK"]       = ""
sales["_Matched_VIOL"]     = ""

def append_trace(current: str, entry: str) -> str:
    return (current + " | " + entry) if current else entry

# ============================================================
# MATCHING: OK then VIOL then UNSAT
# OK condition: purchase_abs_price < sale_price
# nearest: max purchase_abs_price below sale_price
# ============================================================

for sidx, sale in sales.iterrows():
    sale_qty   = int(sale["Number"])
    sale_price = float(sale["Price"])
    sale_date  = sale["Date"]

    # 1) OK: eligible purchases with _P < sale_price, nearest first => _P DESC, Date ASC
    ok = purchases[
        (purchases["_RemainingQty"] > 0) &
        (purchases["_P"] < sale_price)
    ].sort_values(["_P", "Date"], ascending=[False, True])

    for pidx, prow in ok.iterrows():
        if sale_qty <= 0:
            break

        take = min(int(purchases.at[pidx, "_RemainingQty"]), sale_qty)
        if take <= 0:
            continue

        purchases.at[pidx, "_RemainingQty"] -= take
        sale_qty -= take
        sales.at[sidx, "_Matched_OK_Qty"] += take

        # purchase trace uses the sale info
        p_entry = f"{sale_date.date()}@{sale_price}:-{take}"
        purchases.at[pidx, "_Trace_OK"] = append_trace(purchases.at[pidx, "_Trace_OK"], p_entry)

        # sale trace uses the ORIGINAL purchase price (negative), but could also show _P if you prefer
        s_entry = f"{purchases.at[pidx, 'Price']}:-{take}"
        sales.at[sidx, "_Matched_OK"] = append_trace(sales.at[sidx, "_Matched_OK"], s_entry)

    # 2) VIOL: still sell remaining inventory but flag (purchase_abs_price >= sale_price)
    if sale_qty > 0:
        viol = purchases[
            (purchases["_RemainingQty"] > 0) &
            (purchases["_P"] >= sale_price)
        ].sort_values(["_P", "Date"], ascending=[True, True])  # “nearest violation” first

        for pidx, prow in viol.iterrows():
            if sale_qty <= 0:
                break

            take = min(int(purchases.at[pidx, "_RemainingQty"]), sale_qty)
            if take <= 0:
                continue

            purchases.at[pidx, "_RemainingQty"] -= take
            sale_qty -= take
            sales.at[sidx, "_Matched_VIOL_Qty"] += take

            p_entry = f"{sale_date.date()}@{sale_price}:-{take}"
            purchases.at[pidx, "_Trace_VIOL"] = append_trace(purchases.at[pidx, "_Trace_VIOL"], p_entry)

            s_entry = f"{purchases.at[pidx, 'Price']}:-{take}"
            sales.at[sidx, "_Matched_VIOL"] = append_trace(sales.at[sidx, "_Matched_VIOL"], s_entry)

    # 3) true unsatisfied (short)
    if sale_qty > 0:
        sales.at[sidx, "_UnsatisfiedQty"] = sale_qty

# OUTPUT
purchases_out = purchases.rename(columns={
    "Date": "PurchaseDate",
    "Number": "PurchasedQty",
    "Price": "PurchasePrice",
    "_P": "PurchaseAbsPrice",
    "_RemainingQty": "RemainingQty",
    "_Trace_OK": "SalesTrace_OK",
    "_Trace_VIOL": "SalesTrace_VIOL",
})[[
    "PurchaseDate","PurchasedQty","PurchasePrice","PurchaseAbsPrice",
    "RemainingQty","SalesTrace_OK","SalesTrace_VIOL"
]]

sales_out = sales.rename(columns={
    "Date": "SaleDate",
    "Number": "SaleQty",
    "Price": "SalePrice",
    "_Matched_OK_Qty": "Matched_OK_Qty",
    "_Matched_VIOL_Qty": "Matched_VIOL_Qty",
    "_UnsatisfiedQty": "UnsatisfiedQty",
    "_Matched_OK": "Matched_OK_Purchases",
    "_Matched_VIOL": "Matched_VIOL_Purchases",
})[[
    "SaleDate","SaleQty","SalePrice",
    "Matched_OK_Qty","Matched_VIOL_Qty","UnsatisfiedQty",
    "Matched_OK_Purchases","Matched_VIOL_Purchases"
]]

with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
    purchases_out.to_excel(writer, sheet_name="Purchases", index=False)
    sales_out.to_excel(writer, sheet_name="Sales", index=False)

print(f"✅ Written: {Path(OUTPUT_XLSX).resolve()}")
import pandas as pd
from pathlib import Path

# ============================================================
# SETTINGS
# ============================================================

INPUT_XLSX  = "trades.xlsx"
INPUT_SHEET = "Trades"
OUTPUT_XLSX = "FIFO_Comparison.xlsx"

# ============================================================
# LOAD DATA
# ============================================================

df = pd.read_excel(INPUT_XLSX, sheet_name=INPUT_SHEET)
df = df.iloc[:, :3]
df.columns = ["Date", "Number", "Price"]

df["Date"]   = pd.to_datetime(df["Date"], errors="coerce")
df["Number"] = df["Number"].astype(int)
df["Price"]  = df["Price"].astype(float)

# ============================================================
# CORE FIFO ENGINE (PARAMETRIC)
# ============================================================

def run_fifo(purchases, sales):
    p_idx = s_idx = 0

    purchases = purchases.copy()
    sales = sales.copy()

    purchases["RemainingQty"] = purchases["Number"]
    purchases["Sales_OK"]   = ""
    purchases["Sales_VIOL"] = ""

    def add(cur, txt):
        return txt if not cur else f"{cur} | {txt}"

    ok_value = viol_value = viol_qty = 0

    while p_idx < len(purchases) and s_idx < len(sales):

        p_price = purchases.at[p_idx, "PPrice"]
        s_price = sales.at[s_idx, "SPrice"]

        p_rem = purchases.at[p_idx, "RemainingQty"]
        s_rem = sales.at[s_idx, "RemainingQty"]

        take = min(p_rem, s_rem)

        if s_price >= p_price:
            purchases.at[p_idx, "Sales_OK"] = add(
                purchases.at[p_idx, "Sales_OK"],
                f"{s_price}:-{take}"
            )
            ok_value += (s_price - p_price) * take
        else:
            purchases.at[p_idx, "Sales_VIOL"] = add(
                purchases.at[p_idx, "Sales_VIOL"],
                f"{s_price}:-{take}"
            )
            viol_qty += take
            viol_value += (p_price - s_price) * take

        purchases.at[p_idx, "RemainingQty"] -= take
        sales.at[s_idx, "RemainingQty"] -= take

        if purchases.at[p_idx, "RemainingQty"] == 0:
            p_idx += 1
        if sales.at[s_idx, "RemainingQty"] == 0:
            s_idx += 1

    metrics = {
        "TotalPurchaseQty": int(purchases["Number"].sum()),
        "TotalSaleQty": int(sales["OriginalQty"].sum()),
        "RemainingPurchaseQty": int(purchases["RemainingQty"].sum()),
        "ViolationQty": int(viol_qty),
        "OK_Value": ok_value,
        "ViolationValue": viol_value,
        "ViolationToOKRatio": viol_value / ok_value if ok_value else 0
    }

    return purchases, metrics

# ============================================================
# MODEL A — CLASSICAL FIFO (DATE-BASED)
# ============================================================

purchases_A = (
    df[df["Price"] < 0]
    .assign(PPrice=lambda x: -x["Price"])
    .sort_values("Date")
    .reset_index(drop=True)
)

sales_A = (
    df[df["Price"] > 0]
    .assign(SPrice=lambda x: x["Price"])
    .sort_values("Date")
    .assign(
        OriginalQty=lambda x: x["Number"],
        RemainingQty=lambda x: x["Number"]
    )
    .reset_index(drop=True)
)

out_A, metrics_A = run_fifo(purchases_A, sales_A)

# ============================================================
# MODEL B — GLOBAL PRICE-CENTRIC FIFO
# ============================================================

purchases_B = (
    df[df["Price"] < 0]
    .assign(PPrice=lambda x: -x["Price"])
    .sort_values(["PPrice", "Date"])
    .reset_index(drop=True)
)

sales_B = (
    df[df["Price"] > 0]
    .groupby("Price", as_index=False)["Number"].sum()
    .rename(columns={"Price": "SPrice", "Number": "OriginalQty"})
    .sort_values("SPrice")
    .assign(RemainingQty=lambda x: x["OriginalQty"])
    .reset_index(drop=True)
)

out_B, metrics_B = run_fifo(purchases_B, sales_B)

# ============================================================
# METRICS COMPARISON
# ============================================================

metrics_df = pd.DataFrame([
    {"Model": "Classical_FIFO", **metrics_A},
    {"Model": "Global_Price_FIFO", **metrics_B},
])

# ============================================================
# WRITE OUTPUT
# ============================================================

with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as w:
    out_A.to_excel(w, "FIFO_Classical_Purchases", index=False)
    out_B.to_excel(w, "FIFO_PriceCentric_Purchases", index=False)
    metrics_df.to_excel(w, "Metrics_Comparison", index=False)

print(f"✅ Written: {Path(OUTPUT_XLSX).resolve()}")
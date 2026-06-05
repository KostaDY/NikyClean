#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import json
from pathlib import Path

# ============================================================
# FILES
# ============================================================

XLSX_FILE = Path("Base_Portfolio.xlsx")
JSON_FILE = Path("Base_Portfolio.json")
SHEET_NAME = "Port"

# ============================================================
# LOAD EXCEL SHEET
# ============================================================

df = pd.read_excel(XLSX_FILE, sheet_name=SHEET_NAME)

# Replace NaN with None so JSON contains null
df = df.where(pd.notnull(df), None)

# Convert to list of records
data = df.to_dict(orient="records")

# ============================================================
# SAVE JSON
# ============================================================

with open(JSON_FILE, "w", encoding="utf-8") as f:
    json.dump(data, f, indent=4, ensure_ascii=False, default=str)

print(f"✅ Created: {JSON_FILE}")
print(f"✅ Rows exported: {len(data)}")

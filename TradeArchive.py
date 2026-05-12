import pandas as pd
from pathlib import Path
import logging
from datetime import datetime

# ============================================================
# LOGGING SETUP
# ============================================================

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# ============================================================
# SETTINGS
# ============================================================

WORKBOOK = Path("Entry_RSI.xlsm")
WORKSHEET = "Entries"
NAMED_RANGE = "TradeArchive"
OUTPUT_CSV = Path("TradeArchive.csv")

# ============================================================
# READ NAMED RANGE AND EXPORT TO CSV
# ============================================================

logger.info(f"Reading named range '{NAMED_RANGE}' from {WORKBOOK}")

try:
    # Read the entire workbook to access named ranges
    excel_file = pd.ExcelFile(WORKBOOK, engine='openpyxl')
    
    # Load the workbook to get the named range definition
    from openpyxl import load_workbook
    wb = load_workbook(WORKBOOK, data_only=True)
    
    # Get the named range
    if NAMED_RANGE not in wb.defined_names:
        logger.error(f"Named range '{NAMED_RANGE}' not found in workbook")
        raise ValueError(f"Named range '{NAMED_RANGE}' does not exist")
    
    # Get the range reference
    named_range = wb.defined_names[NAMED_RANGE]
    
    # Parse the range (format is typically 'SheetName!$A$1:$Z$100')
    for dest in named_range.destinations:
        sheet_name, cell_range = dest
        logger.info(f"Found range: {sheet_name}!{cell_range}")
        
        # Verify it's the correct worksheet
        if sheet_name != WORKSHEET:
            logger.warning(f"Named range is on sheet '{sheet_name}', expected '{WORKSHEET}'")
        
        # Parse cell range (e.g., $A$1:$Z$100)
        cell_range = cell_range.replace('$', '')
        
        # Read the specific range using pandas
        # Parse start and end cells
        start_cell, end_cell = cell_range.split(':')
        
        # Extract column letters and row numbers
        import re
        start_match = re.match(r'([A-Z]+)(\d+)', start_cell)
        end_match = re.match(r'([A-Z]+)(\d+)', end_cell)
        
        if not start_match or not end_match:
            raise ValueError(f"Could not parse cell range: {cell_range}")
        
        start_col = start_match.group(1)
        start_row = int(start_match.group(2))
        end_col = end_match.group(1)
        end_row = int(end_match.group(2))
        
        # Convert column letters to column indices
        def col_to_num(col):
            num = 0
            for c in col:
                num = num * 26 + (ord(c.upper()) - ord('A')) + 1
            return num - 1  # 0-indexed
        
        start_col_idx = col_to_num(start_col)
        end_col_idx = col_to_num(end_col)
        
        # Read the data from the specific range
        df = pd.read_excel(
            WORKBOOK,
            sheet_name=sheet_name,
            skiprows=start_row - 1,  # Skip rows before the range
            nrows=end_row - start_row + 1,  # Number of rows to read
            usecols=range(start_col_idx, end_col_idx + 1),  # Columns to read
            engine='openpyxl'
        )
        
        logger.info(f"Read {len(df)} rows and {len(df.columns)} columns from named range")
        
        # Export to CSV
        df.to_csv(OUTPUT_CSV, index=False)
        
        logger.info(f"Successfully exported to {OUTPUT_CSV}")
        logger.info(f"Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        break  # Only process the first destination
    
    wb.close()
    
except FileNotFoundError:
    logger.error(f"File not found: {WORKBOOK}")
    raise
except Exception as e:
    logger.error(f"Error reading named range: {e}")
    raise

logger.info("Done!")
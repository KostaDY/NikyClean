import pandas as pd
import yfinance as yf
from datetime import datetime
from pathlib import Path
import time
import logging
import subprocess
import platform

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

WORKBOOK = Path("/Users/kostayanev/NikyClean/YahooDataOutput.xlsx")
INPUT_SHEET = "Data"
OUTPUT_SHEET = "YahooData"
BACKUP_SHEET = "YahooData_Backup"

MAX_RETRIES = 3
RETRY_DELAY_BASE = 2  # seconds

# ============================================================
# HELPER FUNCTIONS
# ============================================================

def close_excel_file(filepath):
    """Close Excel file if open (macOS-specific using AppleScript)"""
    try:
        if platform.system() == "Darwin":  # macOS
            # Simplified AppleScript that works more reliably
            script = f'''
            tell application "Microsoft Excel"
                set workbookName to "{filepath.name}"
                try
                    close workbook workbookName saving no
                    return "Closed: " & workbookName
                on error errMsg
                    return "Not open or error: " & errMsg
                end try
            end tell
            '''
            result = subprocess.run(['osascript', '-e', script], 
                                   capture_output=True, 
                                   text=True, 
                                   timeout=10)
            
            if result.stdout:
                logger.info(f"Excel: {result.stdout.strip()}")
            if result.stderr:
                logger.debug(f"AppleScript stderr: {result.stderr.strip()}")
            
            # Give Excel time to fully close the file
            time.sleep(2)
            
        else:
            logger.warning("Auto-close only supported on macOS. Please close Excel manually.")
    except subprocess.TimeoutExpired:
        logger.error("Timeout while trying to close Excel")
    except Exception as e:
        logger.warning(f"Could not auto-close Excel: {e}")

def open_excel_file(filepath):
    """Open Excel file (macOS-specific)"""
    try:
        if platform.system() == "Darwin":  # macOS
            subprocess.run(['open', str(filepath)], check=False)
            logger.info(f"Opened {filepath.name}")
        elif platform.system() == "Windows":
            subprocess.run(['start', str(filepath)], shell=True, check=False)
            logger.info(f"Opened {filepath.name}")
        else:
            logger.info(f"Please open {filepath} manually")
    except Exception as e:
        logger.warning(f"Could not auto-open Excel: {e}")

def validate_ticker(ticker):
    """Basic ticker validation"""
    if not ticker or not isinstance(ticker, str):
        return False
    ticker = ticker.strip()
    if len(ticker) == 0 or len(ticker) > 10:
        return False
    # Check for common invalid characters
    if any(char in ticker for char in [' ', '\n', '\t']):
        return False
    return True

def fetch_ticker_data(ticker, max_retries=MAX_RETRIES):
    """Fetch ticker data with retry logic"""
    for attempt in range(max_retries):
        try:
            t = yf.Ticker(ticker)
            info = t.info or {}
            
            # Verify we got valid data
            if not info or info.get('regularMarketPrice') is None:
                if attempt < max_retries - 1:
                    logger.warning(f"{ticker}: No data received, retrying...")
                    time.sleep(RETRY_DELAY_BASE ** attempt)
                    continue
                else:
                    logger.error(f"{ticker}: No valid data after {max_retries} attempts")
                    return None, {}
            
            return t, info
            
        except Exception as e:
            if attempt < max_retries - 1:
                delay = RETRY_DELAY_BASE ** attempt
                logger.warning(f"{ticker}: Error on attempt {attempt + 1}/{max_retries} - {e}. Retrying in {delay}s...")
                time.sleep(delay)
            else:
                logger.error(f"{ticker}: Failed after {max_retries} attempts - {e}")
                return None, {}
    
    return None, {}

def get_ex_dividend_date(info):
    """Extract and format ex-dividend date"""
    ex_raw = info.get("exDividendDate")
    if isinstance(ex_raw, (int, float)) and ex_raw > 0:
        try:
            return datetime.fromtimestamp(ex_raw).date()
        except (ValueError, OSError):
            return None
    return None

def get_earnings_date(ticker_obj, ticker_symbol):
    """Extract earnings date with error handling"""
    try:
        ed = ticker_obj.earnings_dates
        if isinstance(ed, pd.DataFrame) and not ed.empty:
            return ed.index[0].date()
    except Exception as e:
        logger.debug(f"{ticker_symbol}: No earnings date available - {e}")
    return None

def get_target_prices(info):
    """Extract mean, low, and high analyst target prices"""
    target_mean = info.get("targetMeanPrice")
    target_low = info.get("targetLowPrice")
    target_high = info.get("targetHighPrice")
    num_analysts = info.get("numberOfAnalystOpinions")
    
    mean = round(target_mean, 2) if isinstance(target_mean, (int, float)) and target_mean > 0 else None
    low = round(target_low, 2) if isinstance(target_low, (int, float)) and target_low > 0 else None
    high = round(target_high, 2) if isinstance(target_high, (int, float)) and target_high > 0 else None
    num_opinions = num_analysts if isinstance(num_analysts, int) and num_analysts > 0 else None
    
    return mean, low, high, num_opinions

def get_dividend_info(info):
    """Extract dividend yield and dividend rate"""
    div_yield = info.get("dividendYield")
    div_rate = info.get("dividendRate")
    
    div_yield_pct = round(div_yield * 100, 2) if isinstance(div_yield, (int, float)) and div_yield >= 0 else None
    div_rate_val = round(div_rate, 2) if isinstance(div_rate, (int, float)) and div_rate >= 0 else None
    
    return div_yield_pct, div_rate_val

# ============================================================
# CLOSE EXCEL FILE IF OPEN
# ============================================================

logger.info("Closing Excel file if open...")
close_excel_file(WORKBOOK)

# ============================================================
# READ TICKERS (A2:A74)
# ============================================================

logger.info(f"Reading tickers from {WORKBOOK}")

try:
    df_tickers = pd.read_excel(
        WORKBOOK,
        sheet_name=INPUT_SHEET,
        usecols=[0],
        skiprows=1,
        nrows=73,
        header=None
    )
except Exception as e:
    logger.error(f"Failed to read Excel file: {e}")
    raise

# Clean and validate tickers
raw_tickers = (
    df_tickers.iloc[:, 0]
    .dropna()
    .astype(str)
    .str.strip()
    .tolist()
)

tickers = [t for t in raw_tickers if validate_ticker(t)]

logger.info(f"Found {len(tickers)} valid tickers (filtered from {len(raw_tickers)} raw entries)")

if len(tickers) != len(raw_tickers):
    invalid = set(raw_tickers) - set(tickers)
    logger.warning(f"Removed invalid tickers: {invalid}")

# ============================================================
# FETCH DATA
# ============================================================

rows = []
successful = 0
failed = 0

logger.info("Starting data fetch...")

for i, ticker in enumerate(tickers, 1):
    logger.info(f"[{i}/{len(tickers)}] Fetching {ticker}...")
    
    t, info = fetch_ticker_data(ticker)
    
    if t is None:
        failed += 1
        # Still add row with ticker to maintain alignment
        rows.append({
            "Ticker": ticker,
            "Ex-Dividend Date": None,
            "Earnings Date": None,
            "1-Year Target (Mean)": None,
            "Target Low": None,
            "Target High": None,
            "# of Analysts": None,
            "Div Yield (%)": None,
            "Div ($)": None
        })
        continue
    
    # Extract all data points
    ex_div_date = get_ex_dividend_date(info)
    earn_date = get_earnings_date(t, ticker)
    target_mean, target_low, target_high, num_analysts = get_target_prices(info)
    div_yield_pct, div_rate = get_dividend_info(info)
    
    rows.append({
        "Ticker": ticker,
        "Ex-Dividend Date": ex_div_date,
        "Earnings Date": earn_date,
        "1-Year Target (Mean)": target_mean,
        "Target Low": target_low,
        "Target High": target_high,
        "# of Analysts": num_analysts,
        "Div Yield (%)": div_yield_pct,
        "Div ($)": div_rate
    })
    
    successful += 1
    
    # Small delay to avoid rate limiting
    if i < len(tickers):
        time.sleep(0.5)

logger.info(f"Fetch complete: {successful} successful, {failed} failed")

# ============================================================
# WRITE TO EXCEL
# ============================================================

out_df = pd.DataFrame(rows)

logger.info(f"Writing results to {WORKBOOK}")

try:
    # Create backup of existing data
    try:
        existing_df = pd.read_excel(WORKBOOK, sheet_name=OUTPUT_SHEET)
        with pd.ExcelWriter(
            WORKBOOK,
            engine="openpyxl",
            mode="a",
            if_sheet_exists="replace"
        ) as writer:
            existing_df.to_excel(writer, sheet_name=BACKUP_SHEET, index=False)
        logger.info(f"Created backup in sheet '{BACKUP_SHEET}'")
    except Exception as e:
        logger.warning(f"Could not create backup: {e}")
    
    # Write new data
    with pd.ExcelWriter(
        WORKBOOK,
        engine="openpyxl",
        mode="a",
        if_sheet_exists="replace"
    ) as writer:
        out_df.to_excel(writer, sheet_name=OUTPUT_SHEET, index=False)
    
    logger.info(f"Successfully wrote {len(out_df)} rows to '{OUTPUT_SHEET}'")
    logger.info(f"Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
except Exception as e:
    logger.error(f"Failed to write to Excel: {e}")
    # Save to CSV as fallback
    backup_csv = WORKBOOK.parent / f"YahooData_fallback_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    out_df.to_csv(backup_csv, index=False)
    logger.info(f"Data saved to fallback CSV: {backup_csv}")
    raise

# ============================================================
# OPEN EXCEL FILE
# ============================================================

logger.info("Opening Excel file...")
open_excel_file(WORKBOOK)

logger.info("Done!")
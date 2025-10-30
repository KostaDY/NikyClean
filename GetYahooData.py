
import pandas as pd
from yahooquery import Ticker
import requests, time, random
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor, as_completed
from tqdm import tqdm

# ------------------------------------------------------------
# Configuration
# ------------------------------------------------------------
INPUT_FILE   = "MEDY.xlsx"
OUTPUT_FILE  = "YahooDataOutput.xlsx"
OUTPUT_SHEET = "YData"

HEADERS      = {"User-Agent": "Mozilla/5.0"}
MAX_WORKERS  = 6
RETRY_DELAY  = (1.0, 2.5)     # seconds between retries
TIMEOUT      = 8              # per request

# ------------------------------------------------------------
def get_calendar_page(ticker):
    """Download and parse Yahoo Finance calendar page."""
    try:
        url = f"https://finance.yahoo.com/quote/{ticker}/calendar"
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
        if r.status_code != 200:
            return None
        return BeautifulSoup(r.text, "lxml")
    except Exception:
        return None


def extract_earnings_date(soup):
    """Extract 'Earnings Date' from the HTML page."""
    if not soup:
        return ""
    try:
        span = soup.find("span", string="Earnings Date")
        if span:
            next_span = span.find_next("span")
            return next_span.text.strip() if next_span else ""
    except Exception:
        pass
    return ""


# ------------------------------------------------------------
def fetch_ticker_data_once(ticker):
    """Fetch Yahoo Finance data for one ticker (single attempt)."""
    try:
        yq = Ticker(ticker)
        info        = yq.summary_detail if isinstance(yq.summary_detail, dict) else {}
        price_info  = yq.price if isinstance(yq.price, dict) else {}
        financials  = yq.financial_data if isinstance(yq.financial_data, dict) else {}

        info        = info.get(ticker, {}) if isinstance(info.get(ticker, {}), dict) else {}
        price_info  = price_info.get(ticker, {}) if isinstance(price_info.get(ticker, {}), dict) else {}
        financials  = financials.get(ticker, {}) if isinstance(financials.get(ticker, {}), dict) else {}

        dividend_yield = info.get("dividendYield")
        ex_div_date    = info.get("exDividendDate")

        # Convert numeric timestamp → date string
        try:
            if isinstance(ex_div_date, (int, float)):
                ex_div_date_str = pd.to_datetime(ex_div_date, unit="s").strftime("%Y-%m-%d")
            else:
                ex_div_date_str = pd.to_datetime(ex_div_date).strftime("%Y-%m-%d")
        except Exception:
            ex_div_date_str = ""

        target_price = financials.get("targetMeanPrice", "")
        currency     = price_info.get("currency", "")

        # Earnings date via API or HTML fallback
        earnings_date = ""
        try:
            cal = yq.calendar_events
            earn_raw = cal.get(ticker, {}).get("earningsDate", [])
            if isinstance(earn_raw, list) and earn_raw:
                earnings_date = ", ".join(
                    pd.to_datetime(d, unit="s").strftime("%Y-%m-%d") for d in earn_raw
                )
        except Exception:
            pass

        if not earnings_date:
            s_calendar = get_calendar_page(ticker)
            earnings_date = extract_earnings_date(s_calendar)

        return {
            "Ticker": ticker,
            "1-Year Target Price": target_price,
            "Ex-Dividend Date": ex_div_date_str,
            "Earnings Date": earnings_date,
            "Dividend Yield (%)": dividend_yield,
            "Currency": currency,
        }

    except Exception as e:
        return {"Ticker": ticker, "Error": str(e)}


# ------------------------------------------------------------
def fetch_ticker_data(ticker):
    """Fetch with one retry if first attempt fails; prints progress per ticker."""
    start = time.time()
    result = fetch_ticker_data_once(ticker)
    if "Error" in result or (not result.get("Earnings Date") and not result.get("Ex-Dividend Date")):
        time.sleep(random.uniform(*RETRY_DELAY))
        result = fetch_ticker_data_once(ticker)

    # Print per-ticker diagnostic
    elapsed = time.time() - start
    edate = result.get("Earnings Date", "")
    print(f"  → {ticker:<10s} handled in {elapsed:4.1f}s | Earnings: {edate or '—'}")

    if "Error" in result:
        print(f"⚠️  {ticker}: {result['Error']}")
        result.pop("Error", None)
    return result


# ------------------------------------------------------------
def main():
    try:
        # --- Read tickers from MEDY.xlsx, sheet "TickerList", range A2:A75 ---
        df = pd.read_excel(
            INPUT_FILE,
            sheet_name="TickerList",
            usecols="A",
            header=0      # first row (A1) is header "Ticker"
        )

        tickers = (
            df.iloc[0:74, 0]           # A2:A75 (74 rows)
            .astype(str)
            .str.strip()
            .tolist()
        )
        tickers = [t for t in tickers if t and t.lower() not in ("nan", "none")]

        print(f"✅ {len(tickers)} tickers loaded from {INPUT_FILE}.\n")

        # --- Concurrent execution (preserve original Excel order) ---
        results_map = {}
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            futures = {executor.submit(fetch_ticker_data, t): t for t in tickers}
            for future in tqdm(as_completed(futures), total=len(futures), desc="Fetching", ncols=90):
                t = futures[future]
                results_map[t] = future.result()

        # Rebuild results in Excel order
        results = [results_map[t] for t in tickers]
        df_result = pd.DataFrame(results)

        # --- Write results to Excel ---
        with pd.ExcelWriter(OUTPUT_FILE, engine="xlsxwriter") as writer:
            df_result.to_excel(writer, index=False, sheet_name=OUTPUT_SHEET)
            wb, ws = writer.book, writer.sheets[OUTPUT_SHEET]
            if "Dividend Yield (%)" in df_result.columns:
                fmt = wb.add_format({"num_format": "0.00%"})
                col = df_result.columns.get_loc("Dividend Yield (%)")
                ws.set_column(col, col, 18, fmt)

        print(f"\n✅ Final data saved to {OUTPUT_FILE} → sheet '{OUTPUT_SHEET}'")

    except FileNotFoundError:
        print(f"❌ Input file '{INPUT_FILE}' not found.")
    except Exception as e:
        print(f"❌ Unexpected error: {e}")


# ------------------------------------------------------------
if __name__ == "__main__":
    main()

import pandas as pd
from yahooquery import Ticker
import requests
from bs4 import BeautifulSoup

HEADERS = {"User-Agent": "Mozilla/5.0"}

def get_calendar_page(ticker):
    url = f"https://finance.yahoo.com/quote/{ticker}/calendar"
    return BeautifulSoup(requests.get(url, headers=HEADERS).text, "lxml")

def extract_earnings_date(soup):
    try:
        span = soup.find("span", string="Earnings Date")
        if span:
            next_span = span.find_next("span")
            return next_span.text.strip() if next_span else ""
    except Exception as e:
        print(f"  [!] Earnings Date not found: {e}")
    return ""

def fetch_ticker_data(ticker):
    print(f"📡 Fetching data for {ticker}...")
    try:
        yq = Ticker(ticker)

        # Safely check for dict-type data before calling .get()
        info = yq.summary_detail if isinstance(yq.summary_detail, dict) else {}
        price_info = yq.price if isinstance(yq.price, dict) else {}
        financials = yq.financial_data if isinstance(yq.financial_data, dict) else {}

        info = info.get(ticker, {}) if isinstance(info.get(ticker, {}), dict) else {}
        price_info = price_info.get(ticker, {}) if isinstance(price_info.get(ticker, {}), dict) else {}
        financials = financials.get(ticker, {}) if isinstance(financials.get(ticker, {}), dict) else {}

        dividend_yield = info.get("dividendYield")
        dividend_yield_pct = dividend_yield if dividend_yield else None

        ex_div_date = info.get("exDividendDate")
        try:
            if isinstance(ex_div_date, (int, float)):
                ex_div_date_str = pd.to_datetime(ex_div_date, unit='s').strftime('%Y-%m-%d')
            else:
                ex_div_date_str = pd.to_datetime(ex_div_date).strftime('%Y-%m-%d')
        except:
            ex_div_date_str = ""

        target_price = financials.get("targetMeanPrice", "")
        currency = price_info.get("currency", "")

        s_calendar = get_calendar_page(ticker)
        earnings_date = extract_earnings_date(s_calendar)

        return {
            "Ticker": ticker,
            "1-Year Target Price": target_price,
            "Ex-Dividend Date": ex_div_date_str,
            "Earnings Date": earnings_date,
            "Dividend Yield (%)": dividend_yield_pct,
            "Currency": currency
        }

    except Exception as e:
        print(f"❌ Error fetching {ticker}: {e}")
        return {
            "Ticker": ticker,
            "1-Year Target Price": "",
            "Ex-Dividend Date": "",
            "Earnings Date": "",
            "Dividend Yield (%)": None,
            "Currency": ""
        }

def main():
    try:
        df = pd.read_excel("TickerList.xlsx")
        tickers = df.iloc[:, 0].dropna().astype(str).tolist()
        results = []

        for ticker in tickers:
            data = fetch_ticker_data(ticker)
            results.append(data)

        df_result = pd.DataFrame(results)

        # Export with percent formatting
        with pd.ExcelWriter("RTMED.xlsm", engine="xlsxwriter") as writer:
            df_result.to_excel(writer, index=False, sheet_name="YData")
            workbook = writer.book
            worksheet = writer.sheets["Data"]

            if "Dividend Yield (%)" in df_result.columns:
                percent_format = workbook.add_format({"num_format": "0.00%"})
                col_idx = df_result.columns.get_loc("Dividend Yield (%)")
                worksheet.set_column(col_idx, col_idx, 18, percent_format)

        print("\n✅ Final data saved to YahooDataOutput.xlsx")

    except FileNotFoundError:
        print("❌ TickerList.xlsx not found.")
    except Exception as e:
        print(f"❌ An unexpected error occurred: {e}")

if __name__ == "__main__":
    main()

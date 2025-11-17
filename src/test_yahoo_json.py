import requests
import json

headers = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "application/json",
    "Accept-Language": "en-US,en;q=0.9",
}

url = "https://query2.finance.yahoo.com/v10/finance/quoteSummary/AAPL"
params = {"modules": "price,summaryDetail"}

r = requests.get(url, headers=headers)
print("STATUS:", r.status_code)

try:
    data = r.json()
    print(json.dumps(data, indent=2))
except Exception as e:
    print("JSON ERROR:", e)
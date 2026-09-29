import os
import sys
import json
import time
import urllib.request
import urllib.parse
from datetime import datetime

DATA_DIR = os.path.join(os.path.dirname(__file__), "data")
os.makedirs(DATA_DIR, exist_ok=True)

DEFAULT_TICKERS = {
    "TQQQ": "ProShares UltraPro QQQ (나스닥 3배)",
    "SOXL": "Direxion Daily Semiconductor Bull 3X (반도체 3배)",
    "UPRO": "ProShares UltraPro S&P 500 (S&P 3배)",
    "QLD": "ProShares Ultra QQQ (나스닥 2배)",
    "QQQ": "Invesco QQQ Trust (나스닥 1배)",
    "SPY": "SPDR S&P 500 ETF (S&P 1배)",
    "122630.KS": "KODEX 레버리지 (코스피 2배)"
}

def fetch_ticker_data(ticker: str, range_str: str = "10y") -> dict:
    url = f"https://query1.finance.yahoo.com/v8/finance/chart/{urllib.parse.quote(ticker)}?range={range_str}&interval=1d"
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    )
    
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            data = json.loads(resp.read().decode('utf-8'))
            
        chart = data.get("chart", {}).get("result", [])
        if not chart:
            print(f"[-] No chart data for {ticker}")
            return None
            
        res0 = chart[0]
        timestamps = res0.get("timestamp", [])
        quote = res0.get("indicators", {}).get("quote", [{}])[0]
        
        opens = quote.get("open", [])
        highs = quote.get("high", [])
        lows = quote.get("low", [])
        closes = quote.get("close", [])
        volumes = quote.get("volume", [])
        
        meta = res0.get("meta", {})
        reg_price = meta.get("regularMarketPrice")
        reg_time = meta.get("regularMarketTime")

        cleaned = []
        for i in range(len(timestamps)):
            c = closes[i]
            if c is None or c <= 0:
                if i == len(timestamps) - 1 and reg_price and reg_price > 0:
                    c = reg_price
                else:
                    continue
            o = opens[i] if opens[i] is not None else c
            h = highs[i] if highs[i] is not None else max(o, c)
            l = lows[i] if lows[i] is not None else min(o, c)
            v = volumes[i] if volumes[i] is not None else 0
            
            dt_str = datetime.utcfromtimestamp(timestamps[i]).strftime("%Y-%m-%d")
            cleaned.append({
                "date": dt_str,
                "open": round(float(o), 4),
                "high": round(float(h), 4),
                "low": round(float(l), 4),
                "close": round(float(c), 4),
                "volume": int(v)
            })

        # meta의 최신 날짜가 timestamps보다 더 최신인 경우 안전하게 추가 보정
        if reg_price and reg_price > 0 and reg_time:
            meta_dt = datetime.utcfromtimestamp(reg_time).strftime("%Y-%m-%d")
            if not cleaned or meta_dt > cleaned[-1]["date"]:
                cleaned.append({
                    "date": meta_dt,
                    "open": round(float(reg_price), 4),
                    "high": round(float(reg_price), 4),
                    "low": round(float(reg_price), 4),
                    "close": round(float(reg_price), 4),
                    "volume": 0
                })
            
        return {
            "ticker": ticker,
            "name": DEFAULT_TICKERS.get(ticker, ticker),
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "data_count": len(cleaned),
            "start_date": cleaned[0]["date"] if cleaned else "",
            "end_date": cleaned[-1]["date"] if cleaned else "",
            "candles": cleaned
        }
    except Exception as e:
        print(f"[!] Error fetching {ticker}: {e}")
        return None

def main():
    print(f"Fetching historical data for DCA Signal ({len(DEFAULT_TICKERS)} tickers)...")
    manifest = {}
    
    for ticker, name in DEFAULT_TICKERS.items():
        print(f" -> Fetching {ticker} ({name})...", end=" ", flush=True)
        t_data = fetch_ticker_data(ticker, range_str="10y")
        if t_data and t_data["candles"]:
            filename = f"{ticker.replace('.', '_')}.json"
            filepath = os.path.join(DATA_DIR, filename)
            with open(filepath, "w", encoding="utf-8") as f:
                json.dump(t_data, f, ensure_ascii=False, indent=2)
            manifest[ticker] = {
                "name": name,
                "filename": filename,
                "data_count": t_data["data_count"],
                "start_date": t_data["start_date"],
                "end_date": t_data["end_date"]
            }
            print(f"OK ({t_data['data_count']} days, {t_data['start_date']} ~ {t_data['end_date']})")
        else:
            print("FAILED")
        time.sleep(1)
        
    manifest_path = os.path.join(DATA_DIR, "manifest.json")
    with open(manifest_path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    print(f"\n[+] All datasets saved to {DATA_DIR}/, manifest generated.")

if __name__ == "__main__":
    main()

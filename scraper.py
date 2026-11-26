import os, re, json
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo
from curl_cffi import requests
import yfinance as yf
import pandas as pd

TZ = ZoneInfo("Africa/Nairobi")
MIN_LOTS = 100.0
BIAS_PCT = 0.75

SESSION_MAP = {
    "CHFJPY": "TIGER", "AUDJPY": "TIGER", "GBPJPY": "TIGER", "NZDJPY": "TIGER",
    "AUDUSD": "TIGER", "EURJPY": "TIGER", "AUDCHF": "TIGER", "NZDUSD": "TIGER",
    "AUDCAD": "TIGER", "AUDNZD": "TIGER", "USDJPY": "TIGER", "CADJPY": "TIGER",
    "EURCHF": "LONDON", "EURUSD": "LONDON", "NZDCHF": "LONDON", "GBPCHF": "LONDON",
    "EURCAD": "LONDON", "CADCHF": "LONDON", "GBPUSD": "LONDON", "EURGBP": "LONDON",
    "GBPCAD": "LONDON", "GBPNZD": "LONDON", "EURNZD": "LONDON", "GBPAUD": "LONDON",
    "EURAUD": "LONDON",
    "NZDCAD": "NEW YORK", "USDCHF": "NEW YORK", "USDCAD": "NEW YORK"
}
ALL_PAIRS = list(SESSION_MAP.keys())

def fetch_sentiment(pairs):
    url = "https://www." + "myfxbook" + ".com/community/outlook"
    res = requests.get(url, impersonate="chrome", timeout=25)
    text = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", res.text))
    names = "|".join(sorted(pairs, key=len, reverse=True))
    not_a_pair = rf"(?:(?!\b(?:{names})\b).)"
    pattern = re.compile(
        rf"\b({names})\b{not_a_pair}{{1,400}}?"
        rf"(?i:short)\s+(\d+(?:\.\d+)?)\s*%\s+([\d,]+(?:\.\d+)?)\s+(?i:lots)"
        rf"{not_a_pair}{{1,150}}?"
        rf"(?i:long)\s+(\d+(?:\.\d+)?)\s*%\s+([\d,]+(?:\.\d+)?)\s+(?i:lots)"
    )
    out = {}
    for m in pattern.finditer(text):
        p = m.group(1)
        if p in out: continue
        sp, sv = float(m.group(2)) / 100.0, float(m.group(3).replace(",", ""))
        lp, lv = float(m.group(4)) / 100.0, float(m.group(5).replace(",", ""))
        tot = sv + lv
        if abs(sp + lp - 1.0) <= 0.02 and (tot == 0 or abs(sv / tot - sp) <= 0.05):
            out[p] = {"net_sale": sp, "net_sale_vol": sv, "net_buy": lp, "net_buy_vol": lv}
    return out

def wilder_rsi(close, period=14):
    if close is None or len(close) < period + 1: return None
    d = close.diff()
    ag = d.clip(lower=0).ewm(alpha=1/period, min_periods=period, adjust=False).mean()
    al = (-d.clip(upper=0)).ewm(alpha=1/period, min_periods=period, adjust=False).mean()
    val = (100 - (100 / (1 + ag / al))).iloc[-1]
    return round(float(val), 1) if pd.notna(val) else None

def fetch_rsi(pairs, period, interval, mins):
    tickers = [f"{p}=X" for p in pairs]
    out = {p: None for p in pairs}
    df = yf.download(tickers, period=period, interval=interval, group_by="ticker", progress=False, threads=True, auto_adjust=False)
    now = datetime.now(timezone.utc)
    for p, t in zip(pairs, tickers):
        try:
            c = df[t]["Close"].dropna()
            if c.empty: continue
            ls = c.index[-1]
            ls = ls.tz_convert("UTC") if ls.tzinfo else ls.tz_localize("UTC")
            if ls + timedelta(minutes=mins) > now: c = c.iloc[:-1]
            out[p] = wilder_rsi(c)
        except Exception: pass
    return out

def main():
    sent = fetch_sentiment(ALL_PAIRS)
    if not sent: return
    r1 = fetch_rsi(ALL_PAIRS, "10d", "1h", 60)
    r15 = fetch_rsi(ALL_PAIRS, "5d", "15m", 15)
    stamp = datetime.now(TZ).strftime("%Y-%m-%d %H:%M EAT")
    rows = []
    for p in ALL_PAIRS:
        d = sent.get(p)
        if not d: continue
        tot = round(d["net_sale_vol"] + d["net_buy_vol"], 2)
        net = round(d["net_buy_vol"] - d["net_sale_vol"], 2)
        bias = "LOW VOL" if tot < MIN_LOTS else ("BUY" if d["net_sale"] >= BIAS_PCT else ("SELL" if d["net_buy"] >= BIAS_PCT else "NONE"))
        rows.append({"pair": p, "session": SESSION_MAP[p], "net_sale": round(d["net_sale"] * 100, 1), "net_sale_vol": d["net_sale_vol"], "net_buy": round(d["net_buy"] * 100, 1), "net_buy_vol": d["net_buy_vol"], "total_lots": tot, "net_lots": net, "rsi_1h": r1.get(p), "rsi_15m": r15.get(p), "bias": bias})
    os.makedirs("docs", exist_ok=True)
    with open("docs/latest.json", "w") as f:
        json.dump({"updated": stamp, "pairs": rows}, f, indent=2)

if __name__ == "__main__":
    main()

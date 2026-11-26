import os, re, json, time, hashlib
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo
from curl_cffi import requests
import yfinance as yf
import pandas as pd

TZ = ZoneInfo("Africa/Nairobi")
MIN_LOTS = 100.0
BIAS_PCT = 0.75

SESSIONS_MAP = {
    "USDJPY": ["TIGER", "LONDON", "LONDON-NY", "NEW YORK"],
    "AUDUSD": ["TIGER", "NEW YORK"],
    "NZDUSD": ["TIGER", "NEW YORK"],
    "EURJPY": ["TIGER", "LONDON", "LONDON-NY"],
    "GBPJPY": ["TIGER", "LONDON", "LONDON-NY"],
    "AUDJPY": ["TIGER", "LONDON"],
    "EURAUD": ["TIGER", "LONDON"],
    "GBPAUD": ["TIGER", "LONDON"],
    "CHFJPY": ["TIGER", "LONDON"],
    "CADJPY": ["TIGER", "NEW YORK"],
    "NZDJPY": ["TIGER"],
    "AUDCAD": ["TIGER", "NEW YORK"],
    "AUDCHF": ["TIGER", "LONDON"],
    "AUDNZD": ["TIGER"],
    "GBPNZD": ["TIGER", "LONDON"],
    "EURNZD": ["TIGER", "LONDON"],
    "EURUSD": ["LONDON", "LONDON-NY", "NEW YORK"],
    "GBPUSD": ["LONDON", "LONDON-NY", "NEW YORK"],
    "USDCHF": ["LONDON", "LONDON-NY", "NEW YORK"],
    "USDCAD": ["LONDON", "LONDON-NY", "NEW YORK"],
    "EURGBP": ["LONDON"],
    "EURCHF": ["LONDON"],
    "EURCAD": ["LONDON", "LONDON-NY", "NEW YORK"],
    "NZDCHF": ["LONDON"],
    "GBPCAD": ["LONDON", "LONDON-NY", "NEW YORK"],
    "GBPCHF": ["LONDON"],
    "CADCHF": ["LONDON", "LONDON-NY", "NEW YORK"],
    "NZDCAD": ["NEW YORK"],
}
ALL_PAIRS = list(SESSIONS_MAP.keys())

def fetch_sentiment(pairs):
    cb = int(time.time())
    host = "www." + "myfxbook" + ".com"
    proxy = "r." + "jina" + ".ai"
    urls = [
        (f"https://{proxy}/https://{host}/community/outlook?cb={cb}", {"X-No-Cache": "true"}),
        (f"https://{host}/community/outlook", {})
    ]
    names = "|".join(sorted(pairs, key=len, reverse=True))
    not_a_pair = rf"(?:(?!\b(?:{names})\b).)"
    pattern = re.compile(
        rf"\b({names})\b{not_a_pair}{{1,400}}?"
        rf"(?i:short)\s+(\d+(?:\.\d+)?)\s*%\s+([\d,]+(?:\.\d+)?)\s+(?i:lots)"
        rf"{not_a_pair}{{1,150}}?"
        rf"(?i:long)\s+(\d+(?:\.\d+)?)\s*%\s+([\d,]+(?:\.\d+)?)\s+(?i:lots)"
    )
    out = {}
    for u, hdrs in urls:
        try:
            res = requests.get(u, headers=hdrs, impersonate="chrome124", timeout=25)
            raw = res.text
            text = re.sub(r"[|\[\]()]+", " ", re.sub(r"<[^>]+>", " ", raw))
            text = re.sub(r"\s+", " ", text)
            for m in pattern.finditer(text):
                p = m.group(1)
                if p in out:
                    continue
                sp, sv = float(m.group(2)) / 100.0, float(m.group(3).replace(",", ""))
                lp, lv = float(m.group(4)) / 100.0, float(m.group(5).replace(",", ""))
                tot = sv + lv
                if abs(sp + lp - 1.0) <= 0.02 and (tot == 0 or abs(sv / tot - sp) <= 0.05):
                    out[p] = {"net_sale": sp, "net_sale_vol": sv, "net_buy": lp, "net_buy_vol": lv}
            if len(out) >= 10:
                data_fp = hashlib.sha256(json.dumps(out, sort_keys=True).encode()).hexdigest()[:12]
                return out, {"source": u, "page_len": len(raw), "data_hash": data_fp}
        except Exception as e:
            print("Fetch error:", u, e)
    return out, {"source": "none", "page_len": 0, "data_hash": ""}

def wilder_rsi(close, period=14):
    if close is None or len(close) < period + 1:
        return None
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
            if c.empty:
                continue
            ls = c.index[-1]
            ls = ls.tz_convert("UTC") if ls.tzinfo else ls.tz_localize("UTC")
            if ls + timedelta(minutes=mins) > now:
                c = c.iloc[:-1]
            out[p] = wilder_rsi(c)
        except Exception:
            pass
    return out

def main():
    sent, meta = fetch_sentiment(ALL_PAIRS)
    if not sent:
        raise RuntimeError("0 pairs matched")
    r1 = fetch_rsi(ALL_PAIRS, "10d", "1h", 60)
    r15 = fetch_rsi(ALL_PAIRS, "5d", "15m", 15)
    stamp = datetime.now(TZ).strftime("%Y-%m-%d %H:%M EAT")

    prev_hash = ""
    if os.path.exists("docs/latest.json"):
        try:
            prev_hash = json.load(open("docs/latest.json")).get("fingerprint", {}).get("data_hash", "")
        except Exception:
            pass
    meta["sentiment_changed"] = (meta["data_hash"] != prev_hash)

    rows = []
    for p in ALL_PAIRS:
        d = sent.get(p)
        if not d:
            continue
        tot = round(d["net_sale_vol"] + d["net_buy_vol"], 2)
        net = round(d["net_buy_vol"] - d["net_sale_vol"], 2)
        bias = "LOW VOL" if tot < MIN_LOTS else ("BUY" if d["net_sale"] >= BIAS_PCT else ("SELL" if d["net_buy"] >= BIAS_PCT else "NONE"))
        sess_list = SESSIONS_MAP[p]
        rows.append({
            "pair": p, "sessions": sess_list, "session": " • ".join(sess_list),
            "net_sale": round(d["net_sale"] * 100, 1), "net_sale_vol": d["net_sale_vol"],
            "net_buy": round(d["net_buy"] * 100, 1), "net_buy_vol": d["net_buy_vol"],
            "total_lots": tot, "net_lots": net,
            "rsi_1h": r1.get(p), "rsi_15m": r15.get(p), "bias": bias
        })

    os.makedirs("docs", exist_ok=True)
    with open("docs/latest.json", "w") as f:
        json.dump({"updated": stamp, "fingerprint": meta, "pairs": rows}, f, indent=2)
    print("Saved:", stamp, meta)

if __name__ == "__main__":
    main()

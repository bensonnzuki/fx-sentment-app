import os, re, json, time, hashlib
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo
from curl_cffi import requests
import yfinance as yf
import pandas as pd

TZ = ZoneInfo("Africa/Nairobi")
MIN_LOTS = 100.0
BIAS_PCT = 0.75
EXCEL_PATH = "/home/bnzuki/CloudStation/Forex-APP/forex_sentiment_tracker.xlsx"

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

def parse_outlook(raw, pairs, out):
    text = re.sub(r"[|\[\]()]+", " ", re.sub(r"<[^>]+>", " ", raw))
    text = re.sub(r"\s+", " ", text)
    names = "|".join(sorted(pairs, key=len, reverse=True))
    not_a_pair = rf"(?:(?!\b(?:{names})\b).)"
    pat_a = re.compile(
        rf"\b({names})\b{not_a_pair}{{1,400}}?"
        rf"(?i:short)\s+(\d+(?:\.\d+)?)\s*%\s+([\d,]+(?:\.\d+)?)\s+(?i:lots)"
        rf"{not_a_pair}{{1,150}}?"
        rf"(?i:long)\s+(\d+(?:\.\d+)?)\s*%\s+([\d,]+(?:\.\d+)?)\s+(?i:lots)"
    )
    pat_b = re.compile(
        rf'symbolName["\':=\s]+({names})["\']?.*?'
        rf'shortPercentage["\':=\s]+(\d+(?:\.\d+)?).*?'
        rf'longPercentage["\':=\s]+(\d+(?:\.\d+)?).*?'
        rf'shortVolume["\':=\s]+([\d.]+).*?'
        rf'longVolume["\':=\s]+([\d.]+)',
        re.S | re.I
    )
    matched = 0
    for m in pat_a.finditer(text):
        p = m.group(1)
        if p in out: continue
        sp, sv = float(m.group(2)) / 100.0, float(m.group(3).replace(",", ""))
        lp, lv = float(m.group(4)) / 100.0, float(m.group(5).replace(",", ""))
        tot = sv + lv
        if abs(sp + lp - 1.0) <= 0.02 and (tot == 0 or abs(sv / tot - sp) <= 0.05):
            out[p] = {"net_sale": sp, "net_sale_vol": sv, "net_buy": lp, "net_buy_vol": lv}
            matched += 1
    if len(out) < len(pairs):
        for m in pat_b.finditer(raw):
            p = m.group(1).upper()
            if p in out: continue
            sp, lp = float(m.group(2)) / 100.0, float(m.group(3)) / 100.0
            sv, lv = float(m.group(4)), float(m.group(5))
            if abs(sp + lp - 1.0) <= 0.02:
                out[p] = {"net_sale": sp, "net_sale_vol": sv, "net_buy": lp, "net_buy_vol": lv}
                matched += 1
    return matched, text

def fetch_browser_html(url):
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            b = p.chromium.launch(headless=True, args=["--disable-blink-features=AutomationControlled"])
            ctx = b.new_context(user_agent="Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36")
            page = ctx.new_page()
            page.goto(url, wait_until="domcontentloaded", timeout=35000)
            page.wait_for_timeout(6000)
            html = page.content()
            b.close()
            return html
    except Exception as e:
        print("Browser fetch warning:", e)
        return ""

def fetch_sentiment(pairs):
    host = "www." + "myfxbook" + ".com"
    proxy = "r." + "jina" + ".ai"
    target_url = f"https://{host}/community/outlook"
    sources = [
        ("browser", target_url, {}),
        ("jina", f"https://{proxy}/{target_url}", {})
    ]
    out, logs, raw_len, used = {}, [], 0, []
    for label, u, hdrs in sources:
        if len(out) == len(pairs):
            break
        try:
            if label == "browser":
                raw = fetch_browser_html(u)
                status = 200 if len(raw) > 10000 else 403
            else:
                res = requests.get(u, headers=hdrs, impersonate="chrome124", timeout=25)
                raw, status = res.text, res.status_code
            if label == "direct":
                open("/tmp/myfxbook_direct_debug.html", "w").write(raw[:200000])
            added, clean_txt = parse_outlook(raw, pairs, out)
            if label == "direct":
                open("/tmp/myfxbook_direct_clean.txt", "w").write(clean_txt[:100000])
            logs.append(f"{label}:status={status},len={len(raw)},added={added}")
            if added > 0:
                used.append(label)
                raw_len = len(raw)
        except Exception as e:
            logs.append(f"{label}:err={type(e).__name__}")
    print("Fetch diagnostics:", " | ".join(logs))
    data_fp = hashlib.sha256(json.dumps(out, sort_keys=True).encode()).hexdigest()[:12] if out else ""
    return out, {"source": "+".join(used) or "none", "page_len": raw_len, "data_hash": data_fp, "diag": " | ".join(logs)}

def wilder_rsi(close, period=14):
    if close is None or len(close) < period + 1:
        return None
    d = close.diff()
    ag = d.clip(lower=0).ewm(alpha=1/period, min_periods=period, adjust=False).mean()
    al = (-d.clip(upper=0)).ewm(alpha=1/period, min_periods=period, adjust=False).mean()
    val = (100 - (100 / (1 + ag / al))).iloc[-1]
    return round(float(val), 1) if pd.notna(val) else None

def fetch_rsi(pairs, period, interval, mins):
    out = {p: None for p in pairs}
    try:
        tickers = [f"{p}=X" for p in pairs]
        df = yf.download(tickers, period=period, interval=interval, group_by="ticker", progress=False, threads=True, auto_adjust=False)
        now = datetime.now(timezone.utc)
        for p, t in zip(pairs, tickers):
            try:
                c = df[t]["Close"].dropna()
                if c.empty: continue
                ls = c.index[-1]
                ls = ls.tz_convert("UTC") if ls.tzinfo else ls.tz_localize("UTC")
                if ls + timedelta(minutes=mins) > now:
                    c = c.iloc[:-1]
                out[p] = wilder_rsi(c)
            except Exception:
                pass
    except Exception as e:
        print(f"RSI ({interval}) fetch warning:", e)
    return out

def update_excel(stamp, rows, changed):
    if not os.path.exists(EXCEL_PATH):
        return
    try:
        import openpyxl
        wb = openpyxl.load_workbook(EXCEL_PATH)
        if "History" in wb.sheetnames and changed:
            ws = wb["History"]
            for r in rows:
                if r["bias"] != "NO DATA":
                    ws.append([stamp, r["pair"], r["session"], r["net_sale"], r["net_buy"], r["total_lots"], r["net_lots"], r["rsi_1h"], r["rsi_15m"], r["bias"]])
        wb.save(EXCEL_PATH)
        print("Synced Excel workbook:", EXCEL_PATH)
    except Exception as e:
        print("Excel sync skipped:", e)

def main():
    sent, meta = fetch_sentiment(ALL_PAIRS)
    if not sent:
        raise RuntimeError("0 pairs matched from all sources")
    r1 = fetch_rsi(ALL_PAIRS, "10d", "1h", 60)
    r15 = fetch_rsi(ALL_PAIRS, "5d", "15m", 15)
    now_eat = datetime.now(TZ)
    stamp = now_eat.strftime("%Y-%m-%d %H:%M EAT")
    hhmm = now_eat.strftime("%H:%M")

    prev = {}
    if os.path.exists("docs/latest.json"):
        try:
            prev = json.load(open("docs/latest.json"))
        except Exception:
            pass
    prev_fp = prev.get("fingerprint", {})
    prev_hash = prev_fp.get("data_hash", "")
    changed = (meta["data_hash"] != prev_hash)
    since = hhmm if (changed or not prev_fp.get("sentiment_since")) else prev_fp.get("sentiment_since")
    meta["sentiment_changed"] = changed
    meta["sentiment_since"] = since

    rows = []
    for p in ALL_PAIRS:
        d = sent.get(p)
        sess_list = SESSIONS_MAP[p]
        if not d:
            rows.append({
                "pair": p, "sessions": sess_list, "session": " • ".join(sess_list),
                "net_sale": 0.0, "net_sale_vol": 0.0, "net_buy": 0.0, "net_buy_vol": 0.0,
                "total_lots": 0.0, "net_lots": 0.0,
                "rsi_1h": r1.get(p), "rsi_15m": r15.get(p), "bias": "NO DATA"
            })
            continue
        tot = round(d["net_sale_vol"] + d["net_buy_vol"], 2)
        net = round(d["net_buy_vol"] - d["net_sale_vol"], 2)
        bias = "LOW VOL" if tot < MIN_LOTS else ("BUY" if d["net_sale"] >= BIAS_PCT else ("SELL" if d["net_buy"] >= BIAS_PCT else "NONE"))
        rows.append({
            "pair": p, "sessions": sess_list, "session": " • ".join(sess_list),
            "net_sale": round(d["net_sale"] * 100, 1), "net_sale_vol": d["net_sale_vol"],
            "net_buy": round(d["net_buy"] * 100, 1), "net_buy_vol": d["net_buy_vol"],
            "total_lots": tot, "net_lots": net,
            "rsi_1h": r1.get(p), "rsi_15m": r15.get(p), "bias": bias
        })

    prev_sig = [(x["pair"], x["net_sale"], x["total_lots"], x["rsi_15m"], x["bias"]) for x in prev.get("pairs", [])]
    curr_sig = [(x["pair"], x["net_sale"], x["total_lots"], x["rsi_15m"], x["bias"]) for x in rows]
    any_data_changed = (curr_sig != prev_sig)

    os.makedirs("docs", exist_ok=True)
    with open("docs/latest.json", "w") as f:
        json.dump({"updated": stamp, "fingerprint": meta, "pairs": rows, "any_changed": any_data_changed}, f, indent=2)
    update_excel(stamp, rows, changed)
    print(f"Saved: {stamp} | changed={changed} | since={since} | any_changed={any_data_changed}")

if __name__ == "__main__":
    main()

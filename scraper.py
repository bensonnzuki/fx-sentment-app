"""SessionEdge data builder.

One Myfxbook fetch per run produces BOTH outputs:
  * docs/latest.json  (the app, published through GitHub Pages)
  * the v4 Excel tracker (main sheet + History sheet)

Sources are tried in order (direct -> browser -> jina) and their results are merged,
so a pair missing from one source can still come from the next.
"""
import os, re, sys, json, time, shutil, hashlib, tempfile
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo

from curl_cffi import requests
import yfinance as yf
import pandas as pd

# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------
TZ = ZoneInfo("Africa/Nairobi")
MIN_LOTS = 100.0          # below this total the % is treated as unreliable (LOW VOL)
BIAS_PCT = 0.75           # crowd share needed for a contrarian bias
HIGHLIGHT_PCT = 0.70      # Excel colour threshold for the % cells
HISTORY_DAYS = 14         # Excel History sheet retention
HEARTBEAT_MIN = 25        # republish at least this often even if nothing changed
JSON_PATH = "docs/latest.json"
EXCEL_FILE = "/home/bnzuki/CloudStation/Forex-APP/forex_net_buy_sale_tracker-v4.xlsx"
DEBUG_DIR = "/tmp/myfxbook_debug"   # raw page of every source, overwritten each run

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
TARGET_URL = "https://www." + "myfxbook" + ".com/community/outlook"
JINA_URL = "https://r." + "jina" + ".ai/" + TARGET_URL


# ---------------------------------------------------------------------------
# Parsing
# ---------------------------------------------------------------------------
def _valid(sp, sv, lp, lv):
    """The two % must add to ~100 and the % must agree with the lots."""
    tot = sv + lv
    return abs(sp + lp - 1.0) <= 0.02 and (tot == 0 or abs(sv / tot - sp) <= 0.05)


def parse_outlook(raw, pairs, out, rejected):
    """Add valid readings for pairs not yet in `out`. Returns (added, clean_text)."""
    text = re.sub(r"[|\[\]()]+", " ", re.sub(r"<[^>]+>", " ", raw))
    text = re.sub(r"\s+", " ", text)

    tokens = []
    for p in sorted(pairs, key=len, reverse=True):
        tokens += [p, f"{p[:3]}/{p[3:]}"]
    names = "|".join(re.escape(t) for t in tokens)
    gap = rf"(?:(?!\b(?:{names})\b).)"          # may not run across another pair
    num = r"(\d+(?:\.\d+)?)\s*%\s+([\d,]+(?:\.\d+)?)\s+(?i:lots)"
    pat_short_first = re.compile(rf"\b({names})\b{gap}{{1,400}}?(?i:short)\s+{num}{gap}{{1,150}}?(?i:long)\s+{num}")
    pat_long_first = re.compile(rf"\b({names})\b{gap}{{1,400}}?(?i:long)\s+{num}{gap}{{1,150}}?(?i:short)\s+{num}")

    # Embedded JSON fallback, bounded: each gap is limited and may not cross another symbol.
    sgap = r"(?:(?!symbolName).){0,300}?"
    pat_json = re.compile(
        rf'symbolName["\':=\s]+["\']?({names})["\']?{sgap}'
        rf'shortPercentage["\':=\s]+(\d+(?:\.\d+)?){sgap}'
        rf'longPercentage["\':=\s]+(\d+(?:\.\d+)?){sgap}'
        rf'shortVolume["\':=\s]+([\d.]+){sgap}'
        rf'longVolume["\':=\s]+([\d.]+)',
        re.S | re.I)

    added = 0

    def record(pair_raw, sp, sv, lp, lv):
        nonlocal added
        pair = pair_raw.replace("/", "").upper()
        if pair in out or pair not in pairs:
            return
        if not _valid(sp, sv, lp, lv):
            rejected.append(f"{pair}: short {sp:.0%} {sv} / long {lp:.0%} {lv}")
            return
        out[pair] = {"net_sale": sp, "net_sale_vol": sv, "net_buy": lp, "net_buy_vol": lv}
        added += 1

    f = lambda s: float(s.replace(",", ""))
    for m in pat_short_first.finditer(text):
        record(m.group(1), f(m.group(2)) / 100, f(m.group(3)), f(m.group(4)) / 100, f(m.group(5)))
    for m in pat_long_first.finditer(text):
        record(m.group(1), f(m.group(4)) / 100, f(m.group(5)), f(m.group(2)) / 100, f(m.group(3)))
    if len(out) < len(pairs):
        for m in pat_json.finditer(raw):
            record(m.group(1), f(m.group(2)) / 100, f(m.group(4)), f(m.group(3)) / 100, f(m.group(5)))
    return added, text


# ---------------------------------------------------------------------------
# Fetching
# ---------------------------------------------------------------------------
def fetch_browser_html(url):
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("Browser source skipped: playwright not installed")
        return ""
    try:
        with sync_playwright() as p:
            b = p.chromium.launch(headless=True, args=["--disable-blink-features=AutomationControlled"])
            ctx = b.new_context(user_agent="Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                                           "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36")
            page = ctx.new_page()
            page.goto(url, wait_until="domcontentloaded", timeout=40000)
            page.wait_for_timeout(12000)   # give the bot check and the data time to load
            html = page.content()
            b.close()
            return html
    except Exception as e:
        print("Browser fetch warning:", e)
        return ""


def save_debug(label, raw, clean):
    try:
        os.makedirs(DEBUG_DIR, exist_ok=True)
        with open(os.path.join(DEBUG_DIR, f"{label}.html"), "w", encoding="utf-8") as fh:
            fh.write(raw[:400000])
        with open(os.path.join(DEBUG_DIR, f"{label}.txt"), "w", encoding="utf-8") as fh:
            fh.write(clean[:200000])
    except Exception as e:
        print("Debug save warning:", e)


def fetch_sentiment(pairs):
    sources = ["direct", "browser", "jina"]   # direct works from the home PC; browser/jina only fill gaps
    out, rejected, logs, used = {}, [], [], []
    for label in sources:
        if len(out) == len(pairs):
            break
        try:
            if label == "browser":
                raw = fetch_browser_html(TARGET_URL)
                status = "ok" if raw else "empty"
            elif label == "direct":
                res = requests.get(TARGET_URL, impersonate="chrome", timeout=25)
                raw, status = res.text, res.status_code
            else:
                res = requests.get(JINA_URL, headers={"X-No-Cache": "true"}, impersonate="chrome124", timeout=30)
                raw, status = res.text, res.status_code
            added, clean = parse_outlook(raw, pairs, out, rejected) if raw else (0, "")
            if raw:
                save_debug(label, raw, clean)
            logs.append(f"{label}:status={status},len={len(raw)},added={added}")
            if added:
                used.append(label)
        except Exception as e:
            logs.append(f"{label}:err={type(e).__name__}")
    for r in rejected:
        print("Rejected:", r)
    print("Fetch diagnostics:", " | ".join(logs))
    data_hash = hashlib.sha256(json.dumps(out, sort_keys=True).encode()).hexdigest()[:12] if out else ""
    return out, {"source": "+".join(used) or "none", "data_hash": data_hash, "diag": " | ".join(logs)}


# ---------------------------------------------------------------------------
# RSI (closed candles only, one batch download per timeframe)
# ---------------------------------------------------------------------------
def wilder_rsi(close, period=14):
    if close is None or len(close) < period + 1:
        return None
    d = close.diff()
    ag = d.clip(lower=0).ewm(alpha=1 / period, min_periods=period, adjust=False).mean()
    al = (-d.clip(upper=0)).ewm(alpha=1 / period, min_periods=period, adjust=False).mean()
    val = (100 - (100 / (1 + ag / al))).iloc[-1]
    return round(float(val), 1) if pd.notna(val) else None


def fetch_rsi(pairs, period, interval, mins):
    out = {p: None for p in pairs}
    try:
        tickers = [f"{p}=X" for p in pairs]
        df = yf.download(tickers, period=period, interval=interval, group_by="ticker",
                         progress=False, threads=True, auto_adjust=False)
        now = datetime.now(timezone.utc)
        for p, t in zip(pairs, tickers):
            try:
                c = df[t]["Close"].dropna()
                if c.empty:
                    continue
                ls = c.index[-1]
                ls = ls.tz_convert("UTC") if ls.tzinfo else ls.tz_localize("UTC")
                if ls + timedelta(minutes=mins) > now:
                    c = c.iloc[:-1]   # drop the candle that is still forming
                out[p] = wilder_rsi(c)
            except Exception:
                pass
    except Exception as e:
        print(f"RSI ({interval}) fetch warning:", e)
    return out


# ---------------------------------------------------------------------------
# Excel (v4 tracker: main sheet + History)
# ---------------------------------------------------------------------------
def update_excel(now_naive, rows, sentiment_changed):
    if not os.path.exists(EXCEL_FILE):
        print("Excel skipped: file not found", EXCEL_FILE)
        return
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    except ImportError:
        print("Excel skipped: openpyxl not installed")
        return

    GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    GREEN_FONT = Font(name="Calibri", size=11, bold=True, color="006100")
    RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    RED_FONT = Font(name="Calibri", size=11, bold=True, color="9C0006")
    GREY_FONT = Font(name="Calibri", size=11, bold=True, color="808080")
    NORM_FONT = Font(name="Calibri", size=11, bold=False, color="000000")
    BOLD_FONT = Font(name="Calibri", size=11, bold=True, color="000000")
    HEADER_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    ZEBRA_FILL = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    WHITE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    NO_FILL = PatternFill(fill_type=None)
    thin = Side(style="thin", color="D9D9D9")
    dbl = Side(style="double", color="8EA9DB")
    top = Side(style="thin", color="8EA9DB")
    CELL_BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
    HEADER_BORDER = Border(left=top, right=top, top=top, bottom=dbl)
    HEADERS = ["PAIR", "SESSION", "NET-SALE", "NET-SALE-VOLUME", "NET-BUY", "NET-BUY-VOLUME",
               "1H RSI", "15M RSI", "TOTAL LOTS", "NET LOTS (BUY-SELL)", "BIAS"]
    HISTORY_HEADERS = ["TIME", "PAIR", "NET-SALE", "NET-BUY", "SALE LOTS", "BUY LOTS",
                       "TOTAL LOTS", "NET LOTS", "1H RSI", "15M RSI", "BIAS"]

    def style(cell, value, fmt=None, fill=None, font=NORM_FONT, align="center"):
        cell.value = value
        if fmt:
            cell.number_format = fmt
        cell.fill = fill or NO_FILL
        cell.font = font
        cell.alignment = Alignment(horizontal=align, vertical="center")
        cell.border = CELL_BORDER

    def rsi_style(v, rf):
        if v is not None and v >= 70:
            return RED_FILL, RED_FONT
        if v is not None and v <= 30:
            return GREEN_FILL, GREEN_FONT
        return rf, NORM_FONT

    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb["Net Buyers & Sellers"] if "Net Buyers & Sellers" in wb.sheetnames else wb.worksheets[0]
        by_pair = {r["pair"]: r for r in rows}

        # Keep the sheet's pair order, add any missing pairs at the end.
        order = []
        for r in range(2, ws.max_row + 1):
            v = ws.cell(row=r, column=1).value
            p = str(v).strip().upper().replace("/", "") if v else ""
            if p in by_pair and p not in order:
                order.append(p)
        order += [p for p in ALL_PAIRS if p not in order]

        for r in range(1, ws.max_row + 5):
            ws.row_dimensions[r].hidden = False
        if ws.max_row > len(order) + 1:
            ws.delete_rows(len(order) + 2, ws.max_row - len(order) - 1)
        for rng in list(ws.merged_cells.ranges):
            ws.unmerge_cells(str(rng))
        for col in range(1, 16):
            c = ws.cell(row=1, column=col)
            c.value, c.fill, c.border = None, NO_FILL, Border()
        ws.row_dimensions[1].height = 28
        for i, title in enumerate(HEADERS, start=1):
            c = ws.cell(row=1, column=i, value=title)
            c.font, c.fill, c.border = BOLD_FONT, HEADER_FILL, HEADER_BORDER
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        history = []
        for idx, pair in enumerate(order):
            row = idx + 2
            d = by_pair[pair]
            rf = ZEBRA_FILL if row % 2 == 1 else WHITE_FILL
            ws.row_dimensions[row].height = 22
            style(ws.cell(row=row, column=1), pair, fill=rf, font=BOLD_FONT)
            style(ws.cell(row=row, column=2), d["session"], fill=rf)
            f1, t1 = rsi_style(d["rsi_1h"], rf)
            f15, t15 = rsi_style(d["rsi_15m"], rf)
            style(ws.cell(row=row, column=7), d["rsi_1h"], "0.0", f1, t1)
            style(ws.cell(row=row, column=8), d["rsi_15m"], "0.0", f15, t15)
            if d["bias"] == "NO DATA":
                for col in (3, 4, 5, 6, 9, 10):
                    style(ws.cell(row=row, column=col), None, fill=rf)
                style(ws.cell(row=row, column=11), "NO DATA", fill=rf, font=GREY_FONT)
                continue
            ns, nb = d["net_sale"] / 100, d["net_buy"] / 100
            style(ws.cell(row=row, column=3), ns, "0.0%", RED_FILL if ns >= HIGHLIGHT_PCT else rf,
                  RED_FONT if ns >= HIGHLIGHT_PCT else NORM_FONT, "right")
            style(ws.cell(row=row, column=4), d["net_sale_vol"], "#,##0.00", rf, align="right")
            style(ws.cell(row=row, column=5), nb, "0.0%", GREEN_FILL if nb >= HIGHLIGHT_PCT else rf,
                  GREEN_FONT if nb >= HIGHLIGHT_PCT else NORM_FONT, "right")
            style(ws.cell(row=row, column=6), d["net_buy_vol"], "#,##0.00", rf, align="right")
            style(ws.cell(row=row, column=9), d["total_lots"], "#,##0.00", rf, align="right")
            style(ws.cell(row=row, column=10), d["net_lots"], "+#,##0.00;-#,##0.00;0.00", rf, align="right")
            bf, bt = {"BUY": (GREEN_FILL, GREEN_FONT), "SELL": (RED_FILL, RED_FONT),
                      "LOW VOL": (rf, GREY_FONT)}.get(d["bias"], (rf, NORM_FONT))
            style(ws.cell(row=row, column=11), d["bias"], fill=bf, font=bt)
            history.append([pair, ns, nb, d["net_sale_vol"], d["net_buy_vol"], d["total_lots"],
                            d["net_lots"], d["rsi_1h"], d["rsi_15m"], d["bias"]])

        for col, w in {"A": 13, "B": 30, "C": 12, "D": 18, "E": 12, "F": 18, "G": 10, "H": 10,
                       "I": 14, "J": 20, "K": 11, "L": 3, "M": 14, "N": 14, "O": 14}.items():
            ws.column_dimensions[col].width = w
        ws.auto_filter.ref = f"A1:K{len(order) + 1}"
        ws.freeze_panes = "A2"
        ws.merge_cells("M1:O1")
        ts = ws["M1"]
        ts.value = f"Last Updated: {now_naive:%Y-%m-%d %H:%M} EAT"
        ts.font = Font(name="Calibri", size=11, bold=True, color="1F4E78")
        ts.fill = HEADER_FILL
        ts.alignment = Alignment(horizontal="center", vertical="center")

        # History: only add rows when the sentiment itself changed (no duplicate readings).
        hs = wb["History"] if "History" in wb.sheetnames else wb.create_sheet("History")
        if hs.cell(row=1, column=1).value != "TIME":
            hs.delete_rows(1, hs.max_row)
            for i, h in enumerate(HISTORY_HEADERS, start=1):
                c = hs.cell(row=1, column=i, value=h)
                c.font, c.fill, c.border = BOLD_FONT, HEADER_FILL, HEADER_BORDER
                c.alignment = Alignment(horizontal="center")
            hs.freeze_panes = "A2"
            for col, w in zip("ABCDEFGHIJK", [18, 10, 10, 10, 12, 12, 12, 12, 9, 9, 10]):
                hs.column_dimensions[col].width = w
        if sentiment_changed:
            for h in history:
                hs.append([now_naive] + h)
                n = hs.max_row
                hs.cell(row=n, column=1).number_format = "yyyy-mm-dd hh:mm"
                for col in (3, 4):
                    hs.cell(row=n, column=col).number_format = "0.0%"
                for col in (5, 6, 7, 8):
                    hs.cell(row=n, column=col).number_format = "#,##0.00"
        cutoff = now_naive - timedelta(days=HISTORY_DAYS)
        old = 0
        for (t,) in hs.iter_rows(min_row=2, max_col=1, values_only=True):
            if isinstance(t, datetime) and t < cutoff:
                old += 1
            else:
                break
        if old:
            hs.delete_rows(2, old)

        # Save: build in /tmp, keep a .bak, then overwrite in place (same file for CloudStation).
        fd, tmp = tempfile.mkstemp(suffix=".xlsx", prefix="forex_tmp_")
        os.close(fd)
        try:
            wb.save(tmp)
            shutil.copy2(EXCEL_FILE, EXCEL_FILE + ".bak")
            with open(tmp, "rb") as src, open(EXCEL_FILE, "wb") as dst:
                dst.write(src.read())
            os.chmod(EXCEL_FILE, 0o664)
        finally:
            if os.path.exists(tmp):
                os.remove(tmp)
        print(f"Excel updated ({len(history)} pairs, history {'added' if sentiment_changed else 'unchanged'})")
        append_oanda_to_excel(rows, sentiment_changed)
    except Exception as e:
        print("Excel update failed:", e)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def fetch_oanda_all(pairs):
    out = {}
    u = "https://www." + "oanda" + ".com/eu-en/market/client-sentiment/"
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as pw:
            b = pw.chromium.launch(headless=True)
            page = b.new_page()
            page.goto(u, wait_until="domcontentloaded", timeout=20000)
            page.wait_for_timeout(3500)
            try:
                page.locator("li, button, a, span").filter(has_text=re.compile(r"^All$")).first.click(timeout=3000)
                page.wait_for_timeout(1500)
            except Exception:
                pass
            txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", page.content()))
            b.close()
        for p in pairs:
            m = re.search(rf"\b{p}\b\s+\b{p}\b\s+(\d+(?:\.\d+)?)\s*%\s+(\d+(?:\.\d+)?)\s*%", txt)
            if not m:
                m = re.search(rf"\b{p}\b.{{1,80}}?(\d+(?:\.\d+)?)\s*%\s+(\d+(?:\.\d+)?)\s*%", txt)
            if m:
                # The header says "% trades that sell / buy", but the numbers are BUY first:
                # the page's own labels match it (GOLD 71% 29% = "71% traders are buying",
                # OILWTI 46% 54% = "54% traders are selling"). First = long, second = short.
                ol_pct, os_pct = float(m.group(1)), float(m.group(2))
                if 98.0 <= (os_pct + ol_pct) <= 102.0:
                    out[p] = {
                        "oanda_short": os_pct, "oanda_long": ol_pct,
                        "stop_pips": None, "stop_side": "", "rvol": None
                    }
    except Exception as e:
        print("OANDA fetch warning:", e)
    print("Public OANDA fetched instruments:", len(out))
    return out

def main():
    sent, meta = fetch_sentiment(ALL_PAIRS)
    if not sent:
        raise RuntimeError("0 pairs matched from all sources")
    r1 = fetch_rsi(ALL_PAIRS, "10d", "1h", 60)
    r15 = fetch_rsi(ALL_PAIRS, "5d", "15m", 15)
    oanda_data = fetch_oanda_all(ALL_PAIRS)
    now = datetime.now(TZ)
    now_naive = now.replace(tzinfo=None, second=0, microsecond=0)
    stamp = now.strftime("%Y-%m-%d %H:%M EAT")

    prev = {}
    if os.path.exists(JSON_PATH):
        try:
            with open(JSON_PATH) as fh:
                prev = json.load(fh)
        except Exception:
            pass
    prev_fp = prev.get("fingerprint", {})
    sentiment_changed = meta["data_hash"] != prev_fp.get("data_hash", "")
    since = stamp if (sentiment_changed or not prev_fp.get("sentiment_since")) else prev_fp["sentiment_since"]
    meta["sentiment_changed"] = sentiment_changed
    meta["sentiment_since"] = since

    rows = []
    for p in ALL_PAIRS:
        d = sent.get(p)
        sess = SESSIONS_MAP[p]
        base = {"pair": p, "sessions": sess, "session": " • ".join(sess),
                "rsi_1h": r1.get(p), "rsi_15m": r15.get(p)}
        if not d:
            rows.append({**base, "net_sale": 0.0, "net_sale_vol": 0.0, "net_buy": 0.0, "net_buy_vol": 0.0,
                         "total_lots": 0.0, "net_lots": 0.0, "bias": "NO DATA"})
            continue
        tot = round(d["net_sale_vol"] + d["net_buy_vol"], 2)
        net = round(d["net_buy_vol"] - d["net_sale_vol"], 2)
        bias = ("LOW VOL" if tot < MIN_LOTS else "BUY" if d["net_sale"] >= BIAS_PCT
                else "SELL" if d["net_buy"] >= BIAS_PCT else "NONE")
        rows.append({**base, "net_sale": round(d["net_sale"] * 100, 1), "net_sale_vol": d["net_sale_vol"],
                     "net_buy": round(d["net_buy"] * 100, 1), "net_buy_vol": d["net_buy_vol"],
                     "total_lots": tot, "net_lots": net, "bias": bias,
            "oanda_short": oanda_data.get(p, {}).get("oanda_short"),
            "oanda_long": oanda_data.get(p, {}).get("oanda_long"),
            "stop_pips": oanda_data.get(p, {}).get("stop_pips"),
            "stop_side": oanda_data.get(p, {}).get("stop_side"),
            "rvol": oanda_data.get(p, {}).get("rvol"),
            "agreed_2of2": (
                (bias == "BUY" and (oanda_data.get(p, {}).get("oanda_short") or 0) >= 70.0) or
                (bias == "SELL" and (oanda_data.get(p, {}).get("oanda_long") or 0) >= 70.0)
            )})

    sig = lambda rs: [(x["pair"], x["net_sale"], x["total_lots"], x["rsi_1h"], x["rsi_15m"], x["bias"]) for x in rs]
    any_changed = sig(rows) != sig(prev.get("pairs", []))
    minutes_since = None
    m = re.match(r"(\d{4}-\d{2}-\d{2} \d{2}:\d{2})", prev.get("updated", ""))
    if m:
        last = datetime.strptime(m.group(1), "%Y-%m-%d %H:%M")
        minutes_since = (now_naive - last).total_seconds() / 60
    heartbeat = minutes_since is None or minutes_since >= HEARTBEAT_MIN

    if any_changed or heartbeat:
        os.makedirs(os.path.dirname(JSON_PATH), exist_ok=True)
        tmp = JSON_PATH + ".tmp"
        with open(tmp, "w") as fh:
            json.dump({"updated": stamp, "fingerprint": meta, "pairs": rows, "any_changed": any_changed}, fh, indent=2)
        os.replace(tmp, JSON_PATH)
        reason = "data changed" if any_changed else f"heartbeat ({minutes_since:.0f} min)" if minutes_since else "first run"
        print(f"Saved {JSON_PATH}: {stamp} ({reason}) | sentiment_changed={sentiment_changed} | since={since}")
        update_excel(now_naive, rows, sentiment_changed)
    else:
        print(f"No change since {prev.get('updated')} - latest.json and Excel left as they are")



def append_oanda_to_excel(rows, changed):
    p = "/home/bnzuki/CloudStation/Forex-APP/forex_net_buy_sale_tracker-v4.xlsx"
    if not os.path.exists(p):
        return
    try:
        import openpyxl
        from copy import copy
        from openpyxl.styles import Font, PatternFill, Alignment

        wb = openpyxl.load_workbook(p)
        by_pair = {r["pair"]: r for r in rows}

        def clone_style(src, dst):
            if src.has_style:
                dst.font = copy(src.font)
                dst.fill = copy(src.fill)
                dst.border = copy(src.border)
                dst.alignment = copy(src.alignment)
                dst.number_format = src.number_format

        for name in wb.sheetnames:
            ws = wb[name]
            # Unmerge M1:O1 if merged for Last Updated
            for rng in list(ws.merged_cells.ranges):
                if rng.min_row == 1 and rng.max_row == 1 and rng.min_col >= 12:
                    ws.unmerge_cells(str(rng))

            # Delete old P/Q/R/S OANDA columns (col >= 15)
            for c in range(ws.max_column, 14, -1):
                hv = str(ws.cell(1, c).value or "").strip().upper()
                if "OANDA" in hv or "AGREED" in hv:
                    ws.delete_cols(c, 1)

            # Move Last Updated note out to P1 (col 16)
            if name != "History":
                lu_txt = ""
                for c in range(12, 18):
                    v = str(ws.cell(1, c).value or "")
                    if v.startswith("Last Updated"):
                        lu_txt = v
                        ws.cell(1, c, None)
                if lu_txt:
                    ws.cell(1, 16, lu_txt).font = Font(bold=True, size=10)

            # Style L1, M1, N1 (cols 12, 13, 14) flush next to K1 (BIAS)
            headers = ["OANDA SHORT %", "OANDA LONG %", "2/2 AGREED"]
            col_widths = [15, 15, 13]
            for idx in range(3):
                col = 12 + idx
                cell = ws.cell(1, col, headers[idx])
                clone_style(ws.cell(1, 11), cell)
                ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = col_widths[idx]

            for r_idx in range(2, ws.max_row + 1):
                pair_val = ""
                for c_idx in range(1, 4):
                    cand = str(ws.cell(r_idx, c_idx).value or "").strip().upper()
                    if cand in by_pair:
                        pair_val = cand
                        break
                if not pair_val:
                    continue
                if name == "History" and r_idx < ws.max_row - len(rows):
                    continue

                d = by_pair[pair_val]
                o_s, o_l = d.get("oanda_short"), d.get("oanda_long")
                ag = "YES" if d.get("agreed_2of2") else "NO"

                c_s = ws.cell(r_idx, 12, (o_s / 100.0) if o_s is not None else "-")
                c_l = ws.cell(r_idx, 13, (o_l / 100.0) if o_l is not None else "-")
                c_a = ws.cell(r_idx, 14, ag)

                for c_dst in (c_s, c_l, c_a):
                    clone_style(ws.cell(r_idx, 11), c_dst)
                    c_dst.alignment = Alignment(horizontal="center", vertical="center")

                if o_s is not None:
                    c_s.number_format = "0.0%"
                    c_l.number_format = "0.0%"

                if ag == "YES":
                    c_a.fill = PatternFill("solid", fgColor="DCFCE7")
                    c_a.font = Font(bold=True, color="166534", size=10)
                else:
                    c_a.fill = PatternFill("solid", fgColor="F1F5F9")
                    c_a.font = Font(bold=True, color="64748B", size=10)

        wb.save(p)
        print("Excel OANDA columns formatted inside table (L, M, N)!")
    except Exception as e:
        print("Excel OANDA sync warning:", e)

if __name__ == "__main__":
    main()

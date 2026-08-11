#!/usr/bin/env python3
"""
generate_telegram_summary.py
----------------------------
Posts an end-of-day (D-1) sales + incentive summary to Telegram.

Runs at 12:01 AM IST daily via GitHub Actions. It always summarises the
PREVIOUS calendar day (D-1); by 12:01 AM the last main-pipeline run
(11:59 PM) has already loaded the full day, so D-1 is complete.

Reads the same public feeds the dashboards use (Sales + Targets + Week-Off)
and sends three Telegram messages:
  1. Daily TSI sell-out summary  (Target / FTD / MTD / %ach / #Zero / #NoUpload)
  2. NPI Tracker                 (FTD / MTD units + % of ISDs participated, MTD)
  3. TSI Incentive summary       (all 7 Aug'26 scheme parameters, Ach/Target : %)

Metric decisions (confirmed with owner):
  - %age achievement = MTD sell-out value / full-month Aug target
  - NPI '% ISDs participated' denominator = ALL of the TSI's ISDs (MTD)
  - Zero-sale / No-upload counts exclude ISDs on their designated week-off day
  - TSIs are attributed by ISP Code -> Targets sheet (not the sales TSI label),
    matching the incentive tracker; Saroj Kumar Rana is remapped to Amit Kumar.

Env vars:
  SALES_CSV_URL, TARGETS_CSV_URL, WEEKOFF_CSV_URL   (public published CSVs)
  TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID
Optional:
  TARGET_DATE   YYYY-MM-DD override (default: yesterday IST) for testing
  DRY_RUN       "1" -> print messages to stdout instead of sending
"""

import os
import io
import re
import csv
import math
import json
import urllib.request
import urllib.parse
from datetime import datetime, timezone, timedelta, date

# ── Config ─────────────────────────────────────────────────────────────────────

IST = timezone(timedelta(hours=5, minutes=30))

# Public published CSVs (same feeds the dashboards use). We fall back to these
# hardcoded URLs whenever the matching env var/secret is missing OR empty
# (an unset GitHub secret arrives as an empty string, which would defeat a
# plain os.environ.get(name, default) — hence the `or`).
SALES_CSV_URL = os.environ.get("SALES_CSV_URL") or \
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vSFtk8td43KgtjhjSIJWePY_WuHG9Xm_JJwtj5OGYZh2BMawqVudIJEICkfWNVWi-dWMfioCeK1R7tr/pub?output=csv"
TARGETS_CSV_URL = os.environ.get("TARGETS_CSV_URL") or \
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vScO4tHpOg-uISI7WnhedipMDO8eVNgsX0UC5AWNNe1eDbAhq2U7ZiaWpL51tvDvAOsVjD3cuwlt2Oz/pub?output=csv"
WEEKOFF_CSV_URL = os.environ.get("WEEKOFF_CSV_URL") or \
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vScO4tHpOg-uISI7WnhedipMDO8eVNgsX0UC5AWNNe1eDbAhq2U7ZiaWpL51tvDvAOsVjD3cuwlt2Oz/pub?gid=1522547784&single=true&output=csv"

TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_CHAT_ID   = os.environ.get("TELEGRAM_CHAT_ID", "")

# ── Incentive scheme constants (mirror incentive-tracker.html CONFIG) ───────────

REMAP = {"Saroj Kumar Rana": "Amit Kumar"}
HC    = {"Amit Kumar": 85, "Joydwip Seal": 75, "Mukesh Kumar": 85, "Nabarun Bhattacharya": 85}
TSIS  = ["Amit Kumar", "Joydwip Seal", "Mukesh Kumar", "Nabarun Bhattacharya"]
RD_EXCLUSIVE = set(TSIS)

SCHEME_YEAR, SCHEME_MONTH = 2026, 8            # August (1-indexed here)
CUT_B1 = date(2026, 8, 9)
CUT_B2 = date(2026, 8, 16)

RE_PREMIUM = re.compile(r"edge|razr|signature", re.I)
RE_FIXEDGE = re.compile(r"edge\s*70\s*pro|edge\s*70\s*max|signature", re.I)
RE_BOOSTER = re.compile(r"edge\s*70\s*pro|edge\s*70\s*max", re.I)
RE_THINGS  = re.compile(r"bud|watch", re.I)

# NPI Tracker models: (label, matcher on lowercased model string)
NPI_MODELS = [
    ("Edge 70 Max",    lambda m: "edge 70 max" in m),
    ("Edge 70 Pro+",   lambda m: "edge 70 pro+" in m),
    ("Edge 70 Pro",    lambda m: "edge 70 pro" in m and "edge 70 pro+" not in m),
    ("Signature",      lambda m: "signature" in m),
    ("Edge 70 Fusion", lambda m: ("edge70 fusion" in m) or ("edge 70 fusion" in m)),
]

_DAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]

# ── Helpers ────────────────────────────────────────────────────────────────────

def rmp(name):
    name = (name or "").strip()
    return REMAP.get(name, name)

def num(x):
    x = ("" if x is None else str(x)).replace(",", "").strip()
    return float(x) if re.match(r"^-?\d+(\.\d+)?$", x) else 0.0

def parse_date(s):
    for fmt in ("%d-%b-%Y", "%d-%b-%y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(str(s).strip(), fmt).date()
        except ValueError:
            pass
    return None

def norm(s):
    return re.sub(r"[\s_\-]+", "", str(s)).lower()

def find_key(row, *cands):
    m = {norm(k): k for k in row}
    for c in cands:
        if norm(c) in m:
            return m[norm(c)]
    return None

def fetch_csv(url):
    req = urllib.request.Request(url + ("&" if "?" in url else "?") + "_t=" +
                                 str(int(datetime.now().timestamp())),
                                 headers={"User-Agent": "moto-telegram-summary"})
    text = urllib.request.urlopen(req, timeout=60).read().decode("utf-8", "replace")
    if text and text[0] == "﻿":
        text = text[1:]
    lines = [l for l in text.splitlines() if l.strip().strip(",")]
    if not lines:
        return []
    # Anchor on the "ISP Code" header row (drop leading totals/spacer rows).
    for idx in range(min(6, len(lines))):
        if re.search(r"isp\s*code", lines[idx], re.I):
            lines = lines[idx:]
            break
    return list(csv.DictReader(lines))

def inr(v):
    """Compact Indian-currency formatting."""
    v = round(v)
    if abs(v) >= 1_00_00_000:
        return "₹%.2f Cr" % (v / 1_00_00_000)
    if abs(v) >= 1_00_000:
        return "₹%.2f L" % (v / 1_00_000)
    return "₹" + format(v, ",d")

def pctf(a, b):
    return round(100 * a / b) if b > 0 else 0

def esc(s):
    """Escape text for Telegram HTML parse_mode."""
    return (str(s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))

STATE_ABBR = {
    "west bengal": "WB", "bihar": "BR", "jharkhand": "JH", "odisha": "OD",
    "orissa": "OD", "assam": "AS", "sikkim": "SK", "tripura": "TR",
    "meghalaya": "ML", "manipur": "MN", "nagaland": "NL", "mizoram": "MZ",
    "arunachal pradesh": "AR", "uttar pradesh": "UP", "jharkand": "JH",
}
CAT_ABBR = {"reliance digital": "RD"}

def loc(info, short=False):
    city = info.get("city") or ""
    st   = info.get("state") or ""
    if short and st:
        st = STATE_ABBR.get(st.strip().lower(), st)
    parts = [p for p in (city, st) if p]
    return ", ".join(parts)

def catshort(cat):
    return CAT_ABBR.get((cat or "").strip().lower(), cat or "—")

def first(name):
    return (name or "").split()[0] if (name or "").strip() else name

def cr(v):
    """Value in ₹ crore, 2 decimals."""
    return v / 1_00_00_000

def pre(lines):
    """Wrap ASCII table lines in a Telegram <pre> monospace block."""
    return "<pre>" + esc("\n".join(lines)) + "</pre>"

# ── Data loading ───────────────────────────────────────────────────────────────

def load_codes():
    """ISP Code -> {tsi, isd, cat, fix, tgt} from the Targets sheet."""
    rows = fetch_csv(TARGETS_CSV_URL)
    r0 = rows[0]
    kC   = find_key(r0, "ISP Code", "ISPCode")
    kN   = find_key(r0, "ISP Name", "ISPName")
    kT   = find_key(r0, "TSI Name", "TSIName")
    kCat = find_key(r0, "Store Category", "Category")
    kFix = find_key(r0, "Fixture Store(Y/N)", "Fixture Store", "Fixture")
    kTgt = find_key(r0, "Aug'26 TGT", "Aug'26 Target", "Aug26 TGT")
    kStr = find_key(r0, "Parent Retailer Name", "Store Name", "Retailer Name")
    kCty = find_key(r0, "City")
    kSt  = find_key(r0, "State")
    codes = {}
    for r in rows:
        c = (r.get(kC) or "").strip()
        if not c or c in codes:
            continue
        codes[c] = {
            "tsi":   rmp(r.get(kT)),
            "isd":   (r.get(kN) or "").strip(),
            "cat":   (r.get(kCat) or "").strip(),
            "fix":   (r.get(kFix) or "").strip(),
            "tgt":   num(r.get(kTgt)),
            "store": (r.get(kStr) or "").strip() if kStr else "",
            "city":  (r.get(kCty) or "").strip() if kCty else "",
            "state": (r.get(kSt)  or "").strip() if kSt  else "",
        }
    return codes

def load_weekoff():
    try:
        rows = fetch_csv(WEEKOFF_CSV_URL)
    except Exception as e:
        print("[TG] WeekOff fetch failed (non-fatal):", e)
        return {}
    if not rows:
        return {}
    r0 = rows[0]
    kC = find_key(r0, "ISP Code", "ISPCode", "Code")
    kD = find_key(r0, "Week Off Day", "WeekOff", "Off Day", "Day")
    out = {}
    for r in rows:
        c = (r.get(kC) or "").strip() if kC else ""
        d = (r.get(kD) or "").strip() if kD else ""
        if c and d:
            out[c] = d
    return out

def load_sales():
    rows = fetch_csv(SALES_CSV_URL)
    r0 = rows[0]
    K = {
        "code":   find_key(r0, "ISP Code", "ISPCode"),
        "model":  find_key(r0, "Model"),
        "cat":    find_key(r0, "ProductCategoryName", "Product Category Name", "Category"),
        "qty":    find_key(r0, "Quantity", "Qty"),
        "price":  find_key(r0, "Price"),
        "status": find_key(r0, "SaleStatus", "Sale Status", "Status"),
        "date":   find_key(r0, "SaleDate", "Sale Date", "Date"),
    }
    out = []
    for r in rows:
        out.append({
            "code":   (r.get(K["code"]) or "").strip(),
            "model":  (r.get(K["model"]) or "").strip(),
            "cat":    (r.get(K["cat"]) or "").strip(),
            "qty":    num(r.get(K["qty"])),
            "price":  num(r.get(K["price"])),
            "status": (r.get(K["status"]) or "").strip(),
            "date":   parse_date(r.get(K["date"])),
        })
    return out

# ── Computation ────────────────────────────────────────────────────────────────

def compute(target_date):
    codes   = load_codes()
    weekoff = load_weekoff()
    sales   = load_sales()

    month_start = date(target_date.year, target_date.month, 1)
    target_day_name = _DAY_NAMES[target_date.weekday()]

    tc = {t: [c for c in codes if codes[c]["tsi"] == t] for t in TSIS}      # TSI -> [codes]

    # Per-code aggregates over the scheme month up to and including D-1 (MTD).
    agg = {c: {"things": 0.0, "premVal": 0.0, "totVal": 0.0, "fixEdge": 0.0,
               "b9": 0.0, "b16": 0.0, "active": False, "sellMTD": 0.0, "sellFTD": 0.0,
               "npiFTD": {m[0]: 0.0 for m in NPI_MODELS},
               "npiMTD": {m[0]: 0.0 for m in NPI_MODELS}} for c in codes}

    # Daily (D-1) status flags per code
    zero_codes = {t: set() for t in TSIS}
    noup_codes = {t: set() for t in TSIS}

    for r in sales:
        c = r["code"]
        if c not in codes:
            continue
        d = r["date"]
        if not d:
            continue
        tsi = codes[c]["tsi"]
        if tsi not in tc:
            continue

        # ── Daily zero / no-upload (D-1 only, exclude week-off) ──
        if d == target_date:
            off = weekoff.get(c, "")
            if not (off and off.lower() == target_day_name.lower()):
                if r["status"] == "Zero Sale Uploaded":
                    zero_codes[tsi].add(c)
                elif r["status"] == "Sale Not Uploaded":
                    noup_codes[tsi].add(c)

        # ── Value/units only from actual uploaded sales, within MTD window ──
        if r["status"] != "Sale Uploaded":
            continue
        q = round(r["qty"])
        if q <= 0:
            continue
        if not (month_start <= d <= target_date):
            continue

        a = agg[c]
        pr, m, cat, ml = r["price"], r["model"], r["cat"], r["model"].lower()
        val = pr * q

        a["sellMTD"] += val
        if d == target_date:
            a["sellFTD"] += val
        if d.year == SCHEME_YEAR and d.month == SCHEME_MONTH and d.day <= 20:
            a["active"] = True
        if cat == "Accessories" and RE_THINGS.search(m):
            a["things"] += q
        if cat == "Smart Phone":
            a["totVal"] += val
            if RE_PREMIUM.search(m):
                a["premVal"] += val
        if RE_FIXEDGE.search(m):
            a["fixEdge"] += q
        if RE_BOOSTER.search(m):
            if d <= CUT_B1:
                a["b9"] += q
            if d <= CUT_B2:
                a["b16"] += q
        for label, match in NPI_MODELS:
            if match(ml):
                a["npiMTD"][label] += q
                if d == target_date:
                    a["npiFTD"][label] += q

    return codes, tc, agg, zero_codes, noup_codes

# ── Message builders ───────────────────────────────────────────────────────────

def build_daily(target_date, codes, tc, agg, zero_codes, noup_codes):
    hdr = target_date.strftime("%d %b %Y")
    tot_t = tot_ftd = tot_mtd = 0.0
    tz = tn = 0
    rows_val = ["%-8s %6s %5s %6s %4s" % ("(Cr)", "Target", "FTD", "MTD", "Ach")]
    rows_gap = ["%-8s %5s %6s" % ("ISDs", "Zero", "NoUpld")]
    for t in TSIS:
        cs  = tc[t]
        tgt = sum(codes[c]["tgt"] for c in cs)
        ftd = sum(agg[c]["sellFTD"] for c in cs)
        mtd = sum(agg[c]["sellMTD"] for c in cs)
        nz  = len(zero_codes[t]); nn = len(noup_codes[t])
        tot_t += tgt; tot_ftd += ftd; tot_mtd += mtd; tz += nz; tn += nn
        rows_val.append("%-8s %6.2f %5.2f %6.2f %3d%%"
                        % (first(t), cr(tgt), cr(ftd), cr(mtd), pctf(mtd, tgt)))
        rows_gap.append("%-8s %5d %6d" % (first(t), nz, nn))
    rows_val.append("-" * 32)
    rows_val.append("%-8s %6.2f %5.2f %6.2f %3d%%"
                    % ("TOTAL", cr(tot_t), cr(tot_ftd), cr(tot_mtd), pctf(tot_mtd, tot_t)))
    rows_gap.append("-" * 21)
    rows_gap.append("%-8s %5d %6d" % ("TOTAL", tz, tn))

    L = ["\U0001F4CA <b>DAILY SELL-OUT</b> — %s <i>(D-1)</i>" % hdr,
         "━" * 20,
         "Sell-out value, in ₹ crore:",
         pre(rows_val),
         "Upload gaps (week-off ISDs excluded):",
         pre(rows_gap)]
    return "\n".join(L)

def build_npi(target_date, codes, tc, agg):
    L = ["\U0001F4F1 <b>NPI TRACKER</b> — %s" % target_date.strftime("%d %b %Y"),
         "━" * 20,
         "<i>FTD &amp; MTD units · %ISD = share of the TSI's ISDs who sold it (MTD)</i>",
         ""]
    for label, _ in NPI_MODELS:
        rows = ["%-8s %4s %4s %5s" % ("", "FTD", "MTD", "%ISD")]
        for t in TSIS:
            cs = tc[t]
            n = len(cs) or 1
            ftd  = sum(agg[c]["npiFTD"][label] for c in cs)
            mtd  = sum(agg[c]["npiMTD"][label] for c in cs)
            part = sum(1 for c in cs if agg[c]["npiMTD"][label] > 0)
            rows.append("%-8s %4d %4d %4d%%"
                        % (first(t), round(ftd), round(mtd), pctf(part, n)))
        L.append("<b>%s</b>" % label)
        L.append(pre(rows))
    return "\n".join(L).rstrip()

def _need(share, base):
    return int(math.ceil(share * base))

def build_incentive(target_date, codes, tc, agg):
    mark = lambda ok: "✅" if ok else "\U0001F538"   # ✅ achieved / 🔸 in progress
    L = ["\U0001F3C6 <b>INCENTIVE TRACKER</b> — %s" % target_date.strftime("%d %b %Y"),
         "<i>Aug'26 scheme · ✅ achieved · 🔸 in progress</i>",
         "━" * 20, ""]
    for t in TSIS:
        cs   = tc[t]
        n    = len(cs)
        rd   = [c for c in cs if codes[c]["cat"] == "Reliance Digital"]
        fx   = [c for c in cs if codes[c]["fix"] == "Yes"]
        hc   = HC.get(t, 0)

        # 1 Moto Things (5 units × ISDs)
        things = sum(agg[c]["things"] for c in cs)
        t_tgt  = 5 * n

        # 2 RD Premium ≥80% (≥70% of RD ISDs)
        rd_ok = sum(1 for c in rd if agg[c]["totVal"] > 0 and
                    agg[c]["premVal"] / agg[c]["totVal"] >= 0.8)
        rd_need = _need(0.7, len(rd))
        rd_prog = min(100, round(100 * (rd_ok / len(rd)) / 0.7)) if rd else 0

        # 3 HC active by 20th (≥90% of sanctioned HC)
        act = sum(1 for c in cs if agg[c]["active"])
        hc_need = _need(0.9, hc)

        # 4 Fixture Edge ≥4 units (≥90% of fixture ISDs)
        fx_ok = sum(1 for c in fx if agg[c]["fixEdge"] >= 4)
        fx_need = _need(0.9, len(fx))
        fx_prog = min(100, round(100 * (fx_ok / len(fx)) / 0.9)) if fx else 0

        # 5 Sell-out achievement (tiered 90% / 100%)
        tgt_sum = sum(codes[c]["tgt"] for c in cs)
        sell    = sum(agg[c]["sellMTD"] for c in cs)
        so_pct  = round(100 * sell / tgt_sum) if tgt_sum else 0

        # 6/7 Boosters
        b1_ok = sum(1 for c in cs if agg[c]["b9"] >= 5)
        b1_need = _need(0.75, n)
        b1_prog = min(100, round(100 * (b1_ok / n) / 0.75)) if n else 0
        b2_ok = sum(1 for c in cs if agg[c]["b16"] >= 15)
        b2_need = _need(0.50, n)
        b2_prog = min(100, round(100 * (b2_ok / n) / 0.50)) if n else 0

        L.append("\U0001F464 <b>%s</b>" % t)
        L.append("<i>%d ISDs · %d RD · %d fixture · HC %d</i>" % (n, len(rd), len(fx), hc))
        L.append("%s Moto Things — %d / %d units (%d%%)"
                 % (mark(things >= t_tgt), round(things), t_tgt, pctf(things, t_tgt)))
        L.append("%s RD Premium ≥80%% — %d / %d ISDs (%d%%, need 70%%)"
                 % (mark(rd_prog >= 100), rd_ok, rd_need, rd_prog))
        L.append("%s HC active by 20th — %d / %d (%d%%)"
                 % (mark(act >= hc_need), act, hc_need, pctf(act, hc_need)))
        L.append("%s Fixture Edge ≥4u — %d / %d ISDs (%d%%, need 90%%)"
                 % (mark(fx_prog >= 100), fx_ok, fx_need, fx_prog))
        L.append("%s Sell-out value — %d%% of target (₹3k@90%%, ₹4k@100%%)"
                 % (mark(so_pct >= 90), so_pct))
        L.append("%s Booster-1 ≥5u by 09 Aug — %d / %d ISDs (%d%%, need 75%%)"
                 % (mark(b1_prog >= 100), b1_ok, b1_need, b1_prog))
        L.append("%s Booster-2 ≥15u by 16 Aug — %d / %d ISDs (%d%%, need 50%%)"
                 % (mark(b2_prog >= 100), b2_ok, b2_need, b2_prog))
        L.append("")
    L.append("<i>Internal reference only — payout subject to activation &amp; "
             "clearance of all eligibility criteria.</i>")
    return "\n".join(L)

def build_isd_lists(target_date, codes, zero_codes, noup_codes):
    """One message per TSI: Zero-sale and No-upload ISDs with full details.

    Returns a list of message strings (one per TSI, in TSIS order).
    """
    hdr = target_date.strftime("%d %b %Y")

    def rows_for(cset):
        # Sort by ISD name for a stable, readable order; two lines per ISD.
        out = []
        for idx, c in enumerate(sorted(cset, key=lambda x: codes[x]["isd"].lower()), 1):
            i = codes[c]
            out.append("%d. <b>%s</b> · %s" % (idx, esc(i["isd"]) or "—",
                                               esc(catshort(i["cat"]))))
            out.append("    %s · %s" % (esc(i["store"]) or "—",
                                        esc(loc(i, short=True)) or "—"))
        return out

    msgs = []
    for t in TSIS:
        zr = rows_for(zero_codes[t])
        nu = rows_for(noup_codes[t])
        L = ["\U0001F4CB <b>%s</b> — ISD Exceptions" % esc(t),
             "%s <i>(D-1)</i>" % hdr,
             "<i>Name · Category / Store · Location</i>",
             "━" * 20, ""]
        L.append("\U0001F534 <b>ZERO SALES UPLOAD (%d)</b>" % (len(zr) // 2))
        L += (zr if zr else ["<i>None 🎉</i>"])
        L.append("")
        L.append("\U0001F7E0 <b>NO SALES UPLOAD (%d)</b>" % (len(nu) // 2))
        L += (nu if nu else ["<i>None 🎉</i>"])
        L.append("")
        L.append("<i>Week-off ISDs excluded.</i>")
        msgs.append("\n".join(L))
    return msgs

# ── Telegram ───────────────────────────────────────────────────────────────────

def send_telegram(text):
    if os.environ.get("DRY_RUN") == "1" or not (TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID):
        print("\n" + "=" * 60 + "\n" + text + "\n")
        return
    # Telegram hard limit is 4096 chars; split on blank lines if needed.
    chunks, cur = [], ""
    for para in text.split("\n\n"):
        add = (("\n\n" if cur else "") + para)
        if len(cur) + len(add) > 3800:
            chunks.append(cur); cur = para
        else:
            cur += add
    if cur:
        chunks.append(cur)
    url = "https://api.telegram.org/bot%s/sendMessage" % TELEGRAM_BOT_TOKEN
    for ch in chunks:
        body = urllib.parse.urlencode({
            "chat_id": TELEGRAM_CHAT_ID, "text": ch,
            "parse_mode": "HTML", "disable_web_page_preview": "true",
        }).encode()
        try:
            resp = urllib.request.urlopen(urllib.request.Request(url, data=body), timeout=30)
            print("[TG] sent chunk (%d chars): HTTP %d" % (len(ch), resp.getcode()))
        except urllib.error.HTTPError as e:
            print("[TG] send FAILED: HTTP %d %s" % (e.code, e.read().decode("utf-8", "replace")))
            raise

# ── Entry point ────────────────────────────────────────────────────────────────

def main():
    override = os.environ.get("TARGET_DATE", "").strip()
    if override:
        target_date = datetime.strptime(override, "%Y-%m-%d").date()
    else:
        target_date = (datetime.now(IST) - timedelta(days=1)).date()

    print("[TG] Summarising D-1 =", target_date)
    codes, tc, agg, zero_codes, noup_codes = compute(target_date)

    daily     = build_daily(target_date, codes, tc, agg, zero_codes, noup_codes)
    npi       = build_npi(target_date, codes, tc, agg)
    incentive = build_incentive(target_date, codes, tc, agg)
    isd_lists = build_isd_lists(target_date, codes, zero_codes, noup_codes)

    send_telegram(daily)
    send_telegram(npi)
    send_telegram(incentive)
    for msg in isd_lists:          # one message per TSI (4 total)
        send_telegram(msg)
    print("[TG] Done.")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""
fetch_competition_nural.py
--------------------------
Pulls the Nural **Competition Sales Report** for today (FTD) and sends a
TSI-wise Telegram message listing the East Reliance ISDs who have **NOT**
reported any competition sale today.

Runs 6 PM & 9 PM IST for today (FTD) and 9 AM IST for yesterday (D-1, via
DAY_OFFSET=1) through GitHub Actions (competition_report.yml).

Pipeline:
  1. Playwright logs into motoindia.nuralsales.app (same portal / login as the
     attendance scraper) only to read the session's authKey + userId from
     localStorage.
  2. Calls the report's own JSON API directly (no Excel export / report-queue):
       POST https://motowebapi.nuralsales.app/api/user/GetCompetitionSaleReport/<userId>
       header authkey:<authKey>, body {dateFrom,dateTo=today, pageIndex, pageSize,...}
     The response rows carry `ispCode`, so matching to the roster is exact.
  3. "Reported" = the ISD has >= 1 competition-sale row today. "Not reported" =
     every ACTIVE roster ISD not in that set, excluding anyone on Weekly-Off /
     Leave today (from the Attendance Tracker WeekOff Log) and Inactive ISDs.
  4. Group the not-reported ISDs by TSI and send one Telegram message to the
     same chat as the other reports (TELEGRAM_CHAT_ID).

Env:
  PORTAL_USER / PORTAL_PASS   Nural login (reused from the sales repo).
  TARGETS_CSV_URL             published Targets CSV (roster / TSI join).
  WEEKOFF_CSV_URL             published Attendance Tracker WeekOff Log CSV
                              (to exclude Weekly-Off / Leave ISDs for today).
  TELEGRAM_BOT_TOKEN / TELEGRAM_CHAT_ID
Optional:
  COMP_JSON        local JSON file (a list of report rows) -> skip Playwright.
  TODAY_OVERRIDE   YYYY-MM-DD -> anchor on this date instead of the wall clock.
  SEND_TELEGRAM=0  build the message but print instead of sending.
  HEADLESS=false   show the browser (local debugging).
"""

import os
import re
import json
import html
from datetime import datetime, timedelta, date, timezone

import fetch_weekoff_nural as WO        # proven Nural login / telegram / roster
import attendance_reports as AR         # parse_log / excluded_codes / _is_inactive

IST = timezone(timedelta(hours=5, minutes=30))
COMP_API = "https://motowebapi.nuralsales.app/api/user/GetCompetitionSaleReport/"
PAGE_SIZE = 2000

# Attendance Tracker · WeekOff Log tab (date-stamped attendance log) — same
# published CSV the Remarks pipeline reads. Env override wins.
WEEKOFF_CSV_URL = os.environ.get("WEEKOFF_CSV_URL") or \
    "https://docs.google.com/spreadsheets/d/1BsShTlrnwYGBUKMcURMI1Qla-pnttWFrPbMmiLNu6Eo/export?format=csv&gid=1706495387"


def env(n, d=""):
    return (os.environ.get(n) or d).strip()

def esc(s):
    return html.escape(str(s or ""), quote=True)

def _norm_name(s):
    """Lower-case, strip punctuation, collapse spaces — name-matching fallback."""
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9 ]", " ", str(s or "").lower())).strip()

# ── Nural competition data (JSON API) ────────────────────────────────────────

# Runs IN the page (same origin as the app, so the authkey header is accepted):
# fetch one page of the competition report for a single day.
_FETCH_JS = """
async ({ds, pageIndex, pageSize}) => {
  const uid = localStorage.getItem('userId');
  const ak  = localStorage.getItem('authKey');
  if (!uid || !ak) return {error: 'no-auth'};
  const res = await fetch('https://motowebapi.nuralsales.app/api/user/GetCompetitionSaleReport/' + uid, {
    method: 'POST',
    headers: {'Content-Type': 'application/json', 'authkey': ak},
    body: JSON.stringify({brandId:0, categoryId:0, modelId:0,
      dateFrom: ds, dateTo: ds, pageIndex, pageSize, ispId:0, retailerId:0})
  });
  const j = await res.json();
  return {total: parseInt(j.totalRecords || 0, 10), list: j.competitionSaleReportList || []};
}
"""

def comp_fetch(day):
    """Log in, then page the competition API for `day`. Returns list of rows."""
    from playwright.sync_api import sync_playwright
    user = env("PORTAL_USER") or env("NURAL_USERNAME")
    pw   = env("PORTAL_PASS") or env("NURAL_PASSWORD")
    if not (user and pw):
        raise RuntimeError("Set PORTAL_USER / PORTAL_PASS (Nural login).")
    ds = day.strftime("%Y-%m-%d")
    headless = env("HEADLESS", "true").lower() != "false"
    rows = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless, args=["--no-sandbox"])
        ctx = browser.new_context(timezone_id="Asia/Kolkata")
        page = ctx.new_page()
        try:
            WO._login(page, user, pw)
            # Land on the report page so the API call shares its origin.
            page.goto(WO.NURAL_ROOT + "/competition-sales-report",
                      wait_until="networkidle", timeout=60000)
            idx = 1
            while True:
                res = page.evaluate(_FETCH_JS, {"ds": ds, "pageIndex": idx, "pageSize": PAGE_SIZE})
                if not isinstance(res, dict) or res.get("error"):
                    raise RuntimeError("Competition API auth/read failed: %s" % res)
                lst = res.get("list") or []
                rows.extend(lst)
                total = int(res.get("total") or 0)
                print("[COMP] page %d: +%d (have %d / %d)" % (idx, len(lst), len(rows), total))
                if len(rows) >= total or not lst:
                    break
                idx += 1
        except Exception:
            try:
                (open("_comp_debug.html", "w", encoding="utf-8")).write(page.content())
                page.screenshot(path="_comp_debug.png", full_page=True)
            except Exception:
                pass
            raise
        finally:
            ctx.close(); browser.close()
    return rows

# ── reporters ────────────────────────────────────────────────────────────────

def reporters_for(day, rows, codes):
    """Set of roster ISP Codes that reported >=1 competition sale on `day`.
    Prefers the ispCode field; falls back to normalised ISP Name."""
    if not rows:
        return set(), "none"
    r0 = rows[0]
    k_code = WO._key(r0, "ispCode", "ISP Code", "ISPCode", "EntityCode")
    k_name = WO._key(r0, "ispName", "ISPNAME", "ISP Name")
    k_date = WO._key(r0, "saleDate", "Sale Date", "Date")

    name_to_codes = {}
    for c, info in codes.items():
        nm = _norm_name(info.get("isd"))
        if nm:
            name_to_codes.setdefault(nm, set()).add(c)

    reported = set()
    for r in rows:
        if k_date:
            d = WO.parse_any_date(r.get(k_date))
            if d and d != day:
                continue
        if k_code and str(r.get(k_code) or "").strip() in codes:
            reported.add(str(r.get(k_code)).strip())
        elif k_name:
            for c in name_to_codes.get(_norm_name(r.get(k_name)), ()):
                reported.add(c)
    return reported, ("code" if k_code else "name")

# ── Weekly-Off / Leave exclusion for today ───────────────────────────────────

def excluded_today(day):
    """ISP Codes on Weekly-Off / Leave on `day`, from the Attendance Tracker
    WeekOff Log. Best-effort — returns empty set on any failure."""
    if not WEEKOFF_CSV_URL:
        return set()
    try:
        rows = WO._fetch_csv(WEEKOFF_CSV_URL)
        log = AR.parse_log(rows)
        excl = AR.excluded_codes(log, day)
        print("[COMP] WeekOff/Leave excluded today:", len(excl))
        return excl
    except Exception as e:
        print("[COMP] WeekOff fetch failed (non-fatal):", str(e)[:120])
        return set()

# ── Telegram message ─────────────────────────────────────────────────────────

def build_messages(not_reported, codes, label, subtitle, match_mode):
    """One self-contained Telegram message per TSI (TSIS order, then any extras).
    A TSI with none still gets a short 'All reported' message, mirroring the
    attendance 'Not Marked' report."""
    by_tsi = {}
    for c in not_reported:
        t = codes[c].get("tsi") or "—"
        by_tsi.setdefault(t, []).append(c)
    for t in by_tsi:
        by_tsi[t].sort(key=lambda c: (codes[c].get("isd") or c).lower())

    order = list(WO.TSIS) + [t for t in by_tsi if t not in WO.TSIS]
    msgs = []
    for t in order:
        lst = by_tsi.get(t, [])
        L = ["🏁 <b>%s — Competition NOT Reported · %s</b>" % (esc(t), label),
             "<i>%s</i>" % subtitle,
             "<i>ISD · Store</i>", "━" * 20, ""]
        if lst:
            for i, c in enumerate(lst, 1):
                info = codes[c]
                line = "%d. <b>%s</b>" % (i, esc(info.get("isd") or c))
                if info.get("store"):
                    line += " — %s" % esc(info["store"])
                L.append(line)
        else:
            L.append("<i>All reported 🎉</i>")
        L.append("")
        L.append("<i>Count: %d</i>" % len(lst))
        if match_mode == "name":
            L.append("<i>Matched by ISD name — verify any unexpected name.</i>")
        msgs.append("\n".join(L))
    return msgs

# ── main ─────────────────────────────────────────────────────────────────────

def main():
    now_ist = datetime.now(IST)
    offset  = int(env("DAY_OFFSET", "0") or "0")   # 0 = today (FTD), 1 = yesterday (D-1)
    _ovr    = env("TODAY_OVERRIDE")
    day = (datetime.strptime(_ovr, "%Y-%m-%d").date() if _ovr
           else now_ist.date() - timedelta(days=offset))
    label = "D-%d" % offset if offset > 0 else "FTD"
    print("[COMP] report:", label, "| date:", day)

    codes = WO.load_codes()
    print("[COMP] roster ISDs:", len(codes))

    comp_json = env("COMP_JSON")
    if comp_json:
        with open(comp_json, encoding="utf-8") as fh:
            rows = json.load(fh)
    else:
        rows = comp_fetch(day)
    print("[COMP] competition rows:", len(rows))

    reported, mode = reporters_for(day, rows, codes)
    print("[COMP] reported today:", len(reported), "| match mode:", mode)

    excl = excluded_today(day)

    not_reported = [c for c, info in codes.items()
                    if not AR._is_inactive(info)
                    and c not in excl
                    and c not in reported]
    print("[COMP] NOT reported:", len(not_reported))

    day_s = "%02d %s %d" % (day.day, day.strftime("%b"), day.year)
    subtitle = (day_s if offset > 0
                else "%s · as of %s IST" % (day_s, now_ist.strftime("%I:%M %p").lstrip("0")))
    msgs = build_messages(not_reported, codes, label, subtitle, mode)
    for m in msgs:
        WO.send_telegram(m)
    print("[COMP] sent %d TSI message(s)." % len(msgs))
    print("[COMP] done.")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""
fetch_weekoff_nural.py
----------------------
Pulls Week-Off / Leave / attendance from Nural and writes a date-stamped log
into the Week-Off tab of the Targets spreadsheet, then (optionally) sends the
daily + MTD attendance reports to Telegram.

Runs 4 PM & 11:59 PM IST via GitHub Actions (fetch_weekoff_nural.yml).

Pipeline:
  1. Playwright logs into motoindia.nuralsales.app (same portal as main.py),
     exports the Attendance + Leave reports for the window (current + prev
     month) as Excel — reusing the proven flow from the Visit Tracker scraper
     (MUI calendar date-pick + exceldownload icon).
  2. Parse both workbooks, filter to East Reliance by USER CODE in Targets.
  3. Derive per ISD per date: Weekly Off / Leave(+type) / Present + Late(>12:30).
  4. Overwrite the Week-Off tab with the log  ->  ISP Code | Date | Status |
     Leave Type | In Time | Late.
  5. Build + send the daily and MTD reports (attendance_reports.py).

Env:
  PORTAL_USER / PORTAL_PASS            Nural login (reused from the sales repo).
    (NURAL_USERNAME / NURAL_PASSWORD also accepted.)
  TARGETS_SPREADSHEET_ID               file id of the spreadsheet holding the
                                       Week-Off tab (the /d/<ID>/edit id).
  WEEKOFF_GID            default 1522547784 — the Week-Off tab to overwrite.
  GOOGLE_CLIENT_ID / GOOGLE_CLIENT_SECRET / GOOGLE_REFRESH_TOKEN
  TARGETS_CSV_URL       published Targets CSV (read-only, for the ISD/TSI join).
  TELEGRAM_BOT_TOKEN / TELEGRAM_CHAT_ID
Optional:
  ATT_XLSX / LEAVE_XLSX  local .xlsx paths -> skip Playwright, parse these
                         (offline validation of the parser against real exports).
  SEND_TELEGRAM=0        build reports but don't send.
  HEADLESS=false         show the browser (local debugging).
  MONTHS_BACK            default 1 (current month + N previous months).
"""

import os
import re
import csv
import io
import urllib.request
from datetime import datetime, timedelta, date, timezone
from pathlib import Path

import attendance_reports as AR

IST = timezone(timedelta(hours=5, minutes=30))
WEEKOFF_GID = int(os.environ.get("WEEKOFF_GID", "1522547784"))
# App-owned output sheet (the pipeline's drive.file-scoped token can only touch
# files it creates, so attendance is written to a dedicated app-owned sheet).
ATT_SHEET_NAME = os.environ.get("ATTENDANCE_SHEET_NAME") or "Attendance Tracker - East Reliance"
ATT_FOLDER_ID  = os.environ.get("ATTENDANCE_FOLDER_ID") or "1a_k-hkkvOgPQWASIhCTvesNSytx4qfTT"
REMAP = {"Saroj Kumar Rana": "Amit Kumar", "Uday Kumar Tiwari": "Mukesh Kumar"}
TSIS = ["Amit Kumar", "Joydwip Seal", "Mukesh Kumar", "Nabarun Bhattacharya"]

TARGETS_CSV_URL = os.environ.get("TARGETS_CSV_URL") or \
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vScO4tHpOg-uISI7WnhedipMDO8eVNgsX0UC5AWNNe1eDbAhq2U7ZiaWpL51tvDvAOsVjD3cuwlt2Oz/pub?output=csv"
NURAL_ROOT = "https://motoindia.nuralsales.app"
ATT_URL   = NURAL_ROOT + "/view-attendance-report"
LEAVE_URL = NURAL_ROOT + "/leave-report"


def env(n, d=""):
    return (os.environ.get(n) or d).strip()

# ── date helpers ─────────────────────────────────────────────────────────────

def parse_any_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v or "").strip()
    if not s:
        return None
    for fmt in ("%d-%b-%Y", "%d/%b/%Y", "%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d",
                "%d-%b-%y", "%d/%m/%y", "%d-%B-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None

def window_bounds(today=None, months_back=1):
    today = today or datetime.now(IST).date()
    m, y = today.month - months_back, today.year
    while m < 1:
        m += 12; y -= 1
    return date(y, m, 1), today

# ── Targets join (read-only published CSV) ───────────────────────────────────

def _fetch_csv(url):
    req = urllib.request.Request(url + ("&" if "?" in url else "?") + "_t=" +
                                 str(int(datetime.now().timestamp())),
                                 headers={"User-Agent": "moto-weekoff"})
    text = urllib.request.urlopen(req, timeout=60).read().decode("utf-8", "replace")
    if text[:1] == "﻿":
        text = text[1:]
    lines = [l for l in text.splitlines() if l.strip().strip(",")]
    for i in range(min(6, len(lines))):
        if re.search(r"isp\s*code", lines[i], re.I):
            lines = lines[i:]; break
    return list(csv.DictReader(lines))

def _key(row, *cands):
    m = {re.sub(r"[\s_]+", "", k).lower(): k for k in row}
    for c in cands:
        kk = re.sub(r"[\s_]+", "", c).lower()
        if kk in m:
            return m[kk]
    return None

def load_codes():
    rows = _fetch_csv(TARGETS_CSV_URL)
    r0 = rows[0]
    kC = _key(r0, "ISP Code", "ISPCode")
    kN = _key(r0, "ISP Name", "ISPName")
    kT = _key(r0, "TSI Name", "TSIName")
    kS = _key(r0, "Parent Retailer Name", "Store Name")
    kCt = _key(r0, "City"); kSt = _key(r0, "State")
    kSt2 = _key(r0, "ISP Status", "Status")
    codes = {}
    for r in rows:
        c = (r.get(kC) or "").strip()
        if not c or c in codes:
            continue
        tsi = (r.get(kT) or "").strip()
        codes[c] = {
            "tsi": REMAP.get(tsi, tsi),
            "isd": (r.get(kN) or "").strip(),
            "store": (r.get(kS) or "").strip() if kS else "",
            "city": (r.get(kCt) or "").strip() if kCt else "",
            "state": (r.get(kSt) or "").strip() if kSt else "",
            "status": (r.get(kSt2) or "").strip() if kSt2 else "",
        }
    return codes

# ── Nural download (Playwright) — adapted from scrape_visit_report.py ─────────

CAL_ICON  = 'svg[data-testid="CalendarTodayIcon"]'
CAL_MODAL = '.MuiModal-root:has(svg[data-testid="ChevronRightIcon"])'
_MON = {m.lower(): i for i, m in enumerate(
    ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"], 1)}

def _login(page, user, pw):
    last = None
    for attempt in range(3):
        try:
            page.goto(NURAL_ROOT + "/login", wait_until="domcontentloaded", timeout=60000)
            page.get_by_placeholder("Enter Username", exact=True).fill(user)
            page.get_by_placeholder("Enter Password", exact=True).fill(pw)
            page.get_by_role("button", name="LOGIN", exact=True).click()
            page.wait_for_function(
                "() => !location.pathname.toLowerCase().includes('/login')", timeout=30000)
            page.wait_for_load_state("networkidle", timeout=60000)
            return
        except Exception as e:
            last = e
            page.wait_for_timeout(1500 * (attempt + 1))
    raise RuntimeError("Nural login failed after 3 attempts: %s" % last)

def _cal_month_year(cal):
    texts = [t.strip() for t in cal.locator("p").all_inner_texts() if t.strip()]
    m = re.search(r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[,\s]+(\d{4})",
                  " ".join(texts), re.I)
    if not m:
        raise RuntimeError("Cannot read calendar month/year: %s" % texts)
    return _MON[m.group(1).lower()[:3]], int(m.group(2))

def _set_date(page, idx, dt):
    icons = page.locator(CAL_ICON)
    box = icons.nth(idx).bounding_box()
    page.mouse.click(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
    cal = page.locator(CAL_MODAL).last
    cal.wait_for(state="visible", timeout=15000)
    for _ in range(24):
        mo, yr = _cal_month_year(cal)
        if yr * 12 + mo == dt.year * 12 + dt.month:
            break
        d = "ChevronRightIcon" if yr * 12 + mo < dt.year * 12 + dt.month else "ChevronLeftIcon"
        cal.locator('button:has(svg[data-testid="%s"])' % d).first.click()
        page.wait_for_timeout(350)
    cal.get_by_text(str(dt.day), exact=True).first.click()
    page.wait_for_timeout(500)

def _download_from_queue(page, since, out_path, keyword):
    """Async fallback: the export was queued -> poll /report-queue for a fresh
    successful row and download it (mirrors scrape_visit_report.py)."""
    cutoff = since - timedelta(minutes=3)
    deadline = datetime.now(IST) + timedelta(minutes=12)
    page.goto(NURAL_ROOT + "/report-queue", wait_until="networkidle", timeout=60000)
    while datetime.now(IST) < deadline:
        rows = page.locator("tr, .MuiDataGrid-row, [role='row']")
        for idx in range(min(rows.count(), 40)):
            row = rows.nth(idx)
            try:
                txt = row.inner_text(timeout=2000)
            except Exception:
                continue
            if keyword and keyword.lower() not in txt.lower():
                continue
            if not re.search(r"process\s+success|successfully|completed|complete", txt, re.I):
                continue
            with page.expect_download(timeout=120000) as dl:
                row.locator('button[title="Download Report"]').click()
            dl.value.save_as(out_path)
            print("[WO] downloaded(queue)", out_path)
            return
        page.wait_for_timeout(15000)
        page.reload(wait_until="networkidle", timeout=60000)
    raise RuntimeError("Nural queue produced no fresh '%s' report in time." % keyword)

def _export_report(page, report_url, from_dt, to_dt, out_path, queue_keyword):
    page.goto(report_url, wait_until="networkidle", timeout=60000)
    if "/login" in page.url.lower():
        raise RuntimeError("Redirected to /login opening %s" % report_url)
    page.wait_for_selector(CAL_ICON, state="visible", timeout=45000)
    _set_date(page, 0, from_dt)
    _set_date(page, 1, to_dt)

    direct, downloads = [], []
    page.on("response", lambda r: r.url.lower().endswith(".xlsx") and direct.append(r))
    page.on("download", lambda d: downloads.append(d))
    started = datetime.now(IST)
    # SEARCH button carries an icon, so match by visible text, not accessible name.
    btn = page.locator("button").filter(has_text=re.compile(r"^\s*SEARCH\s*$")).first
    btn.scroll_into_view_if_needed()
    btn.click()
    deadline = datetime.now(IST) + timedelta(seconds=90)
    while datetime.now(IST) < deadline:
        if downloads:
            downloads[0].save_as(out_path); print("[WO] downloaded", out_path); return
        if direct:
            body = direct[0].body()
            if len(body) >= 1024:
                Path(out_path).write_bytes(body); print("[WO] downloaded(direct)", out_path); return
        if page.get_by_text(re.compile(r"TOTAL RECORDS|USER CODE|LEAVE TYPE", re.I)).count():
            icon = page.locator('img[src*="exceldownload"]').first
            icon.wait_for(state="attached", timeout=30000)
            with page.expect_download(timeout=180000) as dl:
                icon.evaluate("el => el.click()")
            dl.value.save_as(out_path); print("[WO] downloaded", out_path); return
        if page.get_by_text(re.compile(r"status of your request|report queue|added to queue|submitted successfully", re.I)).count():
            break
        page.wait_for_timeout(1000)
    _download_from_queue(page, started, out_path, queue_keyword)

def nural_download(from_dt, to_dt, out_dir):
    from playwright.sync_api import sync_playwright
    user = env("PORTAL_USER") or env("NURAL_USERNAME")
    pw   = env("PORTAL_PASS") or env("NURAL_PASSWORD")
    if not (user and pw):
        raise RuntimeError("Set PORTAL_USER / PORTAL_PASS (Nural login).")
    out_dir = Path(out_dir); out_dir.mkdir(parents=True, exist_ok=True)
    att = out_dir / "attendance.xlsx"
    lv  = out_dir / "leave.xlsx"
    headless = env("HEADLESS", "true").lower() != "false"
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless, args=["--no-sandbox"])
        ctx = browser.new_context(accept_downloads=True, timezone_id="Asia/Kolkata")
        page = ctx.new_page()
        try:
            _login(page, user, pw)
            _export_report(page, ATT_URL,   from_dt, to_dt, str(att), "attendance")
            _export_report(page, LEAVE_URL, from_dt, to_dt, str(lv),  "leave")
        except Exception:
            try:
                (out_dir / "debug.html").write_text(page.content(), encoding="utf-8")
                page.screenshot(path=str(out_dir / "debug.png"), full_page=True)
            except Exception:
                pass
            raise
        finally:
            ctx.close(); browser.close()
    return att, lv

# ── xlsx parsing (header-tolerant) ───────────────────────────────────────────

def _read_xlsx(path):
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = [[c for c in r] for r in ws.iter_rows(values_only=True)]
    wb.close()
    # Find the header row (contains the user/entity code column).
    hdr_i = 0
    for i, r in enumerate(rows[:10]):
        if any(re.sub(r"\s+", "", str(c or "")).lower() in ("usercode", "entitycode")
               for c in r):
            hdr_i = i; break
    header = [str(c or "").strip() for c in rows[hdr_i]]
    out = []
    for r in rows[hdr_i + 1:]:
        if not any(str(c or "").strip() for c in r):
            continue
        out.append({header[j]: r[j] for j in range(min(len(header), len(r)))})
    return out

def _hk(row, *cands):
    m = {re.sub(r"[\s_()\-]+", "", k).lower(): k for k in row}
    for c in cands:
        kk = re.sub(r"[\s_()\-]+", "", c).lower()
        if kk in m:
            return m[kk]
    return None

def parse_attendance(path, codes, lo, hi):
    """Real export cols: EntityCode, AttendanceDate, AttendanceStatus,
    'First visit in-time' / 'Attendence time' (in-time at store)."""
    rows = _read_xlsx(path)
    if not rows:
        return {}
    r0 = rows[0]
    kC = _hk(r0, "EntityCode", "USER CODE")
    kD = _hk(r0, "AttendanceDate", "DATE")
    kS = _hk(r0, "AttendanceStatus", "ATTENDANCE STATUS", "STATUS")
    kI = _hk(r0, "First visit in-time", "Attendence time", "Attendance time", "IN TIME")
    log = {}
    for r in rows:
        code = str(r.get(kC) or "").strip()
        d = parse_any_date(r.get(kD))
        if code not in codes or not d or not (lo <= d <= hi):
            continue
        status = re.sub(r"\s+", " ", str(r.get(kS) or "").strip())
        if "week" in status.lower() and "off" in status.lower():
            status = "Weekly Off"
        in_time = str(r.get(kI) or "").strip()
        log[(code, d)] = {
            "status": status, "leave_type": "",
            "in_time": in_time,
            "late": status.lower() == "present" and AR.is_late(in_time),
        }
    return log

def parse_leave(path, codes, lo, hi):
    """Real export cols: EntityCode, LeaveFromDate, LeaveToDate, LeaveTypeName,
    ApprovalStatus. Rejected leaves are ignored; Pending + Approved count."""
    rows = _read_xlsx(path)
    if not rows:
        return {}
    r0 = rows[0]
    kC = _hk(r0, "EntityCode", "USER CODE")
    kT = _hk(r0, "LeaveTypeName", "LEAVE TYPE")
    kF = _hk(r0, "LeaveFromDate", "FROM")
    kTo = _hk(r0, "LeaveToDate", "TO")
    kA = _hk(r0, "ApprovalStatus")
    out = {}
    for r in rows:
        code = str(r.get(kC) or "").strip()
        if code not in codes:
            continue
        appr = str(r.get(kA) or "").strip().lower() if kA else ""
        if appr == "rejected":
            continue
        f = parse_any_date(r.get(kF)); t = parse_any_date(r.get(kTo))
        if not f:
            continue
        t = t or f
        lt = str(r.get(kT) or "Leave").strip()
        d = max(f, lo)
        while d <= min(t, hi):
            out[(code, d)] = {"status": "Leave", "leave_type": lt,
                              "in_time": "", "late": False}
            d += timedelta(days=1)
    return out

# ── Google Sheets write ──────────────────────────────────────────────────────

def get_creds():
    from google.oauth2.credentials import Credentials
    return Credentials(
        token=None,
        refresh_token=os.environ["GOOGLE_REFRESH_TOKEN"],
        token_uri="https://oauth2.googleapis.com/token",
        client_id=os.environ["GOOGLE_CLIENT_ID"],
        client_secret=os.environ["GOOGLE_CLIENT_SECRET"],
    )

def _tab_title(svc, sid, gid):
    meta = svc.spreadsheets().get(spreadsheetId=sid).execute()
    for sh in meta["sheets"]:
        if sh["properties"]["sheetId"] == gid:
            return sh["properties"]["title"]
    raise RuntimeError("No tab with gid %s in spreadsheet %s" % (gid, sid))

def overwrite_tab(svc, sid, title, values):
    rng = "'%s'!A1" % title
    svc.spreadsheets().values().clear(spreadsheetId=sid, range="'%s'" % title).execute()
    svc.spreadsheets().values().update(
        spreadsheetId=sid, range=rng, valueInputOption="RAW",
        body={"values": values}).execute()
    print("[WO] wrote %d rows to tab '%s'" % (len(values), title))

def ensure_tab(svc, sid, title):
    meta = svc.spreadsheets().get(spreadsheetId=sid).execute()
    if not any(sh["properties"]["title"] == title for sh in meta["sheets"]):
        svc.spreadsheets().batchUpdate(spreadsheetId=sid, body={
            "requests": [{"addSheet": {"properties": {"title": title}}}]}).execute()

def _sheet_id(svc, sid, title):
    meta = svc.spreadsheets().get(spreadsheetId=sid).execute()
    for sh in meta["sheets"]:
        if sh["properties"]["title"] == title:
            return sh["properties"]["sheetId"]
    return None

def _rgb(hexs):
    h = hexs.lstrip("#")
    return {"red": int(h[0:2], 16) / 255, "green": int(h[2:4], 16) / 255,
            "blue": int(h[4:6], 16) / 255}

# Day-cell status text -> background colour for the MTD matrix.
_MTD_COLORS = [
    ("Present",     "TEXT_EQ",       "#B7E1CD"),
    ("Late",        "TEXT_EQ",       "#F6B26B"),
    ("Weekly Off",  "TEXT_EQ",       "#FFE599"),
    ("Leave",       "TEXT_EQ",       "#EA9999"),
    ("Absent",      "TEXT_EQ",       "#E06666"),
    ("Store Close", "TEXT_EQ",       "#D9D9D9"),
    ("Pending",     "TEXT_CONTAINS", "#CFE2F3"),
]

def format_mtd_tab(svc, sid, title, lead_cols, day_cols, data_rows):
    """Colour the MTD day grid by status text + freeze header rows / label cols.
    Programmatic so the colours track the day columns after the count columns are
    inserted. Clears any prior conditional-format rules first (idempotent)."""
    sheet_id = _sheet_id(svc, sid, title)
    if sheet_id is None:
        return
    meta = svc.spreadsheets().get(
        spreadsheetId=sid,
        fields="sheets(properties(sheetId),conditionalFormats)").execute()
    existing = 0
    for sh in meta["sheets"]:
        if sh["properties"]["sheetId"] == sheet_id:
            existing = len(sh.get("conditionalFormats", []) or [])
            break
    reqs = [{"deleteConditionalFormatRule": {"sheetId": sheet_id, "index": i}}
            for i in range(existing - 1, -1, -1)]
    day_range = {"sheetId": sheet_id,
                 "startRowIndex": 3, "endRowIndex": 3 + data_rows,
                 "startColumnIndex": lead_cols, "endColumnIndex": lead_cols + day_cols}
    for text, ctype, color in _MTD_COLORS:
        reqs.append({"addConditionalFormatRule": {"index": 0, "rule": {
            "ranges": [day_range],
            "booleanRule": {
                "condition": {"type": ctype, "values": [{"userEnteredValue": text}]},
                "format": {"backgroundColor": _rgb(color)}}}}})
    reqs.append({"updateSheetProperties": {
        "properties": {"sheetId": sheet_id,
                       "gridProperties": {"frozenRowCount": 3,
                                          "frozenColumnCount": lead_cols}},
        "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"}})
    reqs.append({"repeatCell": {
        "range": {"sheetId": sheet_id, "startRowIndex": 2, "endRowIndex": 3},
        "cell": {"userEnteredFormat": {"textFormat": {"bold": True}}},
        "fields": "userEnteredFormat.textFormat.bold"}})
    svc.spreadsheets().batchUpdate(spreadsheetId=sid, body={"requests": reqs}).execute()
    print("[WO] formatted MTD tab '%s' (%d day cols, %d data rows)"
          % (title, day_cols, data_rows))

def get_or_create_attendance_sheet(drv, svc):
    """Find the app-owned attendance spreadsheet by name (drive.file sees only
    app-created files), or create it in the shared folder and make it
    link-readable so the dashboard / consumers can read it as CSV."""
    q = ("name='%s' and mimeType='application/vnd.google-apps.spreadsheet' "
         "and trashed=false" % ATT_SHEET_NAME.replace("'", "\\'"))
    res = drv.files().list(q=q, fields="files(id,name)", pageSize=5,
                           includeItemsFromAllDrives=True, supportsAllDrives=True,
                           corpora="allDrives").execute()
    if res.get("files"):
        return res["files"][0]["id"]
    sid = svc.spreadsheets().create(
        body={"properties": {"title": ATT_SHEET_NAME}}).execute()["spreadsheetId"]
    if ATT_FOLDER_ID:
        try:
            drv.files().update(fileId=sid, addParents=ATT_FOLDER_ID, removeParents="root",
                               supportsAllDrives=True, fields="id,parents").execute()
        except Exception as e:
            print("[WO] move-to-folder skipped (non-fatal):", str(e)[:120])
    try:
        drv.permissions().create(fileId=sid, body={"type": "anyone", "role": "reader"},
                                 supportsAllDrives=True).execute()
    except Exception as e:
        print("[WO] public-read skipped (non-fatal):", str(e)[:120])
    print("[WO] created attendance sheet", sid)
    return sid

def log_to_rows(log):
    head = ["ISP Code", "Date", "Status", "Leave Type", "In Time", "Late"]
    rows = []
    for (c, d), r in sorted(log.items(), key=lambda kv: (kv[0][1], kv[0][0])):
        rows.append([c, d.strftime("%Y-%m-%d"), r["status"], r["leave_type"],
                     r["in_time"], "Y" if r["late"] else ""])
    return [head] + rows

# ── Telegram ─────────────────────────────────────────────────────────────────

def send_telegram(text):
    import urllib.parse
    tok = env("TELEGRAM_BOT_TOKEN"); chat = env("TELEGRAM_CHAT_ID")
    if env("SEND_TELEGRAM", "1") == "0" or not (tok and chat):
        print("\n" + "=" * 60 + "\n" + text + "\n")
        return
    url = "https://api.telegram.org/bot%s/sendMessage" % tok
    for chunk in _split(text, 3800):
        body = urllib.parse.urlencode({"chat_id": chat, "text": chunk,
            "parse_mode": "HTML", "disable_web_page_preview": "true"}).encode()
        urllib.request.urlopen(urllib.request.Request(url, data=body), timeout=30)

def _split(text, limit):
    """Split into <=limit chunks by lines (long lists have no blank lines, so a
    paragraph splitter would overflow Telegram's 4096 cap)."""
    chunks, cur = [], ""
    for line in text.split("\n"):
        add = (("\n" if cur else "") + line)
        if cur and len(cur) + len(add) > limit:
            chunks.append(cur); cur = line
        else:
            cur += add
        while len(cur) > limit:            # safety for a single over-long line
            chunks.append(cur[:limit]); cur = cur[limit:]
    if cur:
        chunks.append(cur)
    return chunks

# ── main ─────────────────────────────────────────────────────────────────────

def main():
    # TODAY_OVERRIDE (YYYY-MM-DD) lets a one-off/offline run anchor on the data's
    # date instead of the wall clock (e.g. replaying a D-1 export).
    _ovr = env("TODAY_OVERRIDE")
    today = datetime.strptime(_ovr, "%Y-%m-%d").date() if _ovr else datetime.now(IST).date()
    months_back = int(env("MONTHS_BACK", "1"))
    lo, hi = window_bounds(today, months_back)
    print("[WO] window", lo, "->", hi)

    codes = load_codes()
    print("[WO] Targets ISDs:", len(codes))

    att_xlsx, leave_xlsx = env("ATT_XLSX"), env("LEAVE_XLSX")
    if att_xlsx and leave_xlsx:
        print("[WO] offline parse mode")
        att_path, lv_path = Path(att_xlsx), Path(leave_xlsx)
    else:
        att_path, lv_path = nural_download(lo, hi, "_wo_dl")

    log = {}
    log.update(parse_attendance(att_path, codes, lo, hi))
    log.update(parse_leave(lv_path, codes, lo, hi))   # leave overrides same-day
    print("[WO] log entries:", len(log))

    # Write to the app-owned attendance sheet (created/opened in the shared
    # folder). A drive.file-scoped token can only touch files it created.
    from googleapiclient.discovery import build
    creds = get_creds()
    svc = build("sheets", "v4", credentials=creds)
    drv = build("drive", "v3", credentials=creds)
    sid = get_or_create_attendance_sheet(drv, svc)
    ensure_tab(svc, sid, "WeekOff Log")
    overwrite_tab(svc, sid, "WeekOff Log", log_to_rows(log))
    ensure_tab(svc, sid, "Attendance FTD")
    overwrite_tab(svc, sid, "Attendance FTD", AR.attendance_ftd_rows(hi, codes, TSIS, log))
    mtd_tab = "Attendance MTD %s" % hi.strftime("%b %Y")
    ensure_tab(svc, sid, mtd_tab)
    matrix = AR.attendance_matrix_rows(hi.year, hi.month, hi.day, codes, TSIS, log)
    overwrite_tab(svc, sid, mtd_tab, matrix)
    format_mtd_tab(svc, sid, mtd_tab, AR.MATRIX_LEAD_COLS, hi.day, len(matrix) - 3)
    print("[WO] attendance sheet: https://docs.google.com/spreadsheets/d/%s/edit" % sid)

    # Telegram:
    #  - FTD + late daily report (every run),
    #  - one "Not Marked" message per TSI (every run),
    #  - MTD summary only on the end-of-day finalise run (IST hour >= 22) or when
    #    SEND_MTD=1 is forced; SEND_MTD=0 suppresses it entirely.
    send_telegram(AR.build_daily_report(hi, codes, TSIS, log))
    for msg in AR.build_notmarked_messages(hi, codes, TSIS, log):
        send_telegram(msg)
    send_mtd = env("SEND_MTD", "")
    if send_mtd == "1" or (send_mtd != "0" and datetime.now(IST).hour >= 22):
        send_telegram(AR.build_mtd_report(hi.year, hi.month, codes, TSIS, log))
    print("[WO] done.")


if __name__ == "__main__":
    main()

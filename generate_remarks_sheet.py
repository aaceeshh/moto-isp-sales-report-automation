#!/usr/bin/env python3
"""
generate_remarks_sheet.py
Reads yesterday's ISP sales data from the published Sales CSV, filters
Zero Sale Uploaded + Sale Not Uploaded ISDs, and writes a formatted tab
(named by date, e.g. "21st May") into a Google Sheets file called
"Remarks Required for Zero & No Sales Upload" in the shared Drive folder.

Runs at 02:00 AM IST — always processes the previous calendar day.
Skip-if-exists: if a tab for that date already exists it is left untouched,
so any remarks TSIs have already entered are preserved (no overwrite).

Column I (Remarks) is the only editable column for TSI editors.
All other columns are locked via protected ranges (owner-only).

Required env vars:
  GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, GOOGLE_REFRESH_TOKEN
  DRIVE_FOLDER_ID   -- Drive folder id (same as main pipeline)
  SALES_CSV_URL     -- Published CSV of the Sales Google Sheet
  TARGETS_CSV_URL   -- Published CSV of the Targets Sheet
"""

import csv
import os
import re
from datetime import datetime, timezone, timedelta

import requests
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

# ── Timezone ──────────────────────────────────────────────────────────────────

IST = timezone(timedelta(hours=5, minutes=30))

# ── Config ────────────────────────────────────────────────────────────────────

REMARKS_FILE_NAME = "Remarks Required for Zero & No Sales Upload"
SALES_CSV_URL     = os.environ.get("SALES_CSV_URL",   "")
TARGETS_CSV_URL   = os.environ.get("TARGETS_CSV_URL", "")
WEEKOFF_CSV_URL   = os.environ.get("WEEKOFF_CSV_URL",  "")

# TSI email addresses that get editor access to column I (Remarks) only.
# Add / remove addresses here as your team changes.
# Leave empty to skip sharing (you can share the file manually instead).
TSI_EDITOR_EMAILS = [
    "nabarun.bhattacharya@stmpl.co.in",
    "radheshyam.mishra@stmpl.co.in",
    "kumar.mukesh@stmpl.co.in",
    "uday.tiwari@stmpl.co.in",
    "sujit.rajak@stmpl.co.in",
    "madhurjya.tamuly@stmpl.co.in",
    "rajoo.thakur@stmpl.co.in",
    "saroj.rana@stmpl.co.in",
]

# ── Column definition ─────────────────────────────────────────────────────────

COLS = [
    "TSI Name", "ISP Code", "ISD Name", "Store Code",
    "Retailer / Store", "City", "State", "Category", "Remarks",
]
N = len(COLS)

# ── Colour palette ────────────────────────────────────────────────────────────

def _c(r, g, b):
    return {"red": r/255, "green": g/255, "blue": b/255}

C_MOTO_RED   = _c(225, 20,  10)
C_ORANGE_SEC = _c(230, 81,   0)
C_RED_SEC    = _c(183, 28,  28)
C_COL_HDR    = _c(245, 245, 245)
C_ORANGE_ALT = _c(255, 243, 224)
C_RED_ALT    = _c(255, 235, 238)
C_WHITE      = _c(255, 255, 255)
C_REMARKS    = _c(255, 253, 231)
C_TSI_HDR    = _c(197, 202, 233)
C_GRAY_TEXT  = _c(66,  66,  66)
C_WHITE_TEXT = _c(255, 255, 255)

# ── Ordinal / tab name ────────────────────────────────────────────────────────

def _ordinal(n):
    suffix = ("th" if 11 <= (n % 100) <= 13
              else {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th"))
    return str(n) + suffix

def tab_name_for(dt):
    return _ordinal(dt.day) + " " + dt.strftime("%b")   # e.g. "21st May"

# ── CSV helpers ───────────────────────────────────────────────────────────────

def _norm(s):
    return re.sub(r"[\s_\-]+", "", str(s)).lower()

def _find_key(row_dict, *candidates):
    norm_map = {_norm(k): k for k in row_dict}
    for c in candidates:
        m = norm_map.get(_norm(c))
        if m:
            return m
    return None

def _fetch_csv(url):
    resp = requests.get(url + "&_t=" + str(int(datetime.now().timestamp())), timeout=30)
    resp.raise_for_status()
    text = resp.text.lstrip("﻿")
    lines = [l for l in text.splitlines() if l.strip().strip(",")]
    if not lines:
        return []
    # Some sheets carry a leading totals/spacer row above the real header
    # (e.g. ",,,,,,,,,,,,1457.4,1350.9,..." in the Targets sheet). That row
    # has values so the blank-row filter above keeps it. All feeds (Sales,
    # Targets, WeekOff) carry an "ISP Code"/"ISPCode" column, so anchor on
    # that to locate the true header row and drop anything above it.
    for idx in range(min(5, len(lines))):
        if re.search(r"isp\s*code", lines[idx], re.IGNORECASE):
            lines = lines[idx:]
            break
    return list(csv.DictReader(lines))

def _parse_date(s):
    for fmt in ("%d-%b-%y", "%d-%b-%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(str(s).strip(), fmt).date()
        except ValueError:
            pass
    return None

# ── Data loading ──────────────────────────────────────────────────────────────

def load_targets():
    if not TARGETS_CSV_URL:
        return {}
    try:
        rows = _fetch_csv(TARGETS_CSV_URL)
    except Exception as e:
        print("[Remarks] Targets fetch failed (non-fatal): " + str(e))
        return {}
    if not rows:
        return {}
    r0 = rows[0]
    k_code = _find_key(r0, "ISP Code", "ISPCode", "ISP_Code")
    k_cat  = _find_key(r0, "Store Category", "StoreCategory", "Category")
    k_ret  = _find_key(r0, "Parent Retailer Name", "ParentRetailerName",
                       "Retailer Name", "Retailer")
    out = {}
    for r in rows:
        code = (r.get(k_code) or "").strip() if k_code else ""
        if code:
            out[code] = {
                "category": (r.get(k_cat) or "").strip() if k_cat else "",
                "retailer": (r.get(k_ret) or "").strip() if k_ret else "",
            }
    return out

def load_weekoff():
    """Return dict of isp_code -> weekday name (e.g. {'MTRL1753': 'Thursday'})."""
    if not WEEKOFF_CSV_URL:
        return {}
    try:
        rows = _fetch_csv(WEEKOFF_CSV_URL)
    except Exception as e:
        print("[Remarks] WeekOff fetch failed (non-fatal): " + str(e))
        return {}
    if not rows:
        return {}
    r0 = rows[0]
    k_code = _find_key(r0, "ISP Code", "ISPCode", "ISP_Code", "Code")
    k_day  = _find_key(r0, "Week Off Day", "WeekOffDay", "WeekOff", "Off Day", "Day")
    out = {}
    for r in rows:
        code = (r.get(k_code) or "").strip() if k_code else ""
        day  = (r.get(k_day)  or "").strip() if k_day  else ""
        if code and day:
            out[code] = day
    print("[Remarks] WeekOff loaded: " + str(len(out)) + " ISDs")
    return out

def load_sales(target_date, targets, weekoff):
    if not SALES_CSV_URL:
        raise EnvironmentError("SALES_CSV_URL is not set")
    rows = _fetch_csv(SALES_CSV_URL)
    if not rows:
        print("[Remarks] Sales CSV returned no rows.")
        return [], []

    r0 = rows[0]
    def k(*c): return _find_key(r0, *c)

    k_code   = k("ISP Code",             "ISPCode")
    k_name   = k("ISP Name",             "ISPName")
    k_tsi    = k("TSI Name",             "TSIName")
    k_sc     = k("Store Code",           "StoreCode")
    k_ret    = k("Parent Retailer Name", "ParentRetailerName", "Retailer", "Store")
    k_city   = k("City")
    k_state  = k("State")
    k_date   = k("Date", "Sale Date",    "SaleDate")
    k_status = k("Status", "Sale Status","SaleStatus")

    def g(r, key): return (r.get(key) or "").strip() if key else ""

    zero, noup = {}, {}
    for r in rows:
        if _parse_date(g(r, k_date)) != target_date:
            continue
        code   = g(r, k_code)
        status = g(r, k_status)
        if not code:
            continue
        # Skip ISDs on their week-off day
        _DAY_NAMES = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
        _target_day = _DAY_NAMES[target_date.weekday()]
        _off = weekoff.get(code, "")
        if _off and _off.lower() == _target_day.lower():
            continue
        tgt      = targets.get(code, {})
        retailer = g(r, k_ret) or tgt.get("retailer", "")
        isd = dict(
            code=code, name=g(r, k_name), tsi=g(r, k_tsi),
            store_code=g(r, k_sc), retailer=retailer,
            city=g(r, k_city), state=g(r, k_state),
            category=tgt.get("category", ""),
        )
        if status == "Zero Sale Uploaded":
            zero[code] = isd
        elif status == "Sale Not Uploaded":
            noup[code] = isd

    srt = lambda x: (x["tsi"].lower(), x["name"].lower())
    return sorted(zero.values(), key=srt), sorted(noup.values(), key=srt)

# ── Google auth ───────────────────────────────────────────────────────────────

def get_creds():
    return Credentials(
        token=None,
        refresh_token=os.environ["GOOGLE_REFRESH_TOKEN"],
        token_uri="https://oauth2.googleapis.com/token",
        client_id=os.environ["GOOGLE_CLIENT_ID"],
        client_secret=os.environ["GOOGLE_CLIENT_SECRET"],
        # No scopes kwarg — lets Google use whatever was originally authorized
    )

# ── Drive helpers ─────────────────────────────────────────────────────────────

def find_or_create_remarks_file(drv, svc, folder_id):
    safe = REMARKS_FILE_NAME.replace("'", "\\'")
    q = (f"'{folder_id}' in parents "
         f"and mimeType='application/vnd.google-apps.spreadsheet' "
         f"and name='{safe}' and trashed=false")
    files = drv.files().list(q=q, fields="files(id,name)").execute().get("files", [])
    if files:
        print("[Remarks] Found existing file id=" + files[0]["id"])
        return files[0]["id"]
    f = drv.files().create(
        body={"name": REMARKS_FILE_NAME,
              "mimeType": "application/vnd.google-apps.spreadsheet",
              "parents": [folder_id]},
        fields="id",
    ).execute()
    print("[Remarks] Created new file id=" + f["id"])
    return f["id"]

def tab_exists(svc, file_id, tab_name):
    """Return True if a tab with this name already exists in the file."""
    sheets = svc.spreadsheets().get(spreadsheetId=file_id).execute().get("sheets", [])
    return any(s["properties"]["title"] == tab_name for s in sheets)

def delete_tab_if_exists(svc, file_id, tab_name):
    sheets = svc.spreadsheets().get(spreadsheetId=file_id).execute().get("sheets", [])
    for s in sheets:
        if s["properties"]["title"] == tab_name:
            sid = s["properties"]["sheetId"]
            svc.spreadsheets().batchUpdate(
                spreadsheetId=file_id,
                body={"requests": [{"deleteSheet": {"sheetId": sid}}]},
            ).execute()
            print("[Remarks] Deleted existing tab: " + tab_name)
            return

def delete_default_sheet(svc, file_id):
    sheets = svc.spreadsheets().get(spreadsheetId=file_id).execute().get("sheets", [])
    if len(sheets) <= 1:
        return
    for s in sheets:
        if s["properties"]["title"] == "Sheet1":
            svc.spreadsheets().batchUpdate(
                spreadsheetId=file_id,
                body={"requests": [{"deleteSheet": {"sheetId": s["properties"]["sheetId"]}}]},
            ).execute()
            print("[Remarks] Cleaned up default Sheet1")
            break

# ── Sharing & column protection ───────────────────────────────────────────────

def get_owner_email(creds):
    """Return the email of the authenticated Google account."""
    try:
        resp = requests.get(
            "https://www.googleapis.com/oauth2/v3/userinfo",
            headers={"Authorization": "Bearer " + (creds.token or "")},
            timeout=10,
        )
        return resp.json().get("email", "")
    except Exception as e:
        print("[Remarks] Could not get owner email (protection skipped): " + str(e))
        return ""

def share_with_tsis(drv, file_id, emails):
    """
    Share the file with each email as editor (writer role).
    Idempotent — silently skips if already shared.
    No notification email sent (avoids spamming TSIs on every daily run).
    """
    for email in emails:
        try:
            drv.permissions().create(
                fileId=file_id,
                body={"type": "user", "role": "writer", "emailAddress": email},
                sendNotificationEmail=False,
            ).execute()
            print("[Remarks] Shared with " + email)
        except Exception as e:
            print("[Remarks] Share " + email + " skipped (likely already has access): " + str(e))

def protect_tab_columns(svc, file_id, sheet_id, owner_email):
    """
    Lock all columns on this tab EXCEPT column I (index 8, Remarks).
    Only owner_email can edit the locked columns.
    TSI editors (added via share_with_tsis) can only edit the unprotected column I.

    Two protected ranges:
      - Columns A-H  (startColumnIndex 0, endColumnIndex 8)
      - Columns J-Z  (startColumnIndex 9, endColumnIndex 26)
    Column I (index 8) is left unprotected intentionally.
    """
    if not owner_email:
        print("[Remarks] Skipping column protection — owner email not available")
        return

    editors = {"users": [owner_email], "domainUsersCanEdit": False}
    reqs = [
        {"addProtectedRange": {"protectedRange": {
            "range": {"sheetId": sheet_id,
                      "startColumnIndex": 0, "endColumnIndex": 8},
            "description": "Locked - owner edit only",
            "warningOnly": False,
            "editors": editors,
        }}},
        {"addProtectedRange": {"protectedRange": {
            "range": {"sheetId": sheet_id,
                      "startColumnIndex": 9, "endColumnIndex": 26},
            "description": "Locked - owner edit only",
            "warningOnly": False,
            "editors": editors,
        }}},
    ]
    svc.spreadsheets().batchUpdate(
        spreadsheetId=file_id, body={"requests": reqs}
    ).execute()
    print("[Remarks] Column protection applied — column I (Remarks) is editable by TSIs")

# ── Sheet content builder ─────────────────────────────────────────────────────

class _Builder:
    def __init__(self):
        self.values   = []
        self.requests = []
        self._sid     = 0

    def _r(self): return len(self.values)

    def _add(self, cells):
        padded = list(cells) + [""] * max(0, N - len(cells))
        self.values.append(padded[:N])

    def _fmt(self, r1, c1, r2, c2, **kw):
        uf = {}
        if "bg"     in kw: uf["backgroundColor"]    = kw["bg"]
        if "halign" in kw: uf["horizontalAlignment"] = kw["halign"]
        if "wrap"   in kw: uf["wrapStrategy"] = "WRAP" if kw["wrap"] else "OVERFLOW_CELL"
        tf = {}
        if "bold"   in kw: tf["bold"]            = kw["bold"]
        if "italic" in kw: tf["italic"]          = kw["italic"]
        if "fg"     in kw: tf["foregroundColor"] = kw["fg"]
        if "size"   in kw: tf["fontSize"]        = kw["size"]
        if tf: uf["textFormat"] = tf
        if not uf: return
        self.requests.append({"repeatCell": {
            "range": {"sheetId": self._sid,
                      "startRowIndex": r1, "endRowIndex": r2,
                      "startColumnIndex": c1, "endColumnIndex": c2},
            "cell": {"userEnteredFormat": uf},
            "fields": "userEnteredFormat(" + ",".join(uf.keys()) + ")",
        }})

    def _merge(self, r1, c1, r2, c2):
        self.requests.append({"mergeCells": {
            "range": {"sheetId": self._sid,
                      "startRowIndex": r1, "endRowIndex": r2,
                      "startColumnIndex": c1, "endColumnIndex": c2},
            "mergeType": "MERGE_ALL",
        }})

    def _section_header(self, text, color):
        row = self._r()
        self._add([text])
        self._fmt(row, 0, row+1, N, bg=color, bold=True, fg=C_WHITE_TEXT, size=11)
        self._merge(row, 0, row+1, N)

    def _col_headers(self):
        row = self._r()
        self._add(COLS)
        self._fmt(row, 0, row+1, N, bg=C_COL_HDR, bold=True, fg=C_GRAY_TEXT, size=10)

    def _data_rows(self, isds, alt_bg):
        for i, isd in enumerate(isds):
            row = self._r()
            self._add([isd["tsi"], isd["code"], isd["name"], isd["store_code"],
                       isd["retailer"], isd["city"], isd["state"], isd["category"], ""])
            bg = alt_bg if i % 2 == 0 else C_WHITE
            self._fmt(row, 0, row+1, N-1, bg=bg)
            self._fmt(row, N-1, row+1, N, bg=C_REMARKS, wrap=True)

    def _tsi_remarks_block(self, isds, title, alt_bg):
        self._add([])
        row = self._r()
        self._add([title])
        self._fmt(row, 0, row+1, N, bg=C_TSI_HDR, bold=True, size=10)
        self._merge(row, 0, row+1, N)
        row = self._r()
        self._add(["TSI Name", "Overall Remarks"] + [""]*(N-2))
        self._fmt(row, 0, row+1, 1, bg=C_COL_HDR, bold=True, fg=C_GRAY_TEXT)
        self._fmt(row, 1, row+1, N, bg=C_COL_HDR, bold=True, fg=C_GRAY_TEXT)
        seen = list(dict.fromkeys(isd["tsi"] for isd in isds)) or ["--"]
        for tsi in seen:
            row = self._r()
            self._add([tsi, ""] + [""]*(N-2))
            self._fmt(row, 0, row+1, 1, bg=alt_bg, bold=True)
            self._fmt(row, 1, row+1, N, bg=C_REMARKS, wrap=True)

    def build(self, zero_list, noup_list, tab_name, target_date):
        # Title
        title = ("Zero Sale & No Sales Upload  --  Remarks Tracker"
                 "     |     " + tab_name + " " + str(target_date.year))
        row = self._r()
        self._add([title])
        self._fmt(row, 0, row+1, N, bg=C_MOTO_RED, bold=True,
                  fg=C_WHITE_TEXT, size=13, halign="CENTER")
        self._merge(row, 0, row+1, N)
        self._add([])

        # Section A — Zero Sale
        cnt_z = len(zero_list)
        self._section_header(
            "  ZERO SALE ISDs   (" + str(cnt_z) + " ISD" + ("s" if cnt_z != 1 else "") + ")",
            C_ORANGE_SEC)
        self._col_headers()
        if zero_list:
            self._data_rows(zero_list, C_ORANGE_ALT)
        else:
            row = self._r()
            self._add(["No Zero Sale ISDs recorded for this date."])
            self._fmt(row, 0, row+1, N, italic=True, fg=C_GRAY_TEXT)
        self._tsi_remarks_block(zero_list, "TSI Remarks -- Zero Sale ISDs", C_ORANGE_ALT)

        self._add([]); self._add([]); self._add([])   # 3 blank separator rows

        # Section B — No Sales Upload
        cnt_n = len(noup_list)
        self._section_header(
            "  NO SALES UPLOAD ISDs   (" + str(cnt_n) + " ISD" + ("s" if cnt_n != 1 else "") + ")",
            C_RED_SEC)
        self._col_headers()
        if noup_list:
            self._data_rows(noup_list, C_RED_ALT)
        else:
            row = self._r()
            self._add(["No No-Upload ISDs recorded for this date."])
            self._fmt(row, 0, row+1, N, italic=True, fg=C_GRAY_TEXT)
        self._tsi_remarks_block(noup_list, "TSI Remarks -- No Sales Upload ISDs", C_RED_ALT)

        return self.values, self.requests

# ── Tab writer ────────────────────────────────────────────────────────────────

def write_tab(svc, file_id, tab_name, zero_list, noup_list, target_date, owner_email=""):
    # Create sheet tab
    resp = svc.spreadsheets().batchUpdate(
        spreadsheetId=file_id,
        body={"requests": [{"addSheet": {"properties": {
            "title": tab_name,
            "gridProperties": {"rowCount": 600, "columnCount": N+1},
        }}}]},
    ).execute()
    sheet_id = resp["replies"][0]["addSheet"]["properties"]["sheetId"]

    # Build content
    builder = _Builder()
    builder._sid = sheet_id
    values, fmt_reqs = builder.build(zero_list, noup_list, tab_name, target_date)

    # Write values
    svc.spreadsheets().values().update(
        spreadsheetId=file_id,
        range="'" + tab_name + "'!A1",
        valueInputOption="RAW",
        body={"values": values},
    ).execute()

    # Column widths: TSI, Code, ISD Name, StoreCode, Retailer, City, State, Cat, Remarks
    col_px = [145, 90, 165, 90, 185, 90, 90, 80, 245]
    dim_reqs = [{"updateDimensionProperties": {
        "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                  "startIndex": i, "endIndex": i+1},
        "properties": {"pixelSize": px}, "fields": "pixelSize",
    }} for i, px in enumerate(col_px)]

    freeze_req = [{"updateSheetProperties": {
        "properties": {"sheetId": sheet_id,
                       "gridProperties": {"frozenRowCount": 1}},
        "fields": "gridProperties.frozenRowCount",
    }}]

    svc.spreadsheets().batchUpdate(
        spreadsheetId=file_id,
        body={"requests": fmt_reqs + dim_reqs + freeze_req},
    ).execute()

    print("[Remarks] Written tab '" + tab_name + "': "
          + str(len(zero_list)) + " zero-sale, "
          + str(len(noup_list)) + " no-upload ISDs ("
          + str(len(values)) + " rows)")

    # Lock all columns except column I (Remarks)
    protect_tab_columns(svc, file_id, sheet_id, owner_email)

# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    now_ist     = datetime.now(IST)
    target_dt   = now_ist - timedelta(days=1)   # always process yesterday
    target_date = target_dt.date()
    tab         = tab_name_for(target_dt)

    print("[Remarks] Run at " + now_ist.strftime("%d-%b %H:%M IST")
          + " -> processing " + str(target_date)
          + " -> tab '" + tab + "'")

    creds     = get_creds()
    drv       = build("drive",  "v3", credentials=creds)
    svc       = build("sheets", "v4", credentials=creds)
    folder_id = os.environ["GDRIVE_FOLDER_ID"]

    print("[Remarks] Fetching Targets CSV...")
    targets = load_targets()
    print("[Remarks] Targets loaded: " + str(len(targets)) + " ISPs")

    print("[Remarks] Fetching WeekOff CSV...")
    weekoff = load_weekoff()

    print("[Remarks] Fetching Sales CSV...")
    zero_list, noup_list = load_sales(target_date, targets, weekoff)
    print("[Remarks] Flagged ISDs - Zero sale: " + str(len(zero_list))
          + ", No upload: " + str(len(noup_list)))

    file_id = find_or_create_remarks_file(drv, svc, folder_id)

    # Share with TSI editors (idempotent)
    if TSI_EDITOR_EMAILS:
        share_with_tsis(drv, file_id, TSI_EDITOR_EMAILS)
    else:
        print("[Remarks] TSI_EDITOR_EMAILS is empty -- skipping share step")

    # If a tab for this date already exists, do NOT rebuild it. This preserves
    # any remarks TSIs have already entered — a re-run or a delayed/duplicate
    # scheduled run will no longer wipe the day's sheet.
    if tab_exists(svc, file_id, tab):
        print("[Remarks] Tab '" + tab + "' already exists — skipping rebuild "
              "to preserve existing remarks. File id=" + file_id)
        return

    # Get owner email for column protection
    owner_email = get_owner_email(creds)
    print("[Remarks] Owner email: " + (owner_email or "(not found)"))

    write_tab(svc, file_id, tab, zero_list, noup_list, target_date, owner_email)
    delete_default_sheet(svc, file_id)

    print("[Remarks] Complete. File id=" + file_id)


if __name__ == "__main__":
    main()

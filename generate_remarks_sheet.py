#!/usr/bin/env python3
"""
generate_remarks_sheet.py
─────────────────────────────────────────────────────────────────────────────
Reads today's ISP sales data from the published Sales CSV, filters
Zero Sale Uploaded + Sale Not Uploaded ISDs, and writes a richly formatted
tab (named by date, e.g. "21st May") into a Google Sheets file called
"Remarks Required for Zero & No Sales Upload" in the shared Drive folder.

Each tab contains:
  • Section A — ZERO SALE ISDs, sorted alphabetically by TSI Name
      - Data table with per-ISD Remarks column
      - TSI Remarks block (one row per TSI for an overall comment)
  • 3 blank separator rows
  • Section B — NO SALES UPLOAD ISDs, sorted alphabetically by TSI Name
      - Same structure as Section A

Running the script again on the same day performs a full refresh
(the existing tab is deleted and recreated with the latest data).

Required environment variables:
  GOOGLE_CLIENT_ID      — OAuth2 client id
  GOOGLE_CLIENT_SECRET  — OAuth2 client secret
  GOOGLE_REFRESH_TOKEN  — OAuth2 refresh token (Drive + Sheets scopes)
  GDRIVE_FOLDER_ID      — Drive folder id where the Remarks file lives
  SALES_CSV_URL         — Published CSV URL of the auto-converted Sales Sheet
  TARGETS_CSV_URL       — Published CSV URL of the Targets Sheet (for Category)
"""

import csv
import io
import os
import re
from datetime import datetime, timezone, timedelta

import requests
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

# ── Constants ────────────────────────────────────────────────────────────────

IST = timezone(timedelta(hours=5, minutes=30))
REMARKS_FILE_NAME = "Remarks Required for Zero & No Sales Upload"

SALES_CSV_URL   = os.environ.get("SALES_CSV_URL",   "")
TARGETS_CSV_URL = os.environ.get("TARGETS_CSV_URL", "")

# Column list for the data table (order matters — matches write_tab logic)
COLS = [
    "TSI Name", "ISP Code", "ISD Name", "Store Code",
    "Retailer / Store", "City", "State", "Category", "Remarks",
]
N = len(COLS)


# ── Colour palette (all values 0–1 floats) ───────────────────────────────────

def _c(r, g, b):
    return {"red": r / 255, "green": g / 255, "blue": b / 255}

C_MOTO_RED   = _c(225,  20,  10)   # #E1140A  Motorola red  — main title
C_ORANGE_SEC = _c(230,  81,   0)   # #E65100  deep orange   — Zero Sale header
C_RED_SEC    = _c(183,  28,  28)   # #B71C1C  dark red      — No Upload header
C_COL_HDR    = _c(245, 245, 245)   # #F5F5F5  light grey    — column header bg
C_ORANGE_ALT = _c(255, 243, 224)   # #FFF3E0  light orange  — Zero Sale alt rows
C_RED_ALT    = _c(255, 235, 238)   # #FFEBEE  light red     — No Upload alt rows
C_WHITE      = _c(255, 255, 255)
C_REMARKS    = _c(255, 253, 231)   # #FFFDE7  light yellow  — editable Remarks cells
C_TSI_HDR    = _c(197, 202, 233)   # #C5CAE9  indigo tint   — TSI block header
C_GRAY_TEXT  = _c( 66,  66,  66)
C_WHITE_TEXT = _c(255, 255, 255)


# ── Ordinal helpers ───────────────────────────────────────────────────────────

def _ordinal(n):
    """Return ordinal string: 1 → '1st', 22 → '22nd', etc."""
    suffix = (
        "th" if 11 <= (n % 100) <= 13
        else {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    )
    return f"{n}{suffix}"


def tab_name_for(dt):
    """Return Google Sheets tab name, e.g. '21st May'."""
    return f"{_ordinal(dt.day)} {dt.strftime('%b')}"


# ── CSV / column helpers ──────────────────────────────────────────────────────

def _norm(s):
    """Normalise a header name for fuzzy matching."""
    return re.sub(r"[\s_\-]+", "", str(s)).lower()


def _find_key(row_dict, *candidates):
    """Return the actual dict key that matches any of the candidate names."""
    norm_map = {_norm(k): k for k in row_dict}
    for c in candidates:
        match = norm_map.get(_norm(c))
        if match:
            return match
    return None


def _fetch_csv(url):
    """Fetch a published Google Sheets CSV and return list-of-dicts."""
    cache_bust = f"&_t={int(datetime.now().timestamp())}"
    resp = requests.get(url + cache_bust, timeout=30)
    resp.raise_for_status()
    text = resp.text.lstrip("﻿")                # strip BOM
    lines = [l for l in text.splitlines() if l.strip().strip(",")]
    if not lines:
        return []
    return list(csv.DictReader(lines))


def _parse_date(s):
    """Parse date strings like '21-May-26', '21-May-2026'."""
    for fmt in ("%d-%b-%y", "%d-%b-%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(str(s).strip(), fmt).date()
        except ValueError:
            pass
    return None


# ── Data loading ──────────────────────────────────────────────────────────────

def load_targets():
    """Return dict of isp_code → {category: str}."""
    if not TARGETS_CSV_URL:
        return {}
    try:
        rows = _fetch_csv(TARGETS_CSV_URL)
    except Exception as e:
        print(f"[Remarks] Targets fetch failed (non-fatal): {e}")
        return {}
    if not rows:
        return {}
    r0 = rows[0]
    k_code = _find_key(r0, "ISP Code", "ISPCode", "ISP_Code")
    k_cat  = _find_key(r0, "Store Category", "StoreCategory", "Category")
    out = {}
    for r in rows:
        code = (r.get(k_code) or "").strip() if k_code else ""
        if code:
            out[code] = {
                "category": (r.get(k_cat) or "").strip() if k_cat else "",
            }
    return out


def load_sales(today_date, targets):
    """
    Return (zero_list, noup_list) — each a list of ISD dicts for today,
    sorted alphabetically by TSI Name then ISD Name.
    """
    if not SALES_CSV_URL:
        raise EnvironmentError("SALES_CSV_URL is not set")

    rows = _fetch_csv(SALES_CSV_URL)
    if not rows:
        print("[Remarks] Sales CSV returned no rows.")
        return [], []

    r0 = rows[0]
    def k(*cands): return _find_key(r0, *cands)

    k_code   = k("ISP Code",             "ISPCode")
    k_name   = k("ISP Name",             "ISPName")
    k_tsi    = k("TSI Name",             "TSIName")
    k_sc     = k("Store Code",           "StoreCode")
    k_ret    = k("Parent Retailer Name", "ParentRetailerName", "Retailer", "Store")
    k_city   = k("City")
    k_state  = k("State")
    k_date   = k("Date",  "Sale Date",   "SaleDate")
    k_status = k("Status","Sale Status", "SaleStatus")

    def g(r, key): return (r.get(key) or "").strip() if key else ""

    zero, noup = {}, {}
    for r in rows:
        if _parse_date(g(r, k_date)) != today_date:
            continue
        code   = g(r, k_code)
        status = g(r, k_status)
        if not code:
            continue
        tgt = targets.get(code, {})
        isd = dict(
            code=code,
            name=g(r, k_name),
            tsi=g(r, k_tsi),
            store_code=g(r, k_sc),
            retailer=g(r, k_ret),
            city=g(r, k_city),
            state=g(r, k_state),
            category=tgt.get("category", ""),
        )
        if status == "Zero Sale Uploaded":
            zero[code] = isd
        elif status == "Sale Not Uploaded":
            noup[code] = isd

    sort_key = lambda x: (x["tsi"].lower(), x["name"].lower())
    return sorted(zero.values(), key=sort_key), sorted(noup.values(), key=sort_key)


# ── Google API auth ───────────────────────────────────────────────────────────

def get_creds():
    return Credentials(
        token=None,
        refresh_token=os.environ["GOOGLE_REFRESH_TOKEN"],
        token_uri="https://oauth2.googleapis.com/token",
        client_id=os.environ["GOOGLE_CLIENT_ID"],
        client_secret=os.environ["GOOGLE_CLIENT_SECRET"],
        scopes=[
            "https://www.googleapis.com/auth/drive",
            "https://www.googleapis.com/auth/spreadsheets",
        ],
    )


# ── Drive helpers ─────────────────────────────────────────────────────────────

def find_or_create_remarks_file(drv, svc, folder_id):
    """Return the Sheets file id, creating the file if it doesn't exist yet."""
    safe_name = REMARKS_FILE_NAME.replace("'", "\\'")
    q = (
        f"'{folder_id}' in parents "
        f"and mimeType='application/vnd.google-apps.spreadsheet' "
        f"and name='{safe_name}' "
        f"and trashed=false"
    )
    files = drv.files().list(q=q, fields="files(id,name)").execute().get("files", [])
    if files:
        print(f"[Remarks] Found existing file id={files[0]['id']}")
        return files[0]["id"]

    # Create brand-new file
    f = drv.files().create(
        body={
            "name": REMARKS_FILE_NAME,
            "mimeType": "application/vnd.google-apps.spreadsheet",
            "parents": [folder_id],
        },
        fields="id",
    ).execute()
    print(f"[Remarks] Created new Remarks file id={f['id']}")
    return f["id"]


def delete_tab_if_exists(svc, file_id, tab_name):
    """Delete a sheet tab by name if present (full-refresh strategy)."""
    sheets = svc.spreadsheets().get(spreadsheetId=file_id).execute().get("sheets", [])
    for s in sheets:
        if s["properties"]["title"] == tab_name:
            sid = s["properties"]["sheetId"]
            svc.spreadsheets().batchUpdate(
                spreadsheetId=file_id,
                body={"requests": [{"deleteSheet": {"sheetId": sid}}]},
            ).execute()
            print(f'[Remarks] Deleted existing tab "{tab_name}"')
            return


def delete_default_sheet(svc, file_id):
    """Remove the auto-created 'Sheet1' tab if other tabs now exist."""
    sheets = svc.spreadsheets().get(spreadsheetId=file_id).execute().get("sheets", [])
    if len(sheets) <= 1:
        return
    for s in sheets:
        if s["properties"]["title"] == "Sheet1":
            svc.spreadsheets().batchUpdate(
                spreadsheetId=file_id,
                body={"requests": [{"deleteSheet": {"sheetId": s["properties"]["sheetId"]}}]},
            ).execute()
            print("[Remarks] Cleaned up default 'Sheet1' tab")
            break


# ── Sheet content builder ─────────────────────────────────────────────────────

class _Builder:
    """
    Accumulates cell values and Sheets API batchUpdate format requests.
    sheet_id is injected after the tab is created (starts as placeholder 0).
    """

    def __init__(self):
        self.values = []        # list[list[str]]
        self.requests = []      # batchUpdate request dicts
        self._sid = 0           # sheetId placeholder

    # ── Row helpers ──────────────────────────────────────────────────────────

    def _r(self):
        """Current next row index (0-based)."""
        return len(self.values)

    def _add(self, cells):
        padded = list(cells) + [""] * max(0, N - len(cells))
        self.values.append(padded[:N])

    # ── Format / merge helpers ───────────────────────────────────────────────

    def _fmt(self, r1, c1, r2, c2, **kw):
        uf = {}
        if "bg"    in kw: uf["backgroundColor"] = kw["bg"]
        if "halign" in kw: uf["horizontalAlignment"] = kw["halign"]
        if "valign" in kw: uf["verticalAlignment"]   = kw["valign"]
        if "wrap"   in kw: uf["wrapStrategy"] = "WRAP" if kw["wrap"] else "OVERFLOW_CELL"

        tf = {}
        if "bold"   in kw: tf["bold"]            = kw["bold"]
        if "italic" in kw: tf["italic"]          = kw["italic"]
        if "fg"     in kw: tf["foregroundColor"] = kw["fg"]
        if "size"   in kw: tf["fontSize"]        = kw["size"]
        if tf:
            uf["textFormat"] = tf

        if not uf:
            return

        self.requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": self._sid,
                    "startRowIndex": r1, "endRowIndex": r2,
                    "startColumnIndex": c1, "endColumnIndex": c2,
                },
                "cell": {"userEnteredFormat": uf},
                "fields": "userEnteredFormat(" + ",".join(uf.keys()) + ")",
            }
        })

    def _merge(self, r1, c1, r2, c2):
        self.requests.append({
            "mergeCells": {
                "range": {
                    "sheetId": self._sid,
                    "startRowIndex": r1, "endRowIndex": r2,
                    "startColumnIndex": c1, "endColumnIndex": c2,
                },
                "mergeType": "MERGE_ALL",
            }
        })

    # ── Compound block helpers ───────────────────────────────────────────────

    def _section_header(self, text, color):
        row = self._r()
        self._add([text])
        self._fmt(row, 0, row + 1, N, bg=color, bold=True, fg=C_WHITE_TEXT, size=11)
        self._merge(row, 0, row + 1, N)

    def _col_headers(self):
        row = self._r()
        self._add(COLS)
        self._fmt(row, 0, row + 1, N, bg=C_COL_HDR, bold=True, fg=C_GRAY_TEXT, size=10)

    def _data_rows(self, isds, alt_bg):
        for i, isd in enumerate(isds):
            row = self._r()
            self._add([
                isd["tsi"], isd["code"], isd["name"], isd["store_code"],
                isd["retailer"], isd["city"], isd["state"], isd["category"], "",
            ])
            bg = alt_bg if i % 2 == 0 else C_WHITE
            self._fmt(row, 0, row + 1, N - 1, bg=bg)               # data cells
            self._fmt(row, N - 1, row + 1, N, bg=C_REMARKS, wrap=True)  # Remarks cell

    def _tsi_remarks_block(self, isds, title, alt_bg):
        """
        Writes the 'TSI Remarks' block at the end of a section:
          • sub-header row
          • column-header row (TSI Name | Overall Remarks)
          • one editable row per unique TSI
        """
        # blank separator
        self._add([])

        # block header
        row = self._r()
        self._add([title])
        self._fmt(row, 0, row + 1, N, bg=C_TSI_HDR, bold=True, size=10)
        self._merge(row, 0, row + 1, N)

        # column headers
        row = self._r()
        self._add(["TSI Name", "Overall Remarks"] + [""] * (N - 2))
        self._fmt(row, 0, row + 1, 1, bg=C_COL_HDR, bold=True, fg=C_GRAY_TEXT)
        self._fmt(row, 1, row + 1, N, bg=C_COL_HDR, bold=True, fg=C_GRAY_TEXT)

        # one editable row per unique TSI (in the order they appear, i.e. alpha)
        seen = list(dict.fromkeys(isd["tsi"] for isd in isds))  # deduplicated, ordered
        if not seen:
            seen = ["—"]
        for tsi in seen:
            row = self._r()
            self._add([tsi, ""] + [""] * (N - 2))
            self._fmt(row, 0, row + 1, 1, bg=alt_bg, bold=True)
            self._fmt(row, 1, row + 1, N, bg=C_REMARKS, wrap=True)

    # ── Main builder ─────────────────────────────────────────────────────────

    def build(self, zero_list, noup_list, tab_name, today_date):
        # ── Title row ────────────────────────────────────────────────────────
        title = (
            f"Zero Sale & No Sales Upload  —  Remarks Tracker"
            f"     |     {tab_name} {today_date.year}"
        )
        row = self._r()
        self._add([title])
        self._fmt(row, 0, row + 1, N,
                  bg=C_MOTO_RED, bold=True, fg=C_WHITE_TEXT, size=13, halign="CENTER")
        self._merge(row, 0, row + 1, N)
        self._add([])   # blank spacer

        # ─────────────────────────────────────────────────────────────────────
        # SECTION A — Zero Sale ISDs
        # ─────────────────────────────────────────────────────────────────────
        cnt_z = len(zero_list)
        self._section_header(
            f"⚠   ZERO SALE ISDs   ({cnt_z} ISD{'s' if cnt_z != 1 else ''})",
            C_ORANGE_SEC,
        )
        self._col_headers()

        if zero_list:
            self._data_rows(zero_list, C_ORANGE_ALT)
        else:
            row = self._r()
            self._add(["No Zero Sale ISDs recorded for this date."])
            self._fmt(row, 0, row + 1, N, italic=True, fg=C_GRAY_TEXT)

        self._tsi_remarks_block(
            zero_list, "TSI Remarks — Zero Sale ISDs", C_ORANGE_ALT
        )

        # 3 blank rows between sections
        self._add([])
        self._add([])
        self._add([])

        # ─────────────────────────────────────────────────────────────────────
        # SECTION B — No Sales Upload ISDs
        # ─────────────────────────────────────────────────────────────────────
        cnt_n = len(noup_list)
        self._section_header(
            f"🚫   NO SALES UPLOAD ISDs   ({cnt_n} ISD{'s' if cnt_n != 1 else ''})",
            C_RED_SEC,
        )
        self._col_headers()

        if noup_list:
            self._data_rows(noup_list, C_RED_ALT)
        else:
            row = self._r()
            self._add(["No No-Upload ISDs recorded for this date."])
            self._fmt(row, 0, row + 1, N, italic=True, fg=C_GRAY_TEXT)

        self._tsi_remarks_block(
            noup_list, "TSI Remarks — No Sales Upload ISDs", C_RED_ALT
        )

        return self.values, self.requests


# ── Tab writer ────────────────────────────────────────────────────────────────

def write_tab(svc, file_id, tab_name, zero_list, noup_list, today_date):
    """Create a new sheet tab, populate values, and apply all formatting."""

    # 1. Create the new sheet tab
    resp = svc.spreadsheets().batchUpdate(
        spreadsheetId=file_id,
        body={"requests": [{"addSheet": {"properties": {
            "title": tab_name,
            "gridProperties": {"rowCount": 600, "columnCount": N + 1},
        }}}]},
    ).execute()
    sheet_id = resp["replies"][0]["addSheet"]["properties"]["sheetId"]

    # 2. Build content
    builder = _Builder()
    builder._sid = sheet_id
    values, fmt_reqs = builder.build(zero_list, noup_list, tab_name, today_date)

    # 3. Write cell values
    svc.spreadsheets().values().update(
        spreadsheetId=file_id,
        range=f"'{tab_name}'!A1",
        valueInputOption="RAW",
        body={"values": values},
    ).execute()

    # 4. Column widths (px): TSI, Code, ISD Name, StoreCode, Retailer, City, State, Cat, Remarks
    col_px = [145, 90, 165, 90, 185, 90, 90, 80, 245]
    dim_reqs = [
        {"updateDimensionProperties": {
            "range": {
                "sheetId": sheet_id, "dimension": "COLUMNS",
                "startIndex": i, "endIndex": i + 1,
            },
            "properties": {"pixelSize": px},
            "fields": "pixelSize",
        }}
        for i, px in enumerate(col_px)
    ]

    # 5. Freeze first row (title)
    freeze_req = [{"updateSheetProperties": {
        "properties": {
            "sheetId": sheet_id,
            "gridProperties": {"frozenRowCount": 1},
        },
        "fields": "gridProperties.frozenRowCount",
    }}]

    # 6. Apply formatting + dimensions + freeze in one batch
    svc.spreadsheets().batchUpdate(
        spreadsheetId=file_id,
        body={"requests": fmt_reqs + dim_reqs + freeze_req},
    ).execute()

    total_rows = len(values)
    print(
        f'[Remarks] Written tab "{tab_name}": '
        f"{len(zero_list)} zero-sale ISDs, {len(noup_list)} no-upload ISDs "
        f"({total_rows} rows total)"
    )


# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    now_ist    = datetime.now(IST)
    # Script runs at 12:03 PM IST the day AFTER the target date,
    # so always derive data date as yesterday (IST).
    target_dt   = now_ist - timedelta(days=1)
    today_date  = target_dt.date()
    tab         = tab_name_for(target_dt)

    print(f"[Remarks] ── Starting run at {now_ist.strftime('%d-%b %H:%M IST')}  →  processing {today_date}  →  tab '{tab}' ──")

    creds     = get_creds()
    drv       = build("drive",  "v3", credentials=creds)
    svc       = build("sheets", "v4", credentials=creds)
    folder_id = os.environ["GDRIVE_FOLDER_ID"]

    # Load reference data
    print("[Remarks] Fetching Targets CSV …")
    targets = load_targets()
    print(f"[Remarks] Targets loaded: {len(targets)} ISPs")

    print("[Remarks] Fetching Sales CSV …")
    zero_list, noup_list = load_sales(today_date, targets)
    print(f"[Remarks] Today's flagged ISDs — Zero sale: {len(zero_list)}, No upload: {len(noup_list)}")

    # Write to Drive
    file_id = find_or_create_remarks_file(drv, svc, folder_id)
    delete_tab_if_exists(svc, file_id, tab)
    write_tab(svc, file_id, tab, zero_list, noup_list, today_date)
    delete_default_sheet(svc, file_id)

    print(f"[Remarks] ── Complete. File id={file_id} ──")


if __name__ == "__main__":
    main()

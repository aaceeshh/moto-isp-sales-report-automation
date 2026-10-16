#!/usr/bin/env python3
"""
attendance_reports.py
---------------------
Shared logic for the Nural-sourced Week-Off / Leave / Late tracking.

The 4 PM & 11:59 PM scraper (fetch_weekoff_nural.py) writes a *date-stamped
log* into the Week-Off tab of the Targets spreadsheet. This module:
  - parses that published log,
  - exposes the per-date exclusion set (Weekly Off + Leave) and late set,
  - builds the daily and MTD reports (Telegram text + Google-Sheet rows).

Published log schema (Week-Off tab, same gid the pipelines already read):
    ISP Code | Date (YYYY-MM-DD) | Status | Leave Type | In Time | Late
      Status   : "Weekly Off" | "Leave" | "Present" | "Absent"
      Leave Type: e.g. "Casual Leave" / "Sick Leave" (only when Status=Leave)
      In Time  : e.g. "01:03 PM" (Present rows)
      Late     : "Y" when Status=Present and In Time > 12:30 PM, else ""

Exclusion rule (confirmed with owner): ISDs with Status in {Weekly Off, Leave}
are omitted from Zero-Sale / No-Upload flags. Late is a FLAG ONLY — late ISDs
were present and still had to upload, so they are NOT excluded.
"""

import re
import html
from datetime import datetime, date

# Statuses that remove an ISD from the Zero/No-Upload checks for that date.
EXCLUDE_STATUSES = {"weekly off", "leave"}
LATE_CUTOFF = (12, 45)   # 12:45 PM — strictly AFTER this is "late"

_MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
           "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# ── helpers ──────────────────────────────────────────────────────────────────

def esc(s):
    return html.escape(str(s or ""), quote=True)

def _norm(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()

def parse_iso(s):
    try:
        return datetime.strptime(str(s).strip(), "%Y-%m-%d").date()
    except ValueError:
        return None

def parse_clock(s):
    """'01:03 PM' -> (hour24, minute); None if unparseable."""
    s = str(s or "").strip().upper()
    m = re.match(r"^(\d{1,2}):(\d{2})\s*(AM|PM)$", s)
    if not m:
        return None
    h, mn, ap = int(m.group(1)), int(m.group(2)), m.group(3)
    if ap == "AM":
        h = 0 if h == 12 else h
    else:
        h = 12 if h == 12 else h + 12
    return (h, mn)

def is_late(in_time):
    t = parse_clock(in_time)
    if not t:
        return False
    return t > LATE_CUTOFF            # strictly after 12:30 PM

def fmt_day(d):
    return "%02d %s" % (d.day, _MONTHS[d.month - 1])

def _is_inactive(info):
    """True only when the ISD's Targets status is exactly 'Inactive'.
    NB: must be an exact match — the word 'inactive' contains 'active'."""
    if not info:
        return False
    return str(info.get("status", "")).strip().lower() == "inactive"

# ── log parsing ──────────────────────────────────────────────────────────────

def parse_log(rows):
    """rows: list of dicts from the published log CSV.
    Returns {(code, date): {status, leave_type, in_time, late}}."""
    def find_key(r, *cands):
        m = {re.sub(r"[\s_]+", "", k).lower(): k for k in r}
        for c in cands:
            kk = re.sub(r"[\s_]+", "", c).lower()
            if kk in m:
                return m[kk]
        return None
    if not rows:
        return {}
    r0 = rows[0]
    kC = find_key(r0, "ISP Code", "ISPCode", "User Code", "Code")
    kD = find_key(r0, "Date")
    kS = find_key(r0, "Status", "Attendance Status")
    kL = find_key(r0, "Leave Type", "LeaveType")
    kI = find_key(r0, "In Time", "InTime")
    kLate = find_key(r0, "Late")
    out = {}
    for r in rows:
        code = _norm(r.get(kC)) if kC else ""
        d = parse_iso(r.get(kD)) if kD else None
        if not code or not d:
            continue
        status = _norm(r.get(kS)) if kS else ""
        in_time = _norm(r.get(kI)) if kI else ""
        late = (_norm(r.get(kLate)).upper().startswith("Y")) if kLate else False
        if not late and status.lower() == "present":
            late = is_late(in_time)          # derive if column absent
        out[(code, d)] = {
            "status": status,
            "leave_type": _norm(r.get(kL)) if kL else "",
            "in_time": in_time,
            "late": late,
        }
    return out

def excluded_codes(log, target_date):
    """Set of ISP Codes to omit from Zero/No-Upload on target_date."""
    return {code for (code, d), rec in log.items()
            if d == target_date and rec["status"].lower() in EXCLUDE_STATUSES}

def late_entries(log, target_date):
    """{code: in_time} for Present-but-late ISDs on target_date."""
    return {code: rec["in_time"] for (code, d), rec in log.items()
            if d == target_date and rec["late"]}

# ── report builders ──────────────────────────────────────────────────────────
# codes: {ISP Code: {"tsi":..., "isd":..., "store":..., "city":..., "state":...}}
# tsi_order: list of TSI names in display order.

def _by_tsi(codes, tsi_order, items):
    """items: iterable of (code, extra). Returns {tsi: [(code, extra)...]} sorted."""
    g = {t: [] for t in tsi_order}
    for code, extra in items:
        info = codes.get(code)
        if not info or _is_inactive(info):
            continue
        t = info.get("tsi")
        if t in g:
            g[t].append((code, extra))
    for t in g:
        g[t].sort(key=lambda ce: (codes[ce[0]].get("isd") or "").lower())
    return g

def build_daily_report(target_date, codes, tsi_order, log):
    """Telegram HTML for the day's Weekly-Off / Leave / Late, per TSI + names."""
    wo   = [(c, None) for (c, d), r in log.items()
            if d == target_date and r["status"].lower() == "weekly off"]
    lv   = [(c, r["leave_type"]) for (c, d), r in log.items()
            if d == target_date and r["status"].lower() == "leave"]
    late = [(c, r["in_time"]) for (c, d), r in log.items()
            if d == target_date and r["late"]]
    pend = [(c, r["in_time"]) for (c, d), r in log.items()
            if d == target_date and "pending" in r["status"].lower()]
    gwo, glv, glt, gpd = (_by_tsi(codes, tsi_order, x) for x in (wo, lv, late, pend))

    def line(c, suffix=""):
        info = codes[c]
        s = "    • %s" % esc(info.get("isd") or c)
        if info.get("store"):
            s += " — %s" % esc(info["store"])
        if suffix:
            s += " · %s" % suffix
        return s

    L = ["\U0001F4CB <b>Attendance — Week-Off / Leave / Late / Pending</b>",
         "%s" % fmt_day(target_date), "━" * 20, ""]
    for t in tsi_order:
        wol, lvl, ltl, pdl = gwo[t], glv[t], glt[t], gpd[t]
        L.append("\U0001F464 <b>%s</b> — WO %d · Leave %d · Late %d · Pending %d"
                 % (esc(t), len(wol), len(lvl), len(ltl), len(pdl)))
        if wol:
            L.append("  🛌 <b>Weekly Off</b>")
            L += [line(c) for c, _ in wol]
        if lvl:
            L.append("  🧾 <b>Leave</b>")
            L += [line(c, esc(lt or "Leave")) for c, lt in lvl]
        if ltl:
            L.append("  ⏰ <b>Late &gt;12:45</b>")
            L += [line(c, esc(it)) for c, it in ltl]
        if pdl:
            L.append("  ⏳ <b>Approval Pending</b>")
            L += [line(c, esc(it)) for c, it in pdl]
        if not (wol or lvl or ltl or pdl):
            L.append("  <i>— none —</i>")
        L.append("")
    return "\n".join(L).rstrip()

def _collect_mtd(log, year, month):
    """Returns {code: {'wo':[dates], 'leave':[(date,type)], 'late':[(date,intime)]}}."""
    agg = {}
    for (code, d), r in log.items():
        if d.year != year or d.month != month:
            continue
        a = agg.setdefault(code, {"wo": [], "leave": [], "late": []})
        st = r["status"].lower()
        if st == "weekly off":
            a["wo"].append(d)
        elif st == "leave":
            a["leave"].append((d, r["leave_type"]))
        if r["late"]:
            a["late"].append((d, r["in_time"]))
    return agg

def build_mtd_report(year, month, codes, tsi_order, log):
    """Telegram HTML: per-TSI, per-ISD exact dates of WO / Leave / Late."""
    agg = _collect_mtd(log, year, month)
    hdr = "%s %d" % (_MONTHS[month - 1], year)
    # Range end for the "not marked" tally = last day with any data this month.
    month_days = [d.day for (c, d) in log if d.year == year and d.month == month]
    upto = max(month_days) if month_days else 0
    def _active(c):
        return not _is_inactive(codes.get(c))
    L = ["\U0001F4CA <b>Attendance MTD — %s</b>" % hdr,
         "<i>Week-Off · Leave · Late (&gt;12:45) · Not-marked, by ISD</i>", "━" * 20, ""]
    days = lambda ds: ", ".join("%02d" % d.day for d in sorted(ds))
    for t in tsi_order:
        tcodes = [c for c in agg if (codes.get(c) or {}).get("tsi") == t
                  and not _is_inactive(codes.get(c))]
        tcodes.sort(key=lambda c: (codes[c].get("isd") or "").lower())
        two = sum(len(agg[c]["wo"]) for c in tcodes)
        tlv = sum(len(agg[c]["leave"]) for c in tcodes)
        tlt = sum(len(agg[c]["late"]) for c in tcodes)
        # Not-marked ISD-days: active ISDs of this TSI with no record on days 1..upto.
        nomark = sum(1 for c in codes if codes[c].get("tsi") == t and _active(c)
                     for day in range(1, upto + 1)
                     if (c, date(year, month, day)) not in log)
        L.append("\U0001F464 <b>%s</b> — WO %d · Leave %d · Late %d · Not-marked %d"
                 % (esc(t), two, tlv, tlt, nomark))
        for c in tcodes:
            a = agg[c]
            parts = []
            if a["wo"]:
                parts.append("WO " + days(a["wo"]))
            if a["leave"]:
                parts.append("Leave " + days(d for d, _ in a["leave"]))
            if a["late"]:
                parts.append("Late " + days(d for d, _ in a["late"]))
            if parts:
                L.append("  %s — %s" % (esc(codes[c].get("isd") or c), " · ".join(parts)))
        L.append("")
    return "\n".join(L).rstrip()

# Google-Sheet rows (header + data) ------------------------------------------

def daily_sheet_rows(target_date, codes, tsi_order, log):
    head = ["TSI", "ISD Name", "ISP Code", "Store", "City",
            "Status", "Leave Type", "In Time", "Late"]
    rows = []
    for (c, d), r in log.items():
        if d != target_date or c not in codes or _is_inactive(codes.get(c)):
            continue
        st = r["status"].lower()
        if st not in EXCLUDE_STATUSES and not r["late"] and "pending" not in st:
            continue  # only exceptions in the daily tab (WO / Leave / Late / Pending)
        i = codes[c]
        rows.append([i.get("tsi", ""), i.get("isd", ""), c, i.get("store", ""),
                     i.get("city", ""), r["status"], r["leave_type"],
                     r["in_time"], "Y" if r["late"] else ""])
    rows.sort(key=lambda x: (x[0].lower(), x[1].lower()))
    return [head] + rows

def mtd_sheet_rows(year, month, codes, tsi_order, log):
    agg = _collect_mtd(log, year, month)
    head = ["TSI", "ISD Name", "ISP Code",
            "#WeekOff", "WeekOff Dates", "#Leave", "Leave Dates", "#Late", "Late Dates"]
    ds = lambda dates: ", ".join(fmt_day(d) for d in sorted(dates))
    rows = []
    for c, a in agg.items():
        if c not in codes or _is_inactive(codes.get(c)):
            continue
        if not (a["wo"] or a["leave"] or a["late"]):
            continue  # exceptions only — skip ISDs with nothing to report
        i = codes[c]
        rows.append([i.get("tsi", ""), i.get("isd", ""), c,
                     len(a["wo"]), ds(a["wo"]),
                     len(a["leave"]), ds(d for d, _ in a["leave"]),
                     len(a["late"]), ds(d for d, _ in a["late"])])
    rows.sort(key=lambda x: (x[0].lower(), x[1].lower()))
    return [head] + rows


# Matrix / register tabs ------------------------------------------------------

def _isd_order(codes, tsi_order):
    def key(c):
        t = (codes[c] or {}).get("tsi")
        return (tsi_order.index(t) if t in tsi_order else 99,
                (codes[c].get("isd") or "").lower())
    return sorted((c for c in codes if not _is_inactive(codes.get(c))), key=key)

def _status_word(rec):
    """Base attendance word (no late merge): Present / Weekly Off / Leave /
    Absent / Pending."""
    if not rec:
        return ""
    if "pending" in rec["status"].lower():
        return "Pending"
    return rec["status"]            # Present / Weekly Off / Leave / Absent

def _cell(rec):
    """Matrix cell (words): Present rows become 'Late' when late."""
    w = _status_word(rec)
    if w == "Present" and rec and rec["late"]:
        return "Late"
    return w

LEGEND = ("Note: 'Late' = reported after 12:45 PM · 'Leave' includes approved + "
          "pending · blank = no attendance record")

def attendance_matrix_rows(year, month, upto_day, codes, tsi_order, log):
    """MTD register: ISDs down rows, dates 01..upto_day across the top,
    status code in each cell. Overwritten daily."""
    title = "Attendance — %s %d" % (_MONTHS[month - 1], year)
    header = ["TSI", "ISD Name", "ISP Code"] + ["%02d" % d for d in range(1, upto_day + 1)]
    out = [[title], [LEGEND], header]
    for c in _isd_order(codes, tsi_order):
        info = codes[c]
        row = [info.get("tsi", ""), info.get("isd", ""), c]
        for d in range(1, upto_day + 1):
            row.append(_cell(log.get((c, date(year, month, d)))))
        out.append(row)
    return out

def attendance_ftd_rows(target_date, codes, tsi_order, log):
    """FTD register: all ISDs with today's status / in-time / late."""
    head = ["TSI", "ISD Name", "ISP Code", "Store",
            target_date.strftime("%d-%b-%Y"), "In Time", "Late"]
    out = [head]
    for c in _isd_order(codes, tsi_order):
        info = codes[c]
        rec = log.get((c, target_date))
        out.append([info.get("tsi", ""), info.get("isd", ""), c, info.get("store", ""),
                    _status_word(rec), rec["in_time"] if rec else "",
                    "Y" if (rec and rec["late"]) else ""])
    return out


def build_notmarked_messages(target_date, codes, tsi_order, log):
    """One Telegram message per TSI listing active ISDs who marked NEITHER
    attendance NOR leave on target_date (i.e. no record at all in the log)."""
    hdr = fmt_day(target_date)
    def active(c):
        return not _is_inactive(codes.get(c))
    msgs = []
    for t in tsi_order:
        miss = [c for c in codes
                if codes[c].get("tsi") == t and active(c)
                and (c, target_date) not in log]
        miss.sort(key=lambda c: (codes[c].get("isd") or "").lower())
        L = ["\U0001F6AB <b>%s — Not Marked</b>" % esc(t),
             "%s · attendance/leave not marked" % hdr,
             "<i>ISD · Store</i>", "━" * 20, ""]
        if miss:
            for i, c in enumerate(miss, 1):
                info = codes[c]
                line = "%d. <b>%s</b>" % (i, esc(info.get("isd") or c))
                if info.get("store"):
                    line += " — %s" % esc(info["store"])
                L.append(line)
        else:
            L.append("<i>All marked 🎉</i>")
        L.append("")
        L.append("<i>Count: %d</i>" % len(miss))
        msgs.append("\n".join(L))
    return msgs


# ── self-test (synthetic log, no network) ────────────────────────────────────
if __name__ == "__main__":
    codes = {
        "MTRL151": {"tsi": "Amit Kumar",  "isd": "Soumen Mondal",  "store": "Cdit Ultadanga", "city": "Kolkata", "state": "West Bengal"},
        "MTRL154": {"tsi": "Amit Kumar",  "isd": "Dhiraj Kumar",   "store": "Dumdum Rd",      "city": "Kolkata", "state": "West Bengal"},
        "MTRL2812":{"tsi": "Nabarun Bhattacharya", "isd": "Tasif Molla", "store": "Basirhat", "city": "Basirhat","state": "West Bengal"},
        "MTRL3054":{"tsi": "Joydwip Seal","isd": "Rubai Paul",     "store": "Siliguri",       "city": "Siliguri","state": "West Bengal"},
    }
    TSIS = ["Amit Kumar", "Joydwip Seal", "Mukesh Kumar", "Nabarun Bhattacharya"]
    synth = [
        {"ISP Code": "MTRL151", "Date": "2026-10-01", "Status": "Weekly Off", "Leave Type": "", "In Time": "", "Late": ""},
        {"ISP Code": "MTRL154", "Date": "2026-10-01", "Status": "Present", "Leave Type": "", "In Time": "01:03 PM", "Late": "Y"},
        {"ISP Code": "MTRL2812","Date": "2026-10-01", "Status": "Leave", "Leave Type": "Sick Leave", "In Time": "", "Late": ""},
        {"ISP Code": "MTRL3054","Date": "2026-10-01", "Status": "Present", "Leave Type": "", "In Time": "10:05 AM", "Late": ""},
        {"ISP Code": "MTRL151", "Date": "2026-10-08", "Status": "Weekly Off", "Leave Type": "", "In Time": "", "Late": ""},
        {"ISP Code": "MTRL2812","Date": "2026-10-05", "Status": "Leave", "Leave Type": "Casual Leave", "In Time": "", "Late": ""},
        {"ISP Code": "MTRL154", "Date": "2026-10-09", "Status": "Present", "Leave Type": "", "In Time": "12:48 PM", "Late": ""},
    ]
    log = parse_log(synth)
    td = date(2026, 10, 1)
    print("EXCLUDED on", td, "=>", sorted(excluded_codes(log, td)))
    print("LATE on", td, "=>", late_entries(log, td))
    print("\n--- DAILY ---\n" + build_daily_report(td, codes, TSIS, log))
    print("\n--- MTD ---\n" + build_mtd_report(2026, 10, codes, TSIS, log))
    print("\n--- DAILY SHEET ROWS ---")
    for r in daily_sheet_rows(td, codes, TSIS, log):
        print(r)
    print("\n--- MTD SHEET ROWS ---")
    for r in mtd_sheet_rows(2026, 10, codes, TSIS, log):
        print(r)

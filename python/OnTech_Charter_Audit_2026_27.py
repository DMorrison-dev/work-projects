"""Offline OnTech audit implementing charter_invoice_discrepancy_logic.md.

Reads embedded SchoolTool and Disability exports. Never connects to SQL Server,
uses credentials, writes back to SchoolTool, or edits an input workbook.

Usage:
    python OnTech_Charter_Audit_2026_27.py "invoice.xlsx" \
        --output "ONTECH_AUDITED.xlsx" [--calendar approved_calendar.csv]

Requirements: pip install openpyxl lxml

INPUT TABS ARE DETECTED AUTOMATICALLY (case, spacing, and "(17)"-style
suffixes are ignored):
* General roster: "Gen Roster", "General Roster", or "Roster"
* SPED roster:    "SPED Roster" or "SPED"
* SchoolTool:     any tab starting "ST Enroll" (trailing space optional)
* Disability:     any tab starting "Disability" -- "Disability (17)",
                  "Disability (18)", "Disability" all work. If several exist,
                  the one with current-year rows is used.
Override any detection with --gen-tab / --sped-tab / --st-tab / --disability-tab.
Columns are found by header name (any order, any header row in the first 10),
so reordered or extra export columns do not break the live formulas.

IMPORTANT LIMITS
---------------
* A match to StudentID_Disability is an export match, NOT proof of an active IEP.
* The original script's disability code/start/end fields cannot be reconstructed
  from this three-column Disability export and are left explicitly unavailable.
* The governing decision rules are implemented in Python and live Excel
  formulas. Missing entry dates or a one-sided missing exit produce REVIEW.
  Exact school-day gaps need a verified calendar; unknown counts remain REVIEW.
* Program boundaries are the user-supplied 2026-09-03 and 2027-06-30.
* Lookup fields are live INDEX references to the selected source Excel rows.
  Re-run this script if rows/keys are inserted, deleted, sorted, or rematched.
  Enrollment classifications/overall status/notes recalculate live. Identity
  matching, queue membership, and tentative FTE values are run snapshots.
* The embedded historical tab is retained but is not a substitute for the
  current ST Enrollment export or manual SchoolTool enrollment-history review.
"""

from __future__ import annotations

import argparse
import calendar
import csv
import hashlib
import io
import json
import math
import re
import sys
import unicodedata
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path
from zipfile import ZipFile

from lxml import etree as ET
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel, to_excel


YEAR = "2026-2027"
YEAR_START = date(2026, 7, 1)
YEAR_END = date(2027, 6, 30)
FTE_FIRST_DAY = date(2026, 9, 3)  # User-confirmed program/rule first day.
FIRST_SCHOOL_DAY = date(2026, 9, 3)
CHARTER_YEAR_END = date(2027, 6, 27)  # Observed OnTech roster rollover; preserved raw.
SS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
AUDIT_NAMES = ["ONTECH GEN VERIFIED", "ONTECH SPED VERIFIED", "ONTECH REVIEW", "ONTECH AUDIT SUMMARY", "ONTECH DATA CHECKS"]
GENERATED_TABS = set(AUDIT_NAMES) | {"AUDIT CONFIG", "SCHOOL CALENDAR", "VALIDATION REPORT"}

# Required columns per source tab. Matching ignores case, spaces and punctuation.
ROSTER_COLUMNS = ["Last Name", "First Name", "NYSSIS ID", "DOB", "Enroll Date", "Withdraw Date", "OnTECH FTE"]
ST_COLUMNS = ["StudentEnrollment_SchoolYear", "Student_UniversalStudentID", "StudentID", "LastName", "FirstName",
              "Person_DOB", "SchoolName", "StudentEnrollment_StartDate", "StudentEnrollment_EndDate"]
ST_OPTIONAL = ["StudentEnrollment_EntryNote", "StudentEnrollment_ExitNote"]
DIS_COLUMNS = ["SchoolYear_Disability", "Student_UniversalStudentID", "StudentID_Disability"]
COLUMN_ALIASES = {"nyssisid": {"nyssis", "nysisid"}}
ROSTER_DATE_FIELDS = ["DOB", "Enroll Date", "Withdraw Date"]
ST_DATE_FIELDS = ["Person_DOB", "StudentEnrollment_StartDate", "StudentEnrollment_EndDate"]


class AuditInputError(ValueError):
    """A problem with the input files that the user can fix."""


def norm_header(value):
    return re.sub(r"[^a-z0-9]", "", str(value or "").lower())


def norm_tab(name):
    text = re.sub(r"\(\s*\d+\s*\)", " ", str(name).lower())
    text = re.sub(r"[^a-z0-9]+", " ", text).strip()
    return re.sub(r"(\s+\d+)+$", "", text).strip()


TAB_RULES = {
    "gen": ("general roster", lambda n: n in {"gen roster", "general roster", "gen ed roster", "general ed roster", "roster", "gen", "general"}),
    "sped": ("SPED roster", lambda n: n in {"sped roster", "sped", "special ed roster", "special education roster", "sped ed roster"}),
    "st": ("ST Enroll", lambda n: n.startswith("st enroll") or n in {"schooltool enrollment", "st enrollment"}),
    "dis": ("Disability", lambda n: n.startswith("disability")),
}


def find_tab(sheetnames, role, override=None):
    """Return candidate tab names for a role (override must match exactly, ignoring case/spaces)."""
    label, rule = TAB_RULES[role]
    if override:
        hits = [s for s in sheetnames if s == override] or [s for s in sheetnames if s.strip().lower() == override.strip().lower()]
        if not hits:
            raise AuditInputError(f"Tab {override!r} given for {label} was not found. Tabs in workbook: {sheetnames}")
        return hits[:1]
    hits = [s for s in sheetnames if s not in GENERATED_TABS and rule(norm_tab(s))]
    if not hits:
        flag = "disability" if role == "dis" else role
        raise AuditInputError(f"Could not find the {label} tab. Tabs in workbook: {sheetnames}. "
                              f"Rename the tab or pass --{flag}-tab \"<tab name>\".")
    return hits


def school_calendar(calendar_path=None):
    """Unknown days stay unknown. Never infer instructional days from weekdays.

    Optional CSV: Date,IsSchoolDay, one row per calendar date, 1 or 0.
    Only the first/last school dates explicitly supplied by the user are known
    in the absence of a full approved instructional calendar.
    """
    days = {}
    d = FIRST_SCHOOL_DAY
    while d <= YEAR_END:
        days[d] = None
        d += timedelta(days=1)
    days[FIRST_SCHOOL_DAY] = 1
    days[YEAR_END] = 1
    if calendar_path:
        truthy, falsy = {"1", "y", "yes", "true", "x"}, {"0", "n", "no", "false"}
        with open(calendar_path, encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            fields = {norm_header(h): h for h in (reader.fieldnames or [])}
            if "date" not in fields or "isschoolday" not in fields:
                raise AuditInputError(f"Calendar CSV needs columns Date,IsSchoolDay; found {reader.fieldnames}")
            seen = set()
            for line, r in enumerate(reader, 2):
                raw_date = (r.get(fields["date"]) or "").strip()
                if not raw_date:
                    continue
                try:
                    d = parsed_date(raw_date)
                except ValueError:
                    raise AuditInputError(f"Calendar CSV line {line}: unreadable date {raw_date!r}") from None
                if d in seen:
                    raise AuditInputError(f"Calendar CSV line {line}: duplicate date {d}")
                seen.add(d)
                value = (r.get(fields["isschoolday"]) or "").strip().lower()
                if value.endswith(".0"):
                    value = value[:-2]
                if value not in truthy | falsy | {""}:
                    raise AuditInputError(f"Calendar CSV line {line}: IsSchoolDay must be 1, 0, or blank; got {value!r}")
                if d in days:
                    days[d] = None if value == "" else int(value in truthy)
    return days


def school_day_distance(earlier, later, days):
    """Count instructional dates AFTER earlier THROUGH later, within the year."""
    first, last = max(earlier + timedelta(days=1), FIRST_SCHOOL_DAY), min(later, YEAR_END)
    total = 0
    while first <= last:
        flag = days.get(first)
        if flag not in {0, 1}:
            return None
        total += flag
        first += timedelta(days=1)
    return total


def rule_result(status, classification, note, distance=None):
    return {"status": status, "classification": classification, "note": note, "school_days": distance}


def entry_rule(charter_entry, st_entry, days):
    c, s = parsed_date(charter_entry), st_date(st_entry)
    if s is None:
        return rule_result("REVIEW", "REVIEW - MISSING ST ENTRY", "Missing ST entry: review, not a confirmed date discrepancy")
    if c is None:
        return rule_result("REVIEW", "REVIEW - MISSING CHARTER ENTRY", "Missing charter entry: review")
    if s < FIRST_SCHOOL_DAY <= c:
        return rule_result("GOOD", "ST EARLY ENTRY / BEFORE SCHOOL STARTS", "ST early entry / before school starts")
    if c == s:
        return rule_result("GOOD", "MATCH", "Entry dates match")
    if s < c:
        return rule_result("GOOD", "FAVORABLE - ST ENTRY EARLIER", "ST entry earlier than charter entry")
    distance = school_day_distance(c, s, days)
    if distance is not None:
        if distance > 3:
            return rule_result("NOT GOOD", "UNFAVORABLE - ENTRY", "Charter entry earlier than ST entry", distance)
        return rule_result("GOOD", "WITHIN TOLERANCE", "Entry difference is no more than 3 school days", distance)
    # Even without a calendar, a <=3 calendar-day interval cannot contain more
    # than 3 school days. Do not invent an exact count for that interval.
    if (s - c).days <= 3:
        return rule_result("GOOD", "WITHIN TOLERANCE", "At most 3 calendar days, so at most 3 school days; exact count not asserted")
    return rule_result("REVIEW", "REVIEW - SCHOOL CALENDAR REQUIRED", "Exact school-day count needs approved instructional dates")


def exit_rule(charter_exit, st_exit, days):
    # Only 1990 placeholders are cleared under the governing Markdown rules.
    # June 27/June 30 dates are NOT silently changed to blank.
    c, s = st_date(charter_exit), st_date(st_exit)
    if c is None and s is None:
        return rule_result("GOOD", "NO EXIT / ACTIVE", "Both exit dates are blank")
    if c is None or s is None:
        return rule_result("REVIEW", "REVIEW - EXIT DATE MISSING FROM ONE SOURCE", "Exit date exists in only one source")
    if c == s:
        return rule_result("GOOD", "MATCH", "Exit dates match")
    if s > c:
        return rule_result("GOOD", "FAVORABLE - ST EXIT LATER", "ST exit later than charter exit")
    distance = school_day_distance(s, c, days)
    if distance is not None:
        if distance > 3:
            return rule_result("NOT GOOD", "UNFAVORABLE - EXIT", "Charter exit later than ST exit", distance)
        return rule_result("GOOD", "WITHIN TOLERANCE", "Exit difference is no more than 3 school days", distance)
    if (c - s).days <= 3:
        return rule_result("GOOD", "WITHIN TOLERANCE", "At most 3 calendar days, so at most 3 school days; exact count not asserted")
    return rule_result("REVIEW", "REVIEW - SCHOOL CALENDAR REQUIRED", "Exact school-day count needs approved instructional dates")


def overall_rule(entry, exit):
    if entry["status"] == "NOT GOOD" or exit["status"] == "NOT GOOD":
        return "NOT GOOD"
    if entry["status"] == "REVIEW" or exit["status"] == "REVIEW":
        return "REVIEW"
    return "GOOD"


def rule_self_tests():
    """Independent explicit expected cases; test dates are NOT a real calendar."""
    fixture = {}
    d = FIRST_SCHOOL_DAY
    while d <= YEAR_END:
        fixture[d] = int(d.weekday() < 5)
        d += timedelta(days=1)
    fixture[date(2026, 10, 12)] = 0  # Synthetic closure boundary test.
    unknown = school_calendar()
    tests = [
        ("Early July entry", entry_rule, "2026-09-03", "2026-07-01", fixture, "GOOD", "ST EARLY ENTRY / BEFORE SCHOOL STARTS"),
        ("Early August entry", entry_rule, "2026-09-03", "2026-08-28", fixture, "GOOD", "ST EARLY ENTRY / BEFORE SCHOOL STARTS"),
        ("Entry match", entry_rule, "2026-09-10", "2026-09-10", fixture, "GOOD", "MATCH"),
        ("ST earlier after first day", entry_rule, "2026-09-15", "2026-09-10", fixture, "GOOD", "FAVORABLE - ST ENTRY EARLIER"),
        ("Entry three school days", entry_rule, "2026-09-08", "2026-09-11", fixture, "GOOD", "WITHIN TOLERANCE"),
        ("Entry four school days", entry_rule, "2026-09-08", "2026-09-14", fixture, "NOT GOOD", "UNFAVORABLE - ENTRY"),
        ("Entry example five school days", entry_rule, "2026-09-08", "2026-09-15", fixture, "NOT GOOD", "UNFAVORABLE - ENTRY"),
        ("Missing ST entry", entry_rule, "2026-09-08", None, fixture, "REVIEW", "REVIEW - MISSING ST ENTRY"),
        ("Missing charter entry", entry_rule, None, "2026-09-08", fixture, "REVIEW", "REVIEW - MISSING CHARTER ENTRY"),
        ("Both entries blank", entry_rule, None, None, fixture, "REVIEW", "REVIEW - MISSING ST ENTRY"),
        ("Exit match", exit_rule, "2026-10-15", "2026-10-15", fixture, "GOOD", "MATCH"),
        ("ST exits later", exit_rule, "2026-10-15", "2026-10-20", fixture, "GOOD", "FAVORABLE - ST EXIT LATER"),
        ("Exit three school days", exit_rule, "2026-10-20", "2026-10-15", fixture, "GOOD", "WITHIN TOLERANCE"),
        ("Exit example unfavorable", exit_rule, "2026-10-23", "2026-10-15", fixture, "NOT GOOD", "UNFAVORABLE - EXIT"),
        ("Both exits blank", exit_rule, None, None, fixture, "GOOD", "NO EXIT / ACTIVE"),
        ("Missing ST exit", exit_rule, "2026-10-15", None, fixture, "REVIEW", "REVIEW - EXIT DATE MISSING FROM ONE SOURCE"),
        ("Missing charter exit", exit_rule, None, "2026-10-15", fixture, "REVIEW", "REVIEW - EXIT DATE MISSING FROM ONE SOURCE"),
        ("ST 1990 placeholder", exit_rule, None, "1990-01-01", fixture, "GOOD", "NO EXIT / ACTIVE"),
        ("Charter 1990 placeholder", exit_rule, "1990-01-01", None, fixture, "GOOD", "NO EXIT / ACTIVE"),
        ("No hidden year-end exception", exit_rule, "2027-06-27", None, fixture, "REVIEW", "REVIEW - EXIT DATE MISSING FROM ONE SOURCE"),
        ("Holiday within tolerance", entry_rule, "2026-10-09", "2026-10-15", fixture, "GOOD", "WITHIN TOLERANCE"),
        ("Unknown calendar entry", entry_rule, "2026-09-08", "2026-09-15", unknown, "REVIEW", "REVIEW - SCHOOL CALENDAR REQUIRED"),
        ("Unknown calendar exit", exit_rule, "2026-10-23", "2026-10-15", unknown, "REVIEW", "REVIEW - SCHOOL CALENDAR REQUIRED"),
        ("Short gap safely tolerated", entry_rule, "2026-09-08", "2026-09-10", unknown, "GOOD", "WITHIN TOLERANCE"),
    ]
    results = []
    for name, func, c, s, cal, status, classification in tests:
        got = func(c, s, cal)
        ok = got["status"] == status and got["classification"] == classification
        results.append((name, "PASS" if ok else "FAIL", status, classification))
    for e, x, expected in [("NOT GOOD", "REVIEW", "NOT GOOD"), ("REVIEW", "GOOD", "REVIEW"), ("GOOD", "GOOD", "GOOD"), ("GOOD", "NOT GOOD", "NOT GOOD")]:
        got = overall_rule({"status": e}, {"status": x})
        results.append((f"Overall {e}/{x}", "PASS" if got == expected else "FAIL", expected, "Overall priority rule"))
    if any(r[1] != "PASS" for r in results):
        raise RuntimeError(f"Rule self-test failed: {[r for r in results if r[1] != 'PASS']}")
    return results


def blank(value):
    return value is None or value == "" or (isinstance(value, float) and math.isnan(value))


def identifier(value):
    if blank(value):
        return ""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if float(value).is_integer():
            return str(int(value))
    return re.sub(r"\.0$", "", str(value).strip())


def valid_key(value):
    text = identifier(value)
    # A numeric Excel cell cannot store a leading zero, so restore it.
    if isinstance(value, (int, float)) and not isinstance(value, bool) and re.fullmatch(r"[0-9]{1,9}", text):
        text = text.zfill(10)
    return text if re.fullmatch(r"[0-9]{10}", text) else None


def parsed_date(value):
    if blank(value):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and 20000 < value < 60000:
        return from_excel(value).date()
    text = str(value).strip()
    try:
        return datetime.fromisoformat(text.replace("Z", "+00:00")).date()
    except ValueError:
        pass
    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d", "%Y%m%d", "%m-%d-%Y", "%d-%b-%Y", "%d%b%Y", "%b %d, %Y", "%B %d, %Y", "%m/%d/%Y %H:%M:%S", "%m/%d/%Y %I:%M:%S %p"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"Unparseable date: {text!r}")


def st_date(value):
    result = parsed_date(value)
    return None if result and result.year == 1990 else result


def normalize_name(value, first=False):
    text = "" if blank(value) else str(value).strip().lower()
    tokens = [x for x in text.split() if x.rstrip(".") not in {"jr", "sr", "ii", "iii", "iv"}]
    text = tokens[0] if first and tokens else "".join(tokens)
    text = "".join(c for c in unicodedata.normalize("NFKD", text) if not unicodedata.combining(c))
    return "".join(c for c in text if c.isalpha())


def name_dob(last, first, dob):
    last, first, dob = normalize_name(last), normalize_name(first, first=True), parsed_date(dob)
    return (last, first, dob) if last and first and dob else None


def read_table(ws, required, optional=(), students=False, search_rows=10):
    """Read a tab by header NAME, wherever the header row is (first 10 rows).

    Returns (headers, records, columns). Records are keyed by the canonical
    required/optional names, so source column order and header spelling
    (case, spacing, punctuation) do not matter. ``columns`` maps canonical
    name -> 1-based column number for live formulas. ``record['_values']``
    keeps the row exactly as positioned in the source.
    """
    if hasattr(ws, "reset_dimensions"):
        ws.reset_dimensions()  # Some exporters write a wrong <dimension>; read every cell.
    all_rows = list(ws.iter_rows(values_only=True))
    wanted = {norm_header(n): n for n in list(required) + list(optional)}
    for canon, aliases in COLUMN_ALIASES.items():
        if canon in wanted:
            for a in aliases:
                wanted.setdefault(a, wanted[canon])
    best = None
    for i, row in enumerate(all_rows[:search_rows]):
        found = {wanted[norm_header(v)] for v in row if norm_header(v) in wanted}
        if best is None or len(found & set(required)) > len(best[1] & set(required)):
            best = (i, found)
    if best is None or not set(required) <= best[1]:
        missing = sorted(set(required) - (best[1] if best else set()))
        raise AuditInputError(f"Tab {ws.title!r} is missing required column(s) {missing} "
                              f"(searched the first {search_rows} rows for a header row).")
    header_index = best[0]
    headers = [str(x).strip() if x is not None else "" for x in all_rows[header_index]]
    while headers and not headers[-1]:
        headers.pop()
    columns = {}
    for col, h in enumerate(headers, 1):
        canon = wanted.get(norm_header(h))
        if canon and canon not in columns:
            columns[canon] = col
    columns["_header_row"] = header_index + 1
    output = []
    for excel_row, values in enumerate(all_rows[header_index + 1:], header_index + 2):
        values = list(values[:len(headers)]) + [None] * max(0, len(headers) - len(values))
        record = {canon: values[col - 1] for canon, col in columns.items() if not canon.startswith("_")}
        record["_values"] = values
        record["_excel_row"] = excel_row
        if students:
            if blank(record.get("Last Name")) and blank(record.get("First Name")):
                continue
        elif all(blank(x) for x in values):
            continue
        output.append(record)
    return headers, output, columns


def clean_dates(records, fields, columns):
    """Turn unreadable dates into blanks (REVIEW later) instead of crashing."""
    problems = []
    for record in records:
        record["_bad_dates"] = []
        for field in fields:
            raw = record.get(field)
            try:
                parsed_date(raw)
            except ValueError:
                record["_bad_dates"].append(f"{field}={raw!r}")
                record[field] = None
                record["_values"][columns[field] - 1] = None
                problems.append((record["_excel_row"], field, raw))
    return problems


def numeric(value):
    if isinstance(value, bool) or blank(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).strip().replace(",", ""))
    except ValueError:
        return None


def weekly_student_weeks(start, end):
    """Port of this project's existing NYSED-labelled weekly calculation.

    Separate from school-day discrepancy testing. Review boundary cases and
    official calculator results before using differences for payment approval.
    """
    if start is None or end is None or end < start:
        return None
    days = (end - start).days + 1
    if days < 3:
        return 0
    if days == 3:
        return 1

    def weeks_in_month(a, b, apply_first_rule):
        month_end = date(a.year, a.month, calendar.monthrange(a.year, a.month)[1])
        monday = a - timedelta(days=a.weekday())
        count, first_week = 0, True
        while monday <= b:
            sunday = monday + timedelta(days=6)
            days_in = max(0, (min(b, sunday, month_end) - max(a, monday)).days + 1)
            if days_in >= 3 and (not first_week or not apply_first_rule or sunday <= month_end):
                count += 1
            first_week = False
            monday += timedelta(days=7)
        return min(count, 5 if a.month == 7 else 4)

    if (start.year, start.month) == (end.year, end.month):
        return weeks_in_month(start, end, True)
    last_of_start = date(start.year, start.month, calendar.monthrange(start.year, start.month)[1])
    first_of_end = date(end.year, end.month, 1)
    middle_months = (end.year - start.year) * 12 + end.month - start.month - 1
    return weeks_in_month(start, last_of_start, True) + max(0, middle_months) * 4 + weeks_in_month(first_of_end, end, False)


def select_student(record, by_key, by_name):
    raw_key = valid_key(record.get("NYSSIS ID"))
    candidates = by_key.get(raw_key, []) if raw_key else []
    method = "NYSSIS" if candidates else "NOT FOUND"
    if not candidates:
        key = name_dob(record.get("Last Name"), record.get("First Name"), record.get("DOB"))
        candidates = by_name.get(key, []) if key else []
        if candidates:
            method = "NAME+DOB - VERIFY"
    if not candidates:
        return None, method, "NOT FOUND - manual SchoolTool review", 0

    person_ids = {identifier(r.get("StudentID")) for r in candidates if identifier(r.get("StudentID"))}
    if len(person_ids) != 1:
        return None, "AMBIGUOUS", "AMBIGUOUS IDENTITY - do not auto-match", len(candidates)
    target = [r for r in candidates if "ontech" in str(r.get("SchoolName", "")).lower()]
    pool = target or candidates
    pool = sorted(pool, key=lambda r: (r["_start"] or date.min, r["_excel_row"]))
    selected = pool[-1]
    note = ""
    if not target:
        note = "OTHER SCHOOL - manual SchoolTool review"
    elif len(pool) > 1:
        note = "MULTIPLE ONTECH SEGMENTS - review history; not blindly merged"
    elif method != "NYSSIS":
        note = "NAME+DOB CANDIDATE - manual identity verification required"
    else:
        roster_dob = parsed_date(record.get("DOB"))
        st_dob = parsed_date(selected.get("Person_DOB"))
        same_name = (
            normalize_name(record.get("Last Name")) == normalize_name(selected.get("LastName"))
            and normalize_name(record.get("First Name"), first=True) == normalize_name(selected.get("FirstName"), first=True))
        if roster_dob and st_dob and roster_dob != st_dob:
            note = "NYSSIS MATCH / DOB CONFLICT - manual identity review"
        elif not same_name:
            note = "NYSSIS MATCH / NAME DIFFERENCE - verify spelling and identity"
    return selected, method, note, len(pool)


def analyze(record, by_key, by_name, disability_index, general_keys):
    selected, method, identity_note, segments = select_student(record, by_key, by_name)
    key = valid_key(record.get("NYSSIS ID"))
    dis_records = disability_index.get(key, []) if key else []
    dis_numbers = {identifier(r.get("StudentID_Disability")) for r in dis_records if identifier(r.get("StudentID_Disability"))}
    if not dis_records:
        dis_status = "NOT FOUND in current-year Disability export - review"
        dis_record = None
    elif len(dis_numbers) > 1:
        dis_status = "CONFLICTING export StudentID_Disability values - review"
        dis_record = None
    elif not dis_numbers:
        dis_status = "EXPORT MATCH, but StudentID_Disability is blank"
        dis_record = None
    else:
        dis_status = "CURRENT-YEAR EXPORT MATCH - active IEP not independently verified"
        dis_record = next(r for r in dis_records if identifier(r.get("StudentID_Disability")) in dis_numbers)

    ce, cx = parsed_date(record.get("Enroll Date")), parsed_date(record.get("Withdraw Date"))
    # Rollover normalization affects audit-only fields; raw cells are untouched.
    effective_cx = None if cx in {CHARTER_YEAR_END, YEAR_END} else cx
    se = selected["_start"] if selected else None
    sx = selected["_exit"] if selected else None
    if sx == YEAR_END:
        sx = None
    entry_gap = (se - ce).days if se and ce else None
    exit_gap = (effective_cx - sx).days if effective_cx and sx else None
    date_review = False
    if selected is None:
        entry_note = exit_note = "NOT CHECKED - no unambiguous ST match"
    else:
        if ce is None or se is None:
            entry_note, date_review = "MISSING ENTRY DATE - review", True
        elif ce >= se:
            entry_note = "Charter starts on/after ST; no extra entry days claimed"
        else:
            entry_note, date_review = "Charter starts before ST - verify SCHOOL-DAY difference", True
        if effective_cx is None and sx is None:
            exit_note = "Both active/year-end - no midyear exit mismatch"
        elif effective_cx is not None and (sx is None or effective_cx <= sx):
            exit_note = "FAVORABLE direction - charter reports no extra exit days"
        elif effective_cx is None and sx is not None:
            exit_note, date_review = "ST withdrawal / charter active - review invoice coverage and school days", True
        else:
            exit_note, date_review = "Charter exits after ST - verify SCHOOL-DAY difference", True
        if ce and effective_cx and effective_cx < ce:
            entry_note, date_review = "INVALID: charter exit before entry", True
        if se and sx and sx < se:
            exit_note, date_review = "INVALID: ST exit before entry", True

    bad = record.get("_bad_dates") or []
    if bad and not identity_note:
        identity_note = "UNREADABLE DATE IN ROSTER - " + "; ".join(bad)
    verification = identity_note or (
        "DATE REVIEW - instructional calendar/coverage confirmation needed" if date_review
        else "RECORDS ALIGNED - not final billing approval")
    source_fte = numeric(record.get("OnTECH FTE"))
    weeks = tentative_fte = variance = None
    if selected and se and "ontech" in str(selected.get("SchoolName", "")).lower():
        fte_start = max(se, FTE_FIRST_DAY)
        fte_end = sx or cx or YEAR_END
        weeks = weekly_student_weeks(fte_start, fte_end)
        tentative_fte = min(round(weeks / 40, 3), 1.0) if weeks is not None else None
        variance = round(source_fte - tentative_fte, 6) if source_fte is not None and tentative_fte is not None else None
    return {
        "source": record, "st": selected, "dis": dis_record,
        "method": method, "verification": verification, "segments": segments,
        "entry_note": entry_note, "exit_note": exit_note,
        "entry_gap": entry_gap, "exit_gap": exit_gap,
        "dis_status": dis_status, "dis_numbers": "; ".join(sorted(dis_numbers)),
        "in_general": "YES" if key and key in general_keys else "NO/UNRESOLVED - review",
        "charter_fte": source_fte, "weeks": weeks, "fte": tentative_fte, "variance": variance,
    }


def append_audit_sheets(source_path, output_path, audit_wb, style_mapping, cache,
                        part_prefix="ontech_audit"):
    """Append worksheet XML without rewriting any existing worksheet/component.

    Only workbook.xml, workbook.xml.rels and [Content_Types].xml must change to
    register new tabs. Existing worksheet XML, styles, shared strings, formulas,
    images, external links and all other original package parts stay byte-identical.
    """
    stream = io.BytesIO()
    audit_wb.save(stream)
    stream.seek(0)
    modified = {"xl/workbook.xml", "xl/_rels/workbook.xml.rels", "[Content_Types].xml"}
    with ZipFile(source_path) as source, ZipFile(stream) as additions:
        root = ET.fromstring(source.read("xl/workbook.xml"))
        rels = ET.fromstring(source.read("xl/_rels/workbook.xml.rels"))
        types = ET.fromstring(source.read("[Content_Types].xml"))
        sheets = root.find(f"{{{SS}}}sheets")
        existing = {s.get("name") for s in sheets}
        collisions = existing.intersection(audit_wb.sheetnames)
        if collisions:
            raise ValueError(f"Audit tabs already exist; use an original invoice input: {sorted(collisions)}")
        next_id = max(int(s.get("sheetId")) for s in sheets) + 1
        used_rids = {r.get("Id") for r in rels}
        new_parts = {}
        for index, ws in enumerate(audit_wb.worksheets, 1):
            part_name = f"xl/worksheets/{part_prefix}_{index}.xml"
            if part_name in source.namelist():
                raise ValueError(f"Reserved audit part already exists: {part_name}")
            rid = f"rId{part_prefix.replace('_', '')}{index}"
            if rid in used_rids:
                raise ValueError(f"Relationship collision: {rid}")
            ET.SubElement(sheets, f"{{{SS}}}sheet", name=ws.title, sheetId=str(next_id), attrib={f"{{{REL}}}id": rid})
            next_id += 1
            ET.SubElement(rels, f"{{{PKG_REL}}}Relationship", Id=rid, Type=REL + "/worksheet", Target=part_name[3:])
            ET.SubElement(types, f"{{{CT}}}Override", PartName="/" + part_name, ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
            tree = ET.fromstring(additions.read(f"xl/worksheets/sheet{index}.xml"))
            for cell in tree.iter(f"{{{SS}}}c"):
                local_style = int(cell.get("s", "0"))
                original_style = style_mapping.get(local_style, 0)
                if original_style:
                    cell.set("s", str(original_style))
                else:
                    cell.attrib.pop("s", None)
                result = cache.get((ws.title, cell.get("r")))
                if result is not None and cell.find(f"{{{SS}}}f") is not None:
                    for node in list(cell):
                        if node.tag in {f"{{{SS}}}v", f"{{{SS}}}is"}:
                            cell.remove(node)
                    if isinstance(result, (datetime, date)):
                        result = to_excel(result)
                    if isinstance(result, bool):
                        cell.set("t", "b")
                        result = int(result)
                    elif isinstance(result, (int, float)):
                        cell.attrib.pop("t", None)
                    else:
                        cell.set("t", "str")
                    ET.SubElement(cell, f"{{{SS}}}v").text = str(result)
            new_parts[part_name] = ET.tostring(tree, encoding="UTF-8", xml_declaration=True)
        calc = root.find(f"{{{SS}}}calcPr")
        if calc is None:
            calc = ET.SubElement(root, f"{{{SS}}}calcPr")
        calc.set("calcMode", "auto")
        calc.set("fullCalcOnLoad", "1")
        calc.set("forceFullCalc", "1")
        metadata = {
            "xl/workbook.xml": ET.tostring(root, encoding="UTF-8", xml_declaration=True),
            "xl/_rels/workbook.xml.rels": ET.tostring(rels, encoding="UTF-8", xml_declaration=True),
            "[Content_Types].xml": ET.tostring(types, encoding="UTF-8", xml_declaration=True),
        }
        with ZipFile(output_path, "w") as output:
            for info in source.infolist():
                output.writestr(info, metadata.get(info.filename, source.read(info.filename)))
            for name, data in new_parts.items():
                output.writestr(name, data, compress_type=8)

    results = []
    with ZipFile(source_path) as source, ZipFile(output_path) as output:
        if output.testzip() is not None:
            raise RuntimeError("Output ZIP CRC validation failed")
        for info in source.infolist():
            original, final = source.read(info.filename), output.read(info.filename)
            if info.filename not in modified and original != final:
                raise RuntimeError(f"Original package part changed: {info.filename}")
            results.append({"part": info.filename, "status": "METADATA UPDATED" if info.filename in modified else "BYTE-IDENTICAL", "source_sha256": hashlib.sha256(original).hexdigest(), "output_sha256": hashlib.sha256(final).hexdigest()})
        for name in output.namelist():
            if name.endswith(".xml") or name.endswith(".rels"):
                ET.fromstring(output.read(name))
    return results


def is_current_year(value):
    """Accept 2026-2027, 2026-27, '2026 - 2027', etc."""
    return bool(re.fullmatch(r"\s*2026\s*[-/–]\s*(2027|27)\s*", str(value or "")))


def style_id(ws, row, col):
    try:
        cell = ws.cell(row, col)
    except Exception:
        return 0
    return getattr(cell, "_style_id", None) or getattr(cell, "style_id", 0) or 0


def run(source_path, output_path, calendar_path=None, tab_overrides=None):
    source_path, output_path = Path(source_path), Path(output_path)
    tab_overrides = tab_overrides or {}
    if not source_path.is_file():
        raise AuditInputError(f"Invoice workbook not found: {source_path}")
    if source_path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise AuditInputError(f"Invoice must be an .xlsx or .xlsm workbook, not {source_path.suffix or 'no extension'}")
    if calendar_path and not Path(calendar_path).is_file():
        raise AuditInputError(f"Calendar CSV not found: {calendar_path}")
    if output_path.suffix.lower() != source_path.suffix.lower():
        # The output copies the input package, so it must keep the same type
        # (Excel refuses a macro workbook saved with an .xlsx extension).
        output_path = output_path.with_suffix(source_path.suffix.lower())
    if source_path.resolve() == output_path.resolve():
        raise AuditInputError("Input and output paths must be different")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    source_hash = hashlib.sha256(source_path.read_bytes()).hexdigest()
    try:
        wb = load_workbook(source_path, read_only=True, data_only=True, keep_links=True)
    except Exception as exc:
        raise AuditInputError(f"Could not open {source_path.name} as an Excel workbook ({exc}). "
                              "If it is password-protected or an old .xls file, re-save it as .xlsx.") from None
    already = GENERATED_TABS.intersection(wb.sheetnames)
    if already:
        wb.close()
        raise AuditInputError(f"This workbook already contains audit tabs {sorted(already)}. "
                              "Run the script on the ORIGINAL invoice, not a previous audit output.")

    tabs = {}
    for role in ("gen", "sped"):
        hits = find_tab(wb.sheetnames, role, tab_overrides.get(role))
        if len(hits) > 1:
            raise AuditInputError(f"More than one {TAB_RULES[role][0]} tab found: {hits}. "
                                  f"Pass --{role}-tab \"<tab name>\" to choose one.")
        tabs[role] = hits[0]
    gen_headers, gen, gen_cols = read_table(wb[tabs["gen"]], ROSTER_COLUMNS, students=True)
    sped_headers, sped, sped_cols = read_table(wb[tabs["sped"]], ROSTER_COLUMNS, students=True)

    def load_source(role, required, optional, year_field):
        """Pick the tab with current-year rows when several candidates exist."""
        hits = find_tab(wb.sheetnames, role, tab_overrides.get(role))
        loaded = []
        for name in hits:
            try:
                headers, rows, cols = read_table(wb[name], required, optional)
            except AuditInputError:
                if len(hits) == 1:
                    raise
                continue
            loaded.append((name, headers, rows, cols, sum(is_current_year(r.get(year_field)) for r in rows)))
        usable = [x for x in loaded if x[4]]
        if not loaded:
            raise AuditInputError(f"None of the {TAB_RULES[role][0]} tabs {hits} has the required columns {required}")
        if len(usable) > 1:
            flag = "disability" if role == "dis" else role
            raise AuditInputError(f"Several {TAB_RULES[role][0]} tabs have {YEAR} rows: {[x[0] for x in usable]}. "
                                  f"Pass --{flag}-tab \"<tab name>\" to choose one.")
        return (usable or loaded)[0]

    tabs["st"], _, st, st_cols, _ = load_source("st", ST_COLUMNS, ST_OPTIONAL, "StudentEnrollment_SchoolYear")
    tabs["dis"], _, disability, dis_cols, _ = load_source("dis", DIS_COLUMNS, (), "SchoolYear_Disability")

    date_problems = []
    for role, records, cols, fields in [("gen", gen, gen_cols, ROSTER_DATE_FIELDS), ("sped", sped, sped_cols, ROSTER_DATE_FIELDS),
                                        ("st", st, st_cols, ST_DATE_FIELDS)]:
        date_problems += [(tabs[role], *p) for p in clean_dates(records, fields, cols)]
    gen_ws = wb[tabs["gen"]]
    gen_header_row = gen_cols["_header_row"]
    first_data = gen[0]["_excel_row"] if gen else gen_header_row + 1
    original_style_ids = {"header": style_id(gen_ws, gen_header_row, 1),
                          "date": style_id(gen_ws, first_data, gen_cols["Enroll Date"]),
                          "fte": style_id(gen_ws, first_data, gen_cols["OnTECH FTE"])}
    wb.close()
    all_st_years = Counter(str(r.get("StudentEnrollment_SchoolYear")) for r in st)
    all_dis_years = Counter(str(r.get("SchoolYear_Disability")) for r in disability)
    st = [r for r in st if is_current_year(r.get("StudentEnrollment_SchoolYear"))]
    disability = [r for r in disability if is_current_year(r.get("SchoolYear_Disability"))]
    if not st:
        raise AuditInputError(f"No {YEAR} rows in {tabs['st']!r}. School years found: {dict(all_st_years)}")
    if not disability:
        raise AuditInputError(f"No {YEAR} rows in {tabs['dis']!r}. School years found: {dict(all_dis_years)}")

    by_key, by_name, disability_index = defaultdict(list), defaultdict(list), defaultdict(list)
    for record in st:
        record["_start"] = st_date(record.get("StudentEnrollment_StartDate"))
        record["_exit"] = st_date(record.get("StudentEnrollment_EndDate"))
        key = valid_key(record.get("Student_UniversalStudentID"))
        if key:
            by_key[key].append(record)
        nd = name_dob(record.get("LastName"), record.get("FirstName"), record.get("Person_DOB"))
        if nd:
            by_name[nd].append(record)
    for record in disability:
        key = valid_key(record.get("Student_UniversalStudentID"))
        if key:
            disability_index[key].append(record)
    general_keys = {key for r in gen if (key := valid_key(r.get("NYSSIS ID")))}
    gen_results = [analyze(r, by_key, by_name, disability_index, general_keys) for r in gen]
    sped_results = [analyze(r, by_key, by_name, disability_index, general_keys) for r in sped]
    calendar_flags = school_calendar(calendar_path)
    synthetic_tests = rule_self_tests()
    for result in gen_results + sped_results:
        record, selected = result["source"], result["st"]
        e = entry_rule(record.get("Enroll Date"), selected.get("StudentEnrollment_StartDate") if selected else None, calendar_flags)
        x = exit_rule(record.get("Withdraw Date"), selected.get("StudentEnrollment_EndDate") if selected else None, calendar_flags)
        safeguard = ""
        if selected and "ontech" not in str(selected.get("SchoolName", "")).lower():
            safeguard = "ST RECORD AT ANOTHER SCHOOL"
        elif result["method"] in {"NAME+DOB - VERIFY", "AMBIGUOUS"}:
            safeguard = "IDENTITY VERIFICATION REQUIRED"
        elif result["segments"] > 1:
            safeguard = "MULTIPLE SEGMENTS REQUIRE REVIEW"
        elif selected and "DOB CONFLICT" in result["verification"]:
            safeguard = "DOB CONFLICT REQUIRES REVIEW"
        elif record.get("_bad_dates"):
            safeguard = "UNREADABLE ROSTER DATE"
        if safeguard:
            e = rule_result("REVIEW", "REVIEW - " + safeguard, safeguard)
            x = rule_result("REVIEW", "REVIEW - " + safeguard, safeguard)
        previous=result['verification']
        identity_note=(previous if not previous.startswith(('RECORDS ALIGNED','DATE REVIEW')) else '')
        result.update(entry_logic=e, exit_logic=x, overall=overall_rule(e, x), safeguard=safeguard,
                      identity_note=identity_note)
        result['verification']=identity_note or 'MATCH CHECKS CLEAR; see live Overall Status'

    audit = Workbook()
    audit.remove(audit.active)
    cache = {}
    # Style markers are remapped to existing source styles when XML is appended.
    def header(cell):
        cell.font = Font(bold=True)
    def dated(cell):
        cell.number_format = "mm/dd/yyyy"
    def decimal(cell):
        cell.number_format = "0.000000"
    config = audit.create_sheet("AUDIT CONFIG")
    config.append(["Setting", "Value"])
    config.append(["First Day of School (user supplied)", FIRST_SCHOOL_DAY])
    config.append(["Last Day of School (user supplied)", YEAR_END])
    config.append(["Grace School Days", 3])
    config.append(["Governing enrollment logic", "charter_invoice_discrepancy_logic.md"])
    config.append(["Missing dates", "REVIEW, not a confirmed unfavorable discrepancy"])
    config.append(["1990 exits", "Normalized to blank in audit fields only"])
    config.append(["Other year-end dates", "Preserved as actual dates; no unapproved active/year-end exemption"])
    config.append(["Calendar basis", str(calendar_path) if calendar_path else "Only first/last school dates supplied; other days unknown"])
    config.append(["Three-school-day interval", "Count after the earlier date through the later date, within configured year"])
    config.append(["FTE scope", "Tentative weekly project method; not specified by the governing enrollment Markdown"])
    for row in (2, 3): dated(config.cell(row, 2))
    cal_ws = audit.create_sheet("SCHOOL CALENDAR")
    cal_ws.append(["Date", "IsSchoolDay (1/0; blank = unknown)", "Known Flag (formula)"])
    for r, (d, flag) in enumerate(calendar_flags.items(), 2):
        dated(cal_ws.cell(r, 1, d))
        cal_ws.cell(r, 2, flag)
        cell = cal_ws.cell(r, 3, f'=IF(AND(ISNUMBER(B{r}),OR(B{r}=0,B{r}=1)),1,0)')
        cache[(cal_ws.title, cell.coordinate)] = int(flag in {0, 1})
    cal_ws.freeze_panes = "B2"
    cal_ws.column_dimensions['A'].width = 18
    cal_ws.column_dimensions['B'].width = 38
    cal_ws.column_dimensions['C'].width = 24
    calendar_last = cal_ws.max_row
    def live(ws, row, col, sheet, source_col, source_row, value, date_field=False):
        if not source_row or not source_col:
            ws.cell(row, col, "")
            return
        if date_field and not blank(value) and not isinstance(value, (datetime, date, int, float)):
            # Text date in the export: Excel cannot read it reliably in every locale,
            # so write the parsed value instead of a live formula.
            dated(ws.cell(row, col, st_date(value) or ""))
            return
        escaped = sheet.replace("'", "''")
        index = f"INDEX('{escaped}'!${source_col}:${source_col},{source_row})"
        expr = f'IF({index}=0,"",{index})'
        if date_field:
            expr = f'IFERROR(IF({index}=0,"",IF(YEAR({index})=1990,"",INT({index}))),"")'
        cell = ws.cell(row, col, "=" + expr)
        cache[(ws.title, cell.coordinate)] = (st_date(value) or "") if date_field else ("" if blank(value) else value)
        if date_field:
            dated(cell)

    extra_headers = ["Original Excel Row", "Match Method", "Verification Assessment", "Selected ST Excel Row", "ST School (LIVE)", "ST Entry (LIVE)", "ST Exit (LIVE)", "ST Student ID (LIVE, source StudentID)", "ST Entry Note (LIVE)", "ST Exit Note (LIVE)", "Entry Date Verification", "Exit Date Verification", "Entry Calendar-Day Gap (NOT school days)", "Exit Calendar-Day Gap (NOT school days)", "Current Disability Number (LIVE, source StudentID_Disability)", "Disability Export Match Assessment", "On General Roster by NYSSIS", "Project Weekly Student Weeks (tentative)", "Project Weekly FTE (tentative)", "Charter minus Tentative FTE", "ST Segment Count"]
    extra_headers[2] = "Match / Data-Quality Assessment (snapshot)"
    extra_headers += ["Entry Status (LIVE)", "Entry Classification (LIVE)", "Entry School-Day Difference (LIVE)",
                      "Exit Status (LIVE)", "Exit Classification (LIVE)", "Exit School-Day Difference (LIVE)",
                      "Overall Status (LIVE)", "Match Safeguard (snapshot)", "Normalized Charter Exit (LIVE)",
                      "Entry Gap Start (LIVE helper)", "Entry Gap End (LIVE helper)", "Entry Calendar Complete (LIVE)",
                      "Exit Gap Start (LIVE helper)", "Exit Gap End (LIVE helper)", "Exit Calendar Complete (LIVE)",
                      "Notes for Patrick (LIVE)"]

    def write_rule_formulas(ws, row, cols, offset, result):
        """Write tested live versions of the governing Markdown decisions."""
        def ref(n): return f"{get_column_letter(offset+n)}{row}"
        def quoted(text): return '"' + text.replace('"', '""') + '"'
        def choose(cases, default):
            value = default
            for condition, outcome in reversed(cases):
                value = f"IF({condition},{outcome},{value})"
            return value
        def formula(n, text, value, is_date=False):
            cell = ws.cell(row, offset+n, "="+text)
            cache[(ws.title,cell.coordinate)] = "" if value is None else value
            if is_date: dated(cell)
        ce = f"{get_column_letter(cols['Enroll Date'])}{row}"
        cx_raw = f"{get_column_letter(cols['Withdraw Date'])}{row}"
        se, sx = ref(6), ref(7)
        es, ec, ed, xs, xc, xd, ov, guard, cx = [ref(n) for n in range(22,31)]
        eas, eae, eok, xas, xae, xok = [ref(n) for n in range(31,37)]
        ws.cell(row, offset+29, result['safeguard'])
        raw_charter_exit = st_date(result['source'].get('Withdraw Date'))
        formula(30, f'IF({cx_raw}="","",IF(YEAR({cx_raw})=1990,"",INT({cx_raw})))', raw_charter_exit, True)
        record=result['source']; selected=result['st']
        py_ce=parsed_date(record.get('Enroll Date'))
        py_se=st_date(selected.get('StudentEnrollment_StartDate')) if selected else None
        py_sx=st_date(selected.get('StudentEnrollment_EndDate')) if selected else None
        date_range=f"'SCHOOL CALENDAR'!$A$2:$A${calendar_last}"
        flag_range=f"'SCHOOL CALENDAR'!$B$2:$B${calendar_last}"
        known_range=f"'SCHOOL CALENDAR'!$C$2:$C${calendar_last}"
        for gap,start_ref,end_ref,ok_ref,n_start,n_end,n_ok,n_days,a,b,condition in [
            ('ENTRY',eas,eae,eok,31,32,33,24,py_ce,py_se,f'AND({ce}<>"",{se}<>"",{se}>{ce})'),
            ('EXIT',xas,xae,xok,34,35,36,27,py_sx,raw_charter_exit,f'AND({sx}<>"",{cx}<>"",{cx}>{sx})')]:
            earlier_ref,later_ref=(ce,se) if gap=='ENTRY' else (sx,cx)
            a_gap=max(a+timedelta(days=1),FIRST_SCHOOL_DAY) if a and b and b>a else None
            b_gap=min(b,YEAR_END) if a and b and b>a else None
            counted=school_day_distance(a,b,calendar_flags) if a and b and b>a else None
            complete=counted is not None if a and b and b>a else False
            formula(n_start,f'IF({condition},MAX({earlier_ref}+1,\'AUDIT CONFIG\'!$B$2),"")',a_gap,True)
            formula(n_end,f'IF({condition},MIN({later_ref},\'AUDIT CONFIG\'!$B$3),"")',b_gap,True)
            expected=f'MAX(0,{end_ref}-{start_ref}+1)'
            count_rows=f'COUNTIFS({date_range},">="&{start_ref},{date_range},"<="&{end_ref})'
            known=f'SUMIFS({known_range},{date_range},">="&{start_ref},{date_range},"<="&{end_ref})'
            formula(n_ok,f'IF(OR({start_ref}="",{end_ref}=""),FALSE,AND({count_rows}={expected},{known}={expected}))',complete)
            school_sum=f'SUMIFS({flag_range},{date_range},">="&{start_ref},{date_range},"<="&{end_ref})'
            formula(n_days,f'IF({ok_ref},{school_sum},"")',counted if complete else None)
        entry_text=choose([
            (f'{guard}<>""',f'"REVIEW - "&{guard}'),
            (f'{se}=""',quoted('REVIEW - MISSING ST ENTRY')),
            (f'{ce}=""',quoted('REVIEW - MISSING CHARTER ENTRY')),
            (f'AND({se}<\'AUDIT CONFIG\'!$B$2,{ce}>=\'AUDIT CONFIG\'!$B$2)',quoted('ST EARLY ENTRY / BEFORE SCHOOL STARTS')),
            (f'{se}={ce}',quoted('MATCH')),
            (f'{se}<{ce}',quoted('FAVORABLE - ST ENTRY EARLIER')),
            (f'AND({eok},{ed}>\'AUDIT CONFIG\'!$B$4)',quoted('UNFAVORABLE - ENTRY')),
            (f'OR({eok},{se}-{ce}<=\'AUDIT CONFIG\'!$B$4)',quoted('WITHIN TOLERANCE')),
        ],quoted('REVIEW - SCHOOL CALENDAR REQUIRED'))
        exit_text=choose([
            (f'{guard}<>""',f'"REVIEW - "&{guard}'),
            (f'AND({cx}="",{sx}="")',quoted('NO EXIT / ACTIVE')),
            (f'OR({cx}="",{sx}="")',quoted('REVIEW - EXIT DATE MISSING FROM ONE SOURCE')),
            (f'{cx}={sx}',quoted('MATCH')),
            (f'{sx}>{cx}',quoted('FAVORABLE - ST EXIT LATER')),
            (f'AND({xok},{xd}>\'AUDIT CONFIG\'!$B$4)',quoted('UNFAVORABLE - EXIT')),
            (f'OR({xok},{cx}-{sx}<=\'AUDIT CONFIG\'!$B$4)',quoted('WITHIN TOLERANCE')),
        ],quoted('REVIEW - SCHOOL CALENDAR REQUIRED'))
        formula(23,entry_text,result['entry_logic']['classification'])
        formula(26,exit_text,result['exit_logic']['classification'])
        for n,cref,expected in [(22,ec,result['entry_logic']['status']),(25,xc,result['exit_logic']['status'])]:
            formula(n,f'IF(LEFT({cref},11)="UNFAVORABLE","NOT GOOD",IF(LEFT({cref},6)="REVIEW","REVIEW","GOOD"))',expected)
        formula(28,f'IF(OR({es}="NOT GOOD",{xs}="NOT GOOD"),"NOT GOOD",IF(OR({es}="REVIEW",{xs}="REVIEW"),"REVIEW","GOOD"))',result['overall'])
        note=f"{result['entry_logic']['classification']} | {result['exit_logic']['classification']}"
        formula(37,f'{ec}&" | "&{xc}',note)
        formula(11,ec,result['entry_logic']['classification'])
        formula(12,xc,result['exit_logic']['classification'])

    review_rows = []
    st_letter = {name: get_column_letter(c) for name, c in st_cols.items() if not name.startswith("_")}
    dis_letter = get_column_letter(dis_cols["StudentID_Disability"])
    roster_date_positions = {"gen": {gen_cols[f] for f in ROSTER_DATE_FIELDS}, "sped": {sped_cols[f] for f in ROSTER_DATE_FIELDS}}
    for title, headers, results, source_title, role, cols in [
            (AUDIT_NAMES[0], gen_headers, gen_results, tabs["gen"], "gen", gen_cols),
            (AUDIT_NAMES[1], sped_headers, sped_results, tabs["sped"], "sped", sped_cols)]:
        ws = audit.create_sheet(title)
        for col, name in enumerate([f"SOURCE: {h}" for h in headers] + extra_headers, 1):
            header(ws.cell(1, col, name))
            ws.column_dimensions[get_column_letter(col)].width = min(48, max(18, len(name) + 2))
        offset = len(headers)
        for row_no, result in enumerate(results, 2):
            record, selected = result["source"], result["st"]
            for col, value in enumerate(record["_values"], 1):
                if col in roster_date_positions[role] and not blank(value):
                    value = parsed_date(value)  # text dates become real dates so formulas compare correctly
                cell = ws.cell(row_no, col, value)
                if isinstance(value, (datetime, date)):
                    dated(cell)
            extra = [record["_excel_row"], result["method"], result["verification"], selected["_excel_row"] if selected else None,
                     None, None, None, None, None, None, result["entry_note"], result["exit_note"], result["entry_gap"], result["exit_gap"], None,
                     result["dis_status"], result["in_general"], result["weeks"], result["fte"], result["variance"], result["segments"]]
            for col, value in enumerate(extra, offset + 1):
                ws.cell(row_no, col, value)
            sr = selected["_excel_row"] if selected else None
            for new_offset, old_name, date_flag in [
                    (5, "SchoolName", False), (6, "StudentEnrollment_StartDate", True),
                    (7, "StudentEnrollment_EndDate", True), (8, "StudentID", False),
                    (9, "StudentEnrollment_EntryNote", False), (10, "StudentEnrollment_ExitNote", False)]:
                live(ws, row_no, offset + new_offset, tabs["st"], st_letter.get(old_name), sr,
                     selected.get(old_name) if selected else None, date_field=date_flag)
            dr = result["dis"]["_excel_row"] if result["dis"] else None
            live(ws, row_no, offset + 15, tabs["dis"], dis_letter, dr,
                 result["dis"].get("StudentID_Disability") if result["dis"] else None)
            decimal(ws.cell(row_no, offset + 19))
            decimal(ws.cell(row_no, offset + 20))
            write_rule_formulas(ws,row_no,cols,offset,result)
            if (result['overall'] != 'GOOD' or bool(result['identity_note'])
                    or (role == "sped" and (not result["dis"] or result["in_general"] != "YES"))
                    or (result["variance"] is not None and abs(result["variance"]) > 0.0001)):
                review_rows.append([source_title, record["_excel_row"], identifier(record.get("NYSSIS ID")), record.get("Last Name"), record.get("First Name"), result["method"], result['overall'], result['entry_logic']['classification'], result['exit_logic']['classification'], result["dis_status"] if role == "sped" else "Not a billed-SPED eligibility determination", result["variance"]])
        ws.freeze_panes = "C2"
        ws.auto_filter.ref = ws.dimensions

    review = audit.create_sheet(AUDIT_NAMES[2])
    review.append(["Source Roster", "Source Excel Row", "NYSSIS (source)", "Last Name", "First Name", "Match Method", "Verification", "Entry Review", "Exit Review", "Disability Review", "Tentative FTE Variance"])
    for values in review_rows:
        review.append(values)
    review.freeze_panes = "A2"
    review.auto_filter.ref = review.dimensions

    summary = audit.create_sheet(AUDIT_NAMES[3])
    metrics = {"detected_tabs": tabs, "unreadable_dates": len(date_problems), "school_year": YEAR, "gen_rows": len(gen), "sped_rows": len(sped), "gen_source_fte": round(sum(r["charter_fte"] or 0 for r in gen_results), 6), "sped_source_fte": round(sum(r["charter_fte"] or 0 for r in sped_results), 6), "st_current_year_rows": len(st), "disability_current_year_rows": len(disability), "review_rows": len(review_rows)}
    summary.append(["Metric", f"General Roster ({tabs['gen']})", f"SPED Roster ({tabs['sped']})"])
    summary.append(["Source student rows", len(gen), len(sped)])
    summary.append(["Sum of source OnTECH FTE", metrics["gen_source_fte"], metrics["sped_source_fte"]])
    metrics['rule_tests_passed'] = len(synthetic_tests)
    metrics['calendar_known_days'] = sum(v in {0,1} for v in calendar_flags.values())
    metrics['calendar_total_dates'] = len(calendar_flags)
    metrics['overall_status_counts'] = {}
    for status in ['GOOD','NOT GOOD','REVIEW']:
        counts=[]
        output_row=summary.max_row+1
        summary.cell(output_row,1,'OVERALL STATUS: '+status)
        for col,title,headers,results in [(2,AUDIT_NAMES[0],gen_headers,gen_results),(3,AUDIT_NAMES[1],sped_headers,sped_results)]:
            count=sum(r['overall']==status for r in results)
            overall_column=get_column_letter(len(headers)+28)
            expression=f'=COUNTIF(\'{title}\'!{overall_column}2:{overall_column}{len(results)+1},"{status}")'
            cell=summary.cell(output_row,col,expression)
            cache[(summary.title,cell.coordinate)]=count
            counts.append(count)
        metrics['overall_status_counts'][status]={'general':counts[0],'sped':counts[1]}
    for name, test in [
            ("Exact NYSSIS matches", lambda r: r["method"] == "NYSSIS"),
            ("Name+DOB candidates - verification needed", lambda r: r["method"] == "NAME+DOB - VERIFY"),
            ("Not found", lambda r: r["method"] == "NOT FOUND"),
            ("Ambiguous identity", lambda r: r["method"] == "AMBIGUOUS"),
            ("Selected ST school is OnTech", lambda r: r["st"] is not None and "ontech" in str(r["st"].get("SchoolName", "")).lower()),
            ("Selected ST school is not OnTech", lambda r: r["st"] is not None and "ontech" not in str(r["st"].get("SchoolName", "")).lower()),
            ("Identity/record-selection review", lambda r: bool(r['identity_note'])),
            ("Entry status GOOD under governing rules", lambda r: r['entry_logic']['status']=='GOOD'),
            ("Exit status GOOD under governing rules", lambda r: r['exit_logic']['status']=='GOOD'),
            ("Current Disability export / unique number match", lambda r: r["dis"] is not None),
            ("Nonzero tentative weekly FTE variance", lambda r: r["variance"] is not None and abs(r["variance"]) > 0.0001)]:
        a, b = sum(test(r) for r in gen_results), sum(test(r) for r in sped_results)
        summary.append([name, a, b])
        metrics[name] = {"general": a, "sped": b}
    for label, text in [
            ("Mode", "OFFLINE: uploaded exports only; no SQL Server access"),
            ("Active IEP / disability code / disability start and end", "NOT AVAILABLE from these three Disability export columns"),
            ("Governing enrollment logic", "charter_invoice_discrepancy_logic.md; >3 SCHOOL days; missing dates REVIEW"),
            ("Three-school-day discrepancy rule", "Implemented in Python and live formulas; unknown calendar days produce REVIEW rather than guessed counts"),
            ("Weekly FTE", "Tentative project algorithm, 40-week denominator; not a payment approval"),
            ("Charter year-end convention", "No hidden year-end exception: a June 27 charter exit and blank ST exit is REVIEW under the supplied Markdown"),
            ("Live fields", "Selected source-row INDEX formulas; rerun after sorting/inserting/deleting source rows or changing match keys"),
            ("Assessment fields", "Entry, exit, overall and Notes for Patrick are live; matching safeguards, review queue membership and tentative FTE are run snapshots"),
            ("Original formulas", "Some existing source VLOOKUPs use approximate TRUE. They remain unchanged, not endorsed; use new verification tabs"),
            ("Source workbook SHA-256", source_hash),
            ("ST export school years", json.dumps(all_st_years)),
            ("Disability export school years", json.dumps(all_dis_years))]:
        summary.append([label, text])

    checks = audit.create_sheet(AUDIT_NAMES[4])
    checks.append(["Validation Check", "Result", "Details"])
    for role, label in [("gen", "General roster tab"), ("sped", "SPED roster tab"), ("st", "SchoolTool enrollment tab"), ("dis", "Disability tab")]:
        checks.append([label, "DETECTED", tabs[role]])
    checks.append(["Unreadable dates", "PASS" if not date_problems else f"REVIEW - {len(date_problems)} found",
                   "None" if not date_problems else "Treated as blank (REVIEW); source cells unchanged"])
    for tab_name, excel_row, field, raw in date_problems:
        checks.append([f"  {tab_name} row {excel_row}", field, repr(raw)])
    checks.append(["GEN rows retained in verification", "PASS", len(gen_results)])
    checks.append(["SPED rows retained in verification", "PASS", len(sped_results)])
    checks.append(["Source values / IDs", "UNCHANGED", "All original worksheet package parts remain byte-identical"])
    checks.append(["Input workbook", "READ ONLY", "SHA-256 checked after completion"])
    checks.append(["SchoolTool export year", "PASS", YEAR])
    checks.append(["Disability export year", "PASS", YEAR])
    checks.append(["No credentials / network", "PASS", "This script contains no passwords or database connection code"])
    checks.append(["IEP eligibility determination", "NOT PERFORMED", "Not supported by export fields"])
    checks.append(["Original package bytes", "CHECKED ON WRITE", "All parts except three required workbook-registration metadata files"])
    with ZipFile(source_path) as z:
        for info in z.infolist():
            status = "METADATA UPDATED" if info.filename in {"xl/workbook.xml", "xl/_rels/workbook.xml.rels", "[Content_Types].xml"} else "BYTE-IDENTICAL"
            checks.append([info.filename, status, hashlib.sha256(z.read(info.filename)).hexdigest()])

    validation = audit.create_sheet('VALIDATION REPORT')
    validation.append(['Validation Check','Result','Details'])
    validation.append(['Governing rule file','APPLIED','charter_invoice_discrepancy_logic.md; replaces conflicting prior calendar-day and missing-date rules'])
    validation.append(['First/last school date','USER CONFIRMED','2026-09-03 through 2027-06-30'])
    validation.append(['Classification tests','PASS',f'{len(synthetic_tests)} explicit cases; all passed before output'])
    validation.append(['Live enrollment rules','IMPLEMENTED','Entry status/classification, exit status/classification, overall status, and notes are live formulas'])
    validation.append(['Calendar completeness','PASS' if all(v in {0,1} for v in calendar_flags.values()) else 'REVIEW - CALENDAR INCOMPLETE',f'{metrics["calendar_known_days"]} of {metrics["calendar_total_dates"]} dates known; unknown differences are not guessed'])
    validation.append(['Short-gap mathematical bound','APPLIED','No more than 3 calendar dates implies no more than 3 school dates, without asserting an exact school-day count'])
    validation.append(['Missing-date rule','APPLIED','Missing entry, or only one exit, is REVIEW. Both exits blank is GOOD/NO EXIT/ACTIVE'])
    validation.append(['1990 placeholders','APPLIED','Only audit fields normalized; original cells remain unchanged'])
    validation.append(['No unapproved year-end normalization','APPLIED','June 27/June 30 are not automatically changed to blank by the new enrollment rules'])
    validation.append(['Source roster rows','PASS',f'{len(gen)} general; {len(sped)} SPED; no student removed'])
    validation.append(['Source FTE values','PRESERVED',f'General={metrics["gen_source_fte"]}; SPED={metrics["sped_source_fte"]}'])
    validation.append(['Source package integrity','CHECKED ON WRITE','All original parts except the three workbook-registration metadata files must remain byte-identical; code raises on any change'])
    validation.append(['Identity and other-school safeguards','APPLIED','Ambiguous/name+DOB/DOB conflict/other-school/multiple-segment cases remain REVIEW instead of automatic approval'])
    validation.append(['Active IEP confirmation','NOT DETERMINED','Current Disability export presence is not independent active-IEP evidence'])
    validation.append(['FTE scope','SEPARATE / TENTATIVE','Enrollment Markdown does not specify an FTE algorithm; project weekly FTE remains a static, labelled calculation'])
    validation.append(['Queue refresh','SNAPSHOT','Live classifications update; re-run to rebuild review-queue membership after edits'])
    validation.append(['Full old V4 manual compliance','NOT CLAIMED','This build implements the new governing enrollment rules, not every conflicting old manual feature'])
    validation.append([])
    validation.append(['Synthetic rule test','Result','Expected outcome / classification'])
    for name,status,expected,classification in synthetic_tests:
        validation.append([name,status,expected+' / '+classification])

    for ws in [review, summary, checks,validation,config]:
        for cell in ws[1]:
            header(cell)
        for col in range(1, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(col)].width = 32 if col != 2 else 58
    # Force style registration before translating generated worksheet style IDs.
    header_style = audit[AUDIT_NAMES[0]].cell(1, 1).style_id
    date_style = None
    fte_style = audit[AUDIT_NAMES[0]].cell(2, len(gen_headers) + 19).style_id
    for row in audit[AUDIT_NAMES[0]].iter_rows():
        for cell in row:
            if cell.number_format == "mm/dd/yyyy":
                date_style = cell.style_id
                break
        if date_style is not None:
            break
    styles = {0: 0, header_style: original_style_ids["header"], fte_style: original_style_ids["fte"]}
    if date_style is not None:
        styles[date_style] = original_style_ids["date"]

    temp_path = output_path.with_name(output_path.stem + ".partial" + output_path.suffix)
    try:
        results = append_audit_sheets(source_path, temp_path, audit, styles, cache)
        temp_path.replace(output_path)
    finally:
        if temp_path.exists():
            temp_path.unlink()
    if hashlib.sha256(source_path.read_bytes()).hexdigest() != source_hash:
        raise RuntimeError("Input workbook unexpectedly changed")
    for values_only in (True, False):
        reopened = load_workbook(output_path, read_only=True, data_only=values_only, keep_links=True)
        for title, expected in [(AUDIT_NAMES[0], len(gen) + 1), (AUDIT_NAMES[1], len(sped) + 1)]:
            if reopened[title].max_row != expected:
                raise RuntimeError(f"Output row count mismatch: {title}")
        reopened.close()
    metrics["source_sha256"] = source_hash
    metrics["preserved_original_parts"] = sum(r["status"] == "BYTE-IDENTICAL" for r in results)
    metrics["required_metadata_parts_updated"] = 3
    report_path = output_path.with_suffix(".validation.json")
    report_path.write_text(json.dumps({"metrics": metrics, "original_part_validation": results}, indent=2), encoding="utf-8")
    print(json.dumps(metrics, indent=2))
    print("OUTPUT:", str(output_path))
    print("VALIDATION:", str(report_path))
    return output_path


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook")
    parser.add_argument("--output", default="ONTECH_2026_27_DISCREPANCY_AUDIT.xlsx")
    parser.add_argument("--calendar", help="Optional approved CSV with Date,IsSchoolDay (1 or 0) for each calendar date")
    parser.add_argument("--gen-tab", help="Exact name of the general roster tab (default: auto-detect)")
    parser.add_argument("--sped-tab", help="Exact name of the SPED roster tab (default: auto-detect)")
    parser.add_argument("--st-tab", help="Exact name of the SchoolTool enrollment tab (default: auto-detect)")
    parser.add_argument("--disability-tab", help="Exact name of the Disability tab (default: auto-detect)")
    args = parser.parse_args()
    overrides = {"gen": args.gen_tab, "sped": args.sped_tab, "st": args.st_tab, "dis": args.disability_tab}
    try:
        run(args.workbook, args.output, args.calendar, {k: v for k, v in overrides.items() if v})
    except AuditInputError as exc:
        sys.exit(f"ERROR: {exc}")
    except PermissionError as exc:
        sys.exit(f"ERROR: cannot write {exc.filename or args.output}. Close it in Excel and run again.")

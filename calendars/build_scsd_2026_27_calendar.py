"""Build SCSD_2026_27_school_calendar.csv for OnTech_Charter_Audit_2026_27.py.

Source: Syracuse City School District "SCSD Student Calendar 2026-27"
(one-page calendar, revised 09/09/26), linked from
https://www.syracusecityschools.com/calendars
Re-run if the district revises the calendar.
"""
import csv
from datetime import date, timedelta
from pathlib import Path

D = date
CLOSED = {
    **{D(2026, 9, d): "Superintendent Conference Day (no school for students)" for d in (1, 2, 3)},
    D(2026, 9, 4): "No School",
    D(2026, 9, 7): "Labor Day",
    D(2026, 10, 12): "Indigenous Peoples' Day / Columbus Day",
    D(2026, 11, 11): "Veterans Day",
    **{D(2026, 11, d): "Thanksgiving Recess" for d in (25, 26, 27)},
    **{D(2026, 12, d): "Holiday Recess" for d in range(24, 32)},
    D(2027, 1, 1): "Holiday Recess",
    D(2027, 1, 18): "Dr. Martin Luther King, Jr. Day",
    D(2027, 2, 15): "Presidents' Day / Winter Recess",
    **{D(2027, 2, d): "Winter Recess" for d in (16, 17, 18, 19)},
    D(2027, 3, 9): "Eid al-Fitr",
    D(2027, 3, 26): "Good Friday",
    **{D(2027, 4, d): "Spring Recess" for d in range(12, 17)},
    D(2027, 5, 17): "Eid al-Adha / Superintendent Conference Day (no school for students)",
    D(2027, 5, 31): "Memorial Day",
    D(2027, 6, 18): "Juneteenth Observance",
}
NOTES = {
    D(2026, 9, 8): "First Day for Students",
    D(2026, 11, 3): "Election Day - half day for students (counts as a school day)",
    D(2026, 12, 4): "Half day for students (counts as a school day)",
    **{D(2027, 1, d): "Regents Examinations (school day)" for d in range(26, 30)},
    D(2027, 5, 28): "Inclement weather/contingency day - becomes a recess day if no snow days are used",
    **{D(2027, 6, d): "Regents Exams (school day)" for d in (15, 16, 17, 21, 22, 23, 24)},
    D(2027, 6, 25): "Last Day for Students (half day; counts as a school day)",
}
FIRST, LAST = D(2026, 9, 8), D(2027, 6, 25)
# District-published school days per month, used as a check.
PUBLISHED = {(2026, 9): 17, (2026, 10): 21, (2026, 11): 17, (2026, 12): 17, (2027, 1): 19,
             (2027, 2): 15, (2027, 3): 21, (2027, 4): 17, (2027, 5): 19, (2027, 6): 18}

rows, counts = [], {}
d = D(2026, 7, 1)
while d <= D(2027, 6, 30):
    if d.weekday() >= 5:
        flag, note = 0, "Weekend"
    elif d < FIRST:
        flag, note = 0, CLOSED.get(d, "Before first day for students")
    elif d > LAST:
        flag, note = 0, "After last day for students"
    elif d in CLOSED:
        flag, note = 0, CLOSED[d]
    else:
        flag, note = 1, NOTES.get(d, "")
    counts[(d.year, d.month)] = counts.get((d.year, d.month), 0) + flag
    rows.append((d.isoformat(), flag, note))
    d += timedelta(days=1)

bad = {m: (counts.get(m, 0), n) for m, n in PUBLISHED.items() if counts.get(m, 0) != n}
if bad:
    raise SystemExit(f"Month totals do not match the district calendar: {bad}")
out = Path(__file__).with_name("SCSD_2026_27_school_calendar.csv")
with open(out, "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["Date", "IsSchoolDay", "Note"])
    w.writerows(rows)
for (y, m), n in sorted(PUBLISHED.items()):
    print(f"{y}-{m:02d}: {counts[(y, m)]:2d} school days (district says {n})")
print("TOTAL:", sum(r[1] for r in rows), "school days ->", out)

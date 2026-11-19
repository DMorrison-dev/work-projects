# OnTech Charter Invoice Audit (2026-27)

Offline audit of an OnTech charter invoice workbook against the SchoolTool
enrollment and Disability exports embedded in it. Script:
`python/OnTech_Charter_Audit_2026_27.py`.

It never connects to SQL Server or SchoolTool and never edits the input file:
it writes a **new** workbook with audit tabs added, and verifies every original
part of the file is byte-identical.

## Run it

```bash
pip install openpyxl lxml
python python/OnTech_Charter_Audit_2026_27.py "invoice.xlsx" \
    --output "ONTECH_AUDITED.xlsx" --calendar approved_calendar.csv
```

Always run it on the **original** invoice, not on a previous audit output.

## What the workbook needs

**Tabs** (found automatically; case, spacing and `(17)`-style numbers ignored):

| Role | Accepted tab names |
|---|---|
| General roster | `Gen Roster`, `General Roster`, `Roster` |
| SPED roster | `SPED Roster`, `SPED` |
| SchoolTool enrollment | anything starting `ST Enroll` |
| Disability | anything starting `Disability`: `Disability (17)`, `Disability (18)`, `Disability` |

If several Disability or ST tabs exist, the one with 2026-27 rows is used.
To pick a tab yourself: `--gen-tab`, `--sped-tab`, `--st-tab`, `--disability-tab`.

**Columns** are found by header name, in any order, with the header in any of
the first 10 rows:

- Rosters: `Last Name`, `First Name`, `NYSSIS ID`, `DOB`, `Enroll Date`, `Withdraw Date`, `OnTECH FTE`
- ST: `StudentEnrollment_SchoolYear`, `Student_UniversalStudentID`, `StudentID`, `LastName`,
  `FirstName`, `Person_DOB`, `SchoolName`, `StudentEnrollment_StartDate`, `StudentEnrollment_EndDate`
  (optional: `StudentEnrollment_EntryNote`, `StudentEnrollment_ExitNote`)
- Disability: `SchoolYear_Disability`, `Student_UniversalStudentID`, `StudentID_Disability`

School year may be written `2026-2027` or `2026-27`.

## Calendar CSV (strongly recommended)

`Date,IsSchoolDay` — one row per date from 2026-09-03 to 2027-06-30, value
`1`/`0` (also accepts yes/no, true/false). Without it, most date differences
longer than 3 calendar days stay **REVIEW**.

## Rules (summary)

- Entry: charter entry earlier than ST by **more than 3 school days** → NOT GOOD.
- Exit: charter exit later than ST by **more than 3 school days** → NOT GOOD.
- Missing dates, one-sided exits, identity questions, other-school records,
  multiple segments and unreadable dates → REVIEW.
- Overall: NOT GOOD beats REVIEW beats GOOD.
- Disability export match is **not** proof of an active IEP.
- Tentative FTE (weekly, 40-week denominator) is informational, not payment approval.

## Handled automatically

Text dates, `YYYYMMDD`/`03-Sep-2026`-style dates, unreadable dates (flagged,
not crashed), NYSSIS IDs that lost a leading zero, FTE stored as text, wrong
sheet dimensions from exporters, and a locked output file (clear message).
Problems with the input produce a one-line `ERROR:` explaining the fix.

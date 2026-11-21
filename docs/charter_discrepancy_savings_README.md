# Charter Invoice Discrepancy & Savings Calculation (BigQuery)

SQL: [`sql/BigQuery_charter_discrepancy_savings.sql`](../sql/BigQuery_charter_discrepancy_savings.sql)
Project: `charter-invoice-26-27`, dataset `analytics_mart`. All data is de-identified.

Answers two questions for each charter invoice roster:

1. **How many discrepancies were found?**
2. **How much money was saved because of them?**

## Rules

Each roster line is compared to its SchoolTool (ST) enrollment record. The three rule
functions are ports of the Excel master-copy formulas.

| Output | Function | Result |
|---|---|---|
| Entry Dates Validation | `fn_entry_date_validation` | Charter entry more than 3 days **earlier** than ST = UNFAVORABLE |
| Exit Date Validation | `fn_exit_date_validation` | Charter exit more than 3 days **later** than ST, or ST exit with no charter exit = UNFAVORABLE |
| Invoice Validation Status | `fn_invoice_validation_status` | `No Good` (any UNFAVORABLE), `Good / Need ST Update`, `Good`, `Review` |

## FTE

`fn_nysed_fte` = student weeks / 40, capped at 1.000 (NYSED weekly method, ported from the
Python and R audit engines).

- A Monday–Sunday week counts when at least 3 enrolled days fall in that week and month.
- The first week of the start month counts only if its Sunday is in that month.
- Full middle months count as 4 weeks each (July = 5).
- **Charter FTE** uses the charter's dates; **SCSD FTE** uses the SchoolTool dates.

## Savings

| Measure | Meaning |
|---|---|
| Discrepancies found | Roster lines that are not `Good` |
| Billing discrepancies | `No Good` lines only |
| Identified savings | (Charter FTE − SCSD FTE) × rate, on `No Good` lines with an ST record |
| Under review | Full billed amount for students with no ST record (not counted as saved) |
| Confirmed savings | Identified savings on lines marked `Confirmed` in `charter_discrepancy_resolution` |

Rates, school-year dates and the 3-day tolerance live in `charter_savings_config`.
The SPED rate is applied in addition to the Gen Ed rate for SPED students.

## Objects

```
fn_entry_date_validation / fn_exit_date_validation / fn_invoice_validation_status   rule functions
fn_nysed_weeks_in_month / fn_nysed_student_weeks / fn_nysed_fte                     weekly FTE
charter_savings_config                                                              rates, dates, tolerance
charter_discrepancy_resolution                                                      Confirmed / Rejected / Open log
vw_charter_discrepancy_savings                                                      one row per roster line
vw_charter_savings_summary                                                          totals by charter and status
```

## Validation

- Rule functions reproduce the Excel master-copy results on every roster row tested.
- `fn_nysed_student_weeks` matches the Python `weekly_student_weeks()` function on 4,000 generated date pairs.
- The script ends with `ASSERT` checks: no duplicate roster lines, no negative overbilled FTE,
  no FTE above 1.000, and no savings on a line that is not `No Good`.

-- ============================================================================
-- PIPELINE: Charter Invoice Discrepancy & Savings Calculation
-- PROJECT:  charter-invoice-26-27
-- LAYER:    analytics_mart (reads charter_practice roster + SchoolTool enrollment)
--
-- WHAT IT ANSWERS
--   1. How many discrepancies were found?   (rows that are not "Good")
--   2. How much money was saved?            (identified vs confirmed)
--
-- HOW
--   * Entry / exit / status rules are ported from the Excel master copy
--     (Entry Dates Validation, Exit Date Validation, Invoice Validation Status).
--   * FTE uses the NYSED weekly formula (student weeks / 40), ported from
--     weekly_student_weeks() (Python) and nysed_student_weeks() (R engine V5.5).
--   * Savings = (charter FTE - SCSD FTE) x rate, counted only on "No Good" rows.
--   * Students with no SchoolTool record are reported as "under review",
--     not as savings.
--   * Confirmed savings come from the resolution log, once a discrepancy is
--     marked Confirmed.
--
-- All data is de-identified. Idempotent: safe to re-run.
-- ============================================================================

-- ----------------------------------------------------------------------------
-- STEP 1: Rule functions (Excel master copy, Roster columns T / U / V)
-- ----------------------------------------------------------------------------
CREATE OR REPLACE FUNCTION `charter-invoice-26-27.analytics_mart.fn_entry_date_validation`(charter_entry DATE, st_entry DATE, school_start DATE, tolerance_days INT64)
RETURNS STRING AS (
  CASE
    WHEN charter_entry IS NULL AND st_entry IS NULL THEN NULL
    WHEN st_entry IS NULL THEN 'Charter has entry date but ST does not (UNFAVORABLE)'
    WHEN charter_entry IS NULL THEN NULL
    WHEN charter_entry < school_start THEN 'Charter enrolled before school starts (FAVORABLE)'
    WHEN st_entry < school_start AND charter_entry >= school_start THEN 'ST early entry / before school starts (FAVORABLE)'
    WHEN charter_entry = st_entry THEN 'Entry dates match (FAVORABLE)'
    WHEN charter_entry > st_entry THEN CONCAT('Charter entry is ', DATE_DIFF(charter_entry, st_entry, DAY), ' days later than ST entry date (FAVORABLE - update ST discrepancy)')
    WHEN DATE_DIFF(st_entry, charter_entry, DAY) <= tolerance_days THEN CONCAT('Charter entry is ', DATE_DIFF(st_entry, charter_entry, DAY), ' days earlier than ST entry date (FAVORABLE)')
    ELSE CONCAT('Charter entry is ', DATE_DIFF(st_entry, charter_entry, DAY), ' days earlier than ST entry date (UNFAVORABLE)')
  END
);

CREATE OR REPLACE FUNCTION `charter-invoice-26-27.analytics_mart.fn_exit_date_validation`(charter_exit DATE, st_exit DATE, tolerance_days INT64)
RETURNS STRING AS (
  CASE
    WHEN charter_exit IS NULL AND st_exit IS NULL THEN 'No Charter or ST exit dates on file (FAVORABLE)'
    WHEN st_exit IS NULL THEN 'Charter has exit date but ST does not (Discrepancy - update in ST)'
    WHEN charter_exit IS NULL THEN 'ST has exit date but Charter does not (UNFAVORABLE)'
    WHEN charter_exit = st_exit THEN 'Exit dates match (FAVORABLE)'
    WHEN charter_exit < st_exit THEN CONCAT('Charter exit is ', DATE_DIFF(st_exit, charter_exit, DAY), ' days earlier than ST exit date (Discrepancy - update in ST)')
    WHEN DATE_DIFF(charter_exit, st_exit, DAY) <= tolerance_days THEN CONCAT('Charter exit is ', DATE_DIFF(charter_exit, st_exit, DAY), ' days later than ST exit date (FAVORABLE)')
    ELSE CONCAT('Charter exit is ', DATE_DIFF(charter_exit, st_exit, DAY), ' days later than ST exit date (UNFAVORABLE)')
  END
);

CREATE OR REPLACE FUNCTION `charter-invoice-26-27.analytics_mart.fn_invoice_validation_status`(entry_validation STRING, exit_validation STRING)
RETURNS STRING AS (
  CASE
    WHEN IFNULL(entry_validation,'') LIKE '%UNFAVORABLE%' OR IFNULL(exit_validation,'') LIKE '%UNFAVORABLE%' THEN 'No Good'
    WHEN IFNULL(entry_validation,'') LIKE '%update in ST%' OR IFNULL(exit_validation,'') LIKE '%update in ST%'
      OR IFNULL(entry_validation,'') LIKE '%update ST%'    OR IFNULL(exit_validation,'') LIKE '%update ST%' THEN 'Good / Need ST Update'
    WHEN IFNULL(entry_validation,'') LIKE '%FAVORABLE%' AND IFNULL(exit_validation,'') LIKE '%FAVORABLE%' THEN 'Good'
    ELSE 'Review'
  END
);

-- ----------------------------------------------------------------------------
-- STEP 2: NYSED weekly FTE (student weeks / 40 program weeks, capped at 1.000)
--   * A Mon-Sun week counts when >= 3 enrolled days fall in that week and month
--   * First week of the start month counts only if its Sunday is in that month
--   * Full middle months count as 4 weeks each (July = 5)
--   * Fewer than 3 calendar days = 0 weeks; exactly 3 days = 1 week
-- ----------------------------------------------------------------------------
CREATE OR REPLACE FUNCTION `charter-invoice-26-27.analytics_mart.fn_nysed_weeks_in_month`(a DATE, b DATE, apply_first_rule BOOL)
RETURNS INT64 AS ((
  SELECT LEAST(
           COUNTIF(days_in >= 3 AND (NOT is_first OR NOT apply_first_rule OR sunday <= month_end)),
           IF(EXTRACT(MONTH FROM a) = 7, 5, 4))
  FROM (
    SELECT
      DATE_ADD(monday, INTERVAL 6 DAY) AS sunday,
      LAST_DAY(a) AS month_end,
      monday = DATE_TRUNC(a, WEEK(MONDAY)) AS is_first,
      GREATEST(0, DATE_DIFF(LEAST(b, DATE_ADD(monday, INTERVAL 6 DAY), LAST_DAY(a)), GREATEST(a, monday), DAY) + 1) AS days_in
    FROM UNNEST(GENERATE_DATE_ARRAY(DATE_TRUNC(a, WEEK(MONDAY)), b, INTERVAL 7 DAY)) AS monday
  )
));

CREATE OR REPLACE FUNCTION `charter-invoice-26-27.analytics_mart.fn_nysed_student_weeks`(start_date DATE, end_date DATE)
RETURNS INT64 AS (
  CASE
    WHEN start_date IS NULL OR end_date IS NULL THEN NULL
    WHEN end_date < start_date THEN 0
    WHEN DATE_DIFF(end_date, start_date, DAY) + 1 < 3 THEN 0
    WHEN DATE_DIFF(end_date, start_date, DAY) + 1 = 3 THEN 1
    WHEN DATE_TRUNC(start_date, MONTH) = DATE_TRUNC(end_date, MONTH)
      THEN `charter-invoice-26-27.analytics_mart.fn_nysed_weeks_in_month`(start_date, end_date, TRUE)
    ELSE `charter-invoice-26-27.analytics_mart.fn_nysed_weeks_in_month`(start_date, LAST_DAY(start_date), TRUE)
       + GREATEST(0, DATE_DIFF(end_date, start_date, MONTH) - 1) * 4
       + `charter-invoice-26-27.analytics_mart.fn_nysed_weeks_in_month`(DATE_TRUNC(end_date, MONTH), end_date, FALSE)
  END
);

CREATE OR REPLACE FUNCTION `charter-invoice-26-27.analytics_mart.fn_nysed_fte`(start_date DATE, end_date DATE)
RETURNS NUMERIC AS (
  LEAST(ROUND(CAST(`charter-invoice-26-27.analytics_mart.fn_nysed_student_weeks`(start_date, end_date) AS NUMERIC) / 40, 3), NUMERIC '1')
);

-- ----------------------------------------------------------------------------
-- STEP 3: Config (one active row) and resolution log
-- ----------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS `charter-invoice-26-27.analytics_mart.charter_savings_config`
(
  school_year        STRING NOT NULL,
  school_start_date  DATE   NOT NULL OPTIONS(description='First day of school; entries before this date are treated as favorable, as in the Excel entry rule'),
  school_end_date    DATE   NOT NULL OPTIONS(description='Last billable day; used when the charter roster has no withdraw date'),
  tolerance_days     INT64  NOT NULL OPTIONS(description='Date variances up to this many days are favorable; more is UNFAVORABLE'),
  gen_ed_rate_usd    NUMERIC NOT NULL OPTIONS(description='Annual per-FTE rate from the General Education invoice'),
  sped_rate_usd      NUMERIC NOT NULL OPTIONS(description='Annual per-FTE Special Needs Reimbursement rate from the SPED invoice, billed in addition to the Gen Ed rate'),
  is_active          BOOL   NOT NULL OPTIONS(description='Exactly one row should be TRUE; the views use that row')
);

INSERT INTO `charter-invoice-26-27.analytics_mart.charter_savings_config`
SELECT '2025-26', DATE '2025-09-03', DATE '2026-06-30', 3, NUMERIC '16959', NUMERIC '23299', TRUE
FROM (SELECT 1) WHERE NOT EXISTS (SELECT 1 FROM `charter-invoice-26-27.analytics_mart.charter_savings_config`);

CREATE TABLE IF NOT EXISTS `charter-invoice-26-27.analytics_mart.charter_discrepancy_resolution`
(
  student_id            STRING NOT NULL OPTIONS(description='De-identified student key, same as in charter_enrollment_raw'),
  charter_school_name   STRING NOT NULL,
  entry_date            DATE   NOT NULL,
  resolution_status     STRING NOT NULL OPTIONS(description='Confirmed (charter corrected / credit issued), Rejected (billing was valid), or Open'),
  credited_amount_usd   NUMERIC OPTIONS(description='Actual dollar amount credited or removed from the invoice; leave NULL to use the calculated amount'),
  resolved_date         DATE,
  notes                 STRING
)
OPTIONS(description='Log of how each flagged charter billing discrepancy was resolved. One row per student + charter school + entry date. Drives confirmed (realized) savings.');

-- ----------------------------------------------------------------------------
-- STEP 4: Row-level view - rules, FTE, savings
-- ----------------------------------------------------------------------------
CREATE OR REPLACE VIEW `charter-invoice-26-27.analytics_mart.vw_charter_discrepancy_savings`
OPTIONS(description='One row per charter roster line: Excel-rule entry/exit validation against SchoolTool, invoice validation status, NYSED weekly FTE (student weeks / 40) on charter dates and on SchoolTool dates, and identified vs confirmed savings.')
AS
WITH cfg AS (
  SELECT * FROM `charter-invoice-26-27.analytics_mart.charter_savings_config` WHERE is_active
),
school_codes AS (
  SELECT DISTINCT school_name, charter_code FROM `charter-invoice-26-27.charter_practice.charter_roster`
),
st AS (
  SELECT s.*, sc.charter_code
  FROM `charter-invoice-26-27.charter_practice.st_enrollment` AS s
  JOIN school_codes AS sc USING (school_name)
),
matched AS (
  -- one SchoolTool record per roster line: same student and charter; exact school name first, then closest entry date
  SELECT
    r.*,
    s.st_row_id IS NOT NULL AS found_in_st,
    s.entry_date AS st_entry_date,
    s.exit_date  AS st_exit_date,
    s.exit_reason AS st_exit_reason,
    ROW_NUMBER() OVER (
      PARTITION BY r.row_id
      ORDER BY IF(s.school_name = r.school_name, 0, 1),
               ABS(DATE_DIFF(s.entry_date, r.enroll_date, DAY)),
               s.st_row_id) AS st_rank
  FROM `charter-invoice-26-27.charter_practice.charter_roster` AS r
  LEFT JOIN st AS s
    ON s.student_key = r.student_key AND s.charter_code = r.charter_code
),
ruled AS (
  SELECT
    m.*, cfg.gen_ed_rate_usd, cfg.sped_rate_usd,
    `charter-invoice-26-27.analytics_mart.fn_entry_date_validation`(m.enroll_date, m.st_entry_date, cfg.school_start_date, cfg.tolerance_days) AS entry_dates_validation,
    `charter-invoice-26-27.analytics_mart.fn_exit_date_validation`(m.withdraw_date, m.st_exit_date, cfg.tolerance_days) AS exit_date_validation,
    -- Charter FTE: NYSED weekly formula on the charter's own dates (this roster table has no charter-billed FTE column)
    `charter-invoice-26-27.analytics_mart.fn_nysed_fte`(
        GREATEST(m.enroll_date, cfg.school_start_date),
        COALESCE(m.withdraw_date, cfg.school_end_date)) AS charter_fte,
    -- SCSD FTE: same formula on SchoolTool dates. Start = ST entry (charter entry if ST has none), not before the first day.
    -- End = ST exit if before the official last day, else charter exit, else the official last day.
    `charter-invoice-26-27.analytics_mart.fn_nysed_fte`(
        GREATEST(COALESCE(m.st_entry_date, m.enroll_date), cfg.school_start_date),
        IF(m.st_exit_date IS NOT NULL AND m.st_exit_date < cfg.school_end_date, m.st_exit_date,
           COALESCE(m.withdraw_date, cfg.school_end_date))) AS scsd_fte_calc
  FROM matched AS m CROSS JOIN cfg
  WHERE m.st_rank = 1
),
priced AS (
  SELECT
    r.*,
    `charter-invoice-26-27.analytics_mart.fn_invoice_validation_status`(entry_dates_validation, exit_date_validation) AS invoice_validation_status,
    STARTS_WITH(IFNULL(sped_flag, ''), 'SPED') AS is_sped,
    charter_fte - scsd_fte_calc AS fte_variance
  FROM ruled AS r
),
final AS (
  SELECT
    p.*,
    -- Dollars only where the Excel rules say No Good, SchoolTool has a record, and the charter FTE is higher
    IF(invoice_validation_status = 'No Good' AND found_in_st, GREATEST(fte_variance, 0), NUMERIC '0') AS overbilled_fte,
    -- No SchoolTool record at all: nothing to compare, so the whole FTE is under review rather than counted as saved
    IF(NOT found_in_st, charter_fte, NUMERIC '0') AS under_review_fte
  FROM priced AS p
)
SELECT
  f.row_id,
  f.charter_code,
  f.school_name,
  f.charter_month,
  f.student_key,
  f.grade_level,
  f.sped_flag,
  f.enroll_date,
  f.withdraw_date,
  f.found_in_st,
  f.st_entry_date,
  f.st_exit_date,
  f.st_exit_reason,
  f.entry_dates_validation,
  f.exit_date_validation,
  f.invoice_validation_status,
  f.invoice_validation_status IN ('No Good', 'Good / Need ST Update', 'Review') AS is_discrepancy,
  f.invoice_validation_status = 'No Good'      AS is_billing_discrepancy,
  f.charter_fte,
  f.scsd_fte_calc                              AS scsd_fte,
  f.scsd_fte                                   AS source_scsd_fte,
  f.fte_variance,
  f.overbilled_fte,
  f.under_review_fte,
  ROUND(f.charter_fte * f.gen_ed_rate_usd, 2)                                   AS billed_gen_ed_usd,
  ROUND(IF(f.is_sped, f.charter_fte * f.sped_rate_usd, 0), 2)                   AS billed_sped_usd,
  ROUND(f.overbilled_fte * f.gen_ed_rate_usd, 2)                                AS identified_gen_ed_savings_usd,
  ROUND(IF(f.is_sped, f.overbilled_fte * f.sped_rate_usd, 0), 2)                AS identified_sped_savings_usd,
  ROUND(f.overbilled_fte * (f.gen_ed_rate_usd + IF(f.is_sped, f.sped_rate_usd, 0)), 2) AS identified_savings_usd,
  ROUND(f.under_review_fte * (f.gen_ed_rate_usd + IF(f.is_sped, f.sped_rate_usd, 0)), 2) AS under_review_usd,
  CASE WHEN f.invoice_validation_status != 'No Good' THEN 'Not applicable'
       ELSE IFNULL(res.resolution_status, 'Open') END                           AS resolution_status,
  CASE WHEN f.invoice_validation_status = 'No Good' AND res.resolution_status = 'Confirmed'
       THEN IFNULL(res.credited_amount_usd,
                   ROUND(IF(f.found_in_st, f.overbilled_fte, f.under_review_fte) * (f.gen_ed_rate_usd + IF(f.is_sped, f.sped_rate_usd, 0)), 2))
       ELSE NUMERIC '0' END                                                     AS confirmed_savings_usd,
  res.resolved_date,
  res.notes AS resolution_notes
FROM final AS f
LEFT JOIN `charter-invoice-26-27.analytics_mart.charter_discrepancy_resolution` AS res
  ON  res.student_id = f.student_key
  AND res.charter_school_name = f.school_name
  AND res.entry_date = f.enroll_date;

-- ----------------------------------------------------------------------------
-- STEP 5: Summary view - by charter and status, with totals
-- ----------------------------------------------------------------------------
CREATE OR REPLACE VIEW `charter-invoice-26-27.analytics_mart.vw_charter_savings_summary`
OPTIONS(description='Discrepancies found and money saved by charter and invoice validation status, with charter totals and a grand total.')
AS
SELECT
  IF(GROUPING(code) = 1, 'ALL CHARTERS', code)   AS charter_code,
  IF(GROUPING(status) = 1, 'TOTAL', status)      AS invoice_validation_status,
  COUNT(*)                                       AS roster_lines,
  COUNTIF(is_discrepancy)                        AS discrepancies_found,
  COUNTIF(is_billing_discrepancy)                AS billing_discrepancies,
  COUNTIF(NOT found_in_st)                       AS not_found_in_st,
  SUM(charter_fte)                               AS charter_fte,
  SUM(scsd_fte)                                  AS scsd_fte,
  SUM(overbilled_fte)                            AS overbilled_fte,
  SUM(billed_gen_ed_usd + billed_sped_usd)       AS billed_usd,
  SUM(identified_gen_ed_savings_usd)             AS identified_gen_ed_savings_usd,
  SUM(identified_sped_savings_usd)               AS identified_sped_savings_usd,
  SUM(identified_savings_usd)                    AS identified_savings_usd,
  SUM(under_review_usd)                          AS under_review_usd,
  SUM(confirmed_savings_usd)                     AS confirmed_savings_usd,
  SUM(IF(resolution_status = 'Open', identified_savings_usd, 0)) AS open_savings_usd,
  ROUND(100 * SAFE_DIVIDE(SUM(identified_savings_usd), SUM(billed_gen_ed_usd + billed_sped_usd)), 2) AS identified_pct_of_billed
FROM (SELECT *, charter_code AS code, invoice_validation_status AS status
      FROM `charter-invoice-26-27.analytics_mart.vw_charter_discrepancy_savings`)
GROUP BY ROLLUP(code, status);

-- ----------------------------------------------------------------------------
-- STEP 6: Data-quality assertions (fail the run on bad output)
-- ----------------------------------------------------------------------------
ASSERT (
  SELECT COUNT(*) = COUNT(DISTINCT row_id) AND COUNT(*) > 0
  FROM `charter-invoice-26-27.analytics_mart.vw_charter_discrepancy_savings`
) AS 'DQ ERROR: savings view is empty or has duplicate roster lines.';

ASSERT (
  SELECT COUNTIF(overbilled_fte < 0 OR charter_fte > 1 OR scsd_fte > 1
                 OR (invoice_validation_status != 'No Good' AND identified_savings_usd != 0)) = 0
  FROM `charter-invoice-26-27.analytics_mart.vw_charter_discrepancy_savings`
) AS 'DQ ERROR: negative overbilled FTE, FTE above 1.000, or savings on a row that is not No Good.';

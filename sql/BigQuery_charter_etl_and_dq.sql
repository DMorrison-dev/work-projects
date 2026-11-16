-- ============================================================================
-- PIPELINE: Charter Enrollment ETL & Discrepancy Audit Engine
-- PROJECT:  charter-invoice-26-27
-- LAYER:    Staging (raw_staging) -> Clean & Mart (analytics_mart)
-- SCHEDULE: Run daily via BigQuery Scheduled Queries
-- ============================================================================

-- STEP 1: Build Partitioned & Clustered Clean Table (Idempotent)
CREATE OR REPLACE TABLE `charter-invoice-26-27.analytics_mart.charter_enrollment_audit`
PARTITION BY DATE_TRUNC(entry_date, MONTH)
CLUSTER BY charter_school_name, audit_status AS

WITH deduplicated_records AS (
  SELECT
    student_id,
    charter_school_name,
    CAST(entry_date AS DATE) AS entry_date,
    CAST(exit_date AS DATE) AS exit_date,
    enrollment_type, -- 'GEN_ED' or 'SPED'
    CAST(fte_claimed AS NUMERIC) AS fte_claimed,
    in_district_residency,
    has_active_iep,
    -- Flag duplicate student enrollments across schools on the same date
    ROW_NUMBER() OVER (
      PARTITION BY student_id, entry_date
      ORDER BY updated_at DESC
    ) AS row_num,
    COUNT(*) OVER (
      PARTITION BY student_id
    ) AS total_student_records
  FROM `charter-invoice-26-27.raw_staging.charter_enrollment_raw`
  WHERE student_id IS NOT NULL
)

SELECT
  student_id,
  charter_school_name,
  entry_date,
  exit_date,
  enrollment_type,
  fte_claimed,
  -- Apply District Tuition Rates ($15,471 Gen Ed | $22,687 SPED)
  CASE
    WHEN enrollment_type = 'SPED' THEN ROUND(fte_claimed * 22687.00, 2)
    ELSE ROUND(fte_claimed * 15471.00, 2)
  END AS billed_amount_usd,
  -- Automated Discrepancy Rules Engine
  CASE
    WHEN total_student_records > 1 THEN 'FLAG: Duplicate Enrollment ID'
    WHEN in_district_residency = FALSE THEN 'FLAG: Out-of-District Residency'
    WHEN enrollment_type = 'SPED' AND has_active_iep = FALSE THEN 'FLAG: Missing Valid IEP'
    WHEN DATE_DIFF(COALESCE(exit_date, CURRENT_DATE()), entry_date, DAY) <= 3 THEN 'FLAG: <=3 Day Short Enrollment'
    ELSE 'VERIFIED'
  END AS audit_status,
  CURRENT_TIMESTAMP() AS pipeline_run_utc
FROM deduplicated_records
WHERE row_num = 1;

-- STEP 2: Automated Data Quality Assertions (Fails Pipeline on Bad Data)
ASSERT (
  SELECT COUNT(*) FROM `charter-invoice-26-27.analytics_mart.charter_enrollment_audit`
) > 0 AS 'DQ ERROR: Output table is empty after transformation.';

ASSERT (
  SELECT COUNT(*) FROM `charter-invoice-26-27.analytics_mart.charter_enrollment_audit`
  WHERE student_id IS NULL OR billed_amount_usd < 0
) = 0 AS 'DQ ERROR: Null student IDs or negative billing amounts detected.';

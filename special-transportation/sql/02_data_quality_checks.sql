-- Special Transportation Listing: data-quality checks (BigQuery)
-- Each check returns one row per snapshot; "expect" columns say what a clean load looks like.

-- 1. Row counts, distinct students, exact-duplicate rows, required-field nulls
SELECT
  snapshot_date,
  COUNT(*)                                                   AS row_ct,
  COUNT(DISTINCT student_id)                                 AS distinct_student_ids,
  COUNT(*) - COUNT(DISTINCT TO_JSON_STRING(t))               AS duplicate_rows,          -- expect 0
  COUNTIF(student_id IS NULL)                                AS null_student_id,         -- expect 0 / 0 / 2 / 3
  COUNTIF(school IS NULL)                                    AS null_school,             -- expect 0 / 0 / 2 / 3
  COUNTIF(need IS NULL OR transportation_recommendation IS NULL
          OR cr_decision_status IS NULL)                     AS null_required_fields     -- expect 0
FROM `school-district-work.special_transportation.special_transportation_listing_synthetic` AS t
GROUP BY snapshot_date
ORDER BY snapshot_date;

-- 2. DI consistency: a date exists if and only if DI Finalized = 'Yes'; no date after its snapshot
SELECT
  snapshot_date,
  COUNTIF(di_finalized = 'Yes' AND di_date_finalized IS NULL)      AS yes_without_date,   -- expect 0
  COUNTIF(di_finalized = 'No'  AND di_date_finalized IS NOT NULL)  AS no_with_date,       -- expect 0
  COUNTIF(di_finalized IS NULL AND snapshot_date >= '2026-08-01')  AS missing_di_flag,    -- expect 0
  COUNTIF(di_date_finalized > snapshot_date)                       AS date_after_snapshot -- expect 0
FROM `school-district-work.special_transportation.special_transportation_listing_synthetic`
GROUP BY snapshot_date
ORDER BY snapshot_date;

-- 3. De-identification guards: every ID is synthetic (9xxxxxxxx) and each ID maps to one name and one school per snapshot
SELECT
  COUNTIF(student_id NOT BETWEEN 900000000 AND 999999999)  AS non_synthetic_ids,          -- expect 0
  (SELECT COUNT(*) FROM (
     SELECT student_id
     FROM `school-district-work.special_transportation.special_transportation_listing_synthetic`
     WHERE student_id IS NOT NULL
     GROUP BY student_id
     HAVING COUNT(DISTINCT CONCAT(last_name, '|', first_name)) > 1))  AS ids_with_multiple_names,  -- expect 0
  (SELECT COUNT(*) FROM (
     SELECT snapshot_date, student_id
     FROM `school-district-work.special_transportation.special_transportation_listing_synthetic`
     WHERE student_id IS NOT NULL
     GROUP BY snapshot_date, student_id
     HAVING COUNT(DISTINCT school) > 1))                              AS ids_with_multiple_schools  -- expect 0
FROM `school-district-work.special_transportation.special_transportation_listing_synthetic`;

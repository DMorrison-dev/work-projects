-- Special Transportation Listing: analysis queries (BigQuery)

-- 1. Need mix by snapshot
SELECT
  snapshot_date,
  need,
  COUNT(*)                                                              AS row_ct,
  COUNT(DISTINCT student_id)                                            AS students,
  ROUND(100 * COUNT(*) / SUM(COUNT(*)) OVER (PARTITION BY snapshot_date), 1) AS pct_of_snapshot
FROM `school-district-work.special_transportation.special_transportation_listing_synthetic`
GROUP BY snapshot_date, need
ORDER BY snapshot_date, row_ct DESC;

-- 2. Latest snapshot: students and recommendations by school level
SELECT
  IFNULL(school_level, '(blank)')       AS school_level,
  COUNT(DISTINCT school)                AS schools,
  COUNT(DISTINCT student_id)            AS students,
  COUNT(*)                              AS recommendations,
  ROUND(COUNT(*) / NULLIF(COUNT(DISTINCT student_id), 0), 2) AS recs_per_student
FROM `school-district-work.special_transportation.special_transportation_listing_synthetic`
WHERE snapshot_date = (SELECT MAX(snapshot_date)
                       FROM `school-district-work.special_transportation.special_transportation_listing_synthetic`)
GROUP BY school_level
ORDER BY students DESC;

-- 3. DI finalization rate by school, latest snapshot (lowest first)
SELECT
  school,
  school_level,
  COUNT(*)                                           AS recommendations,
  COUNTIF(di_finalized = 'Yes')                      AS finalized,
  ROUND(100 * COUNTIF(di_finalized = 'Yes') / COUNT(*), 1) AS pct_finalized
FROM `school-district-work.special_transportation.special_transportation_listing_synthetic`
WHERE snapshot_date = (SELECT MAX(snapshot_date)
                       FROM `school-district-work.special_transportation.special_transportation_listing_synthetic`)
  AND school IS NOT NULL
GROUP BY school, school_level
ORDER BY pct_finalized, recommendations DESC;

-- 4. Snapshot-over-snapshot change: students added and dropped versus the previous snapshot
WITH s AS (
  SELECT DISTINCT snapshot_date, student_id
  FROM `school-district-work.special_transportation.special_transportation_listing_synthetic`
  WHERE student_id IS NOT NULL
),
pairs AS (
  SELECT snapshot_date, prev_date
  FROM (SELECT snapshot_date, LAG(snapshot_date) OVER (ORDER BY snapshot_date) AS prev_date
        FROM (SELECT DISTINCT snapshot_date FROM s))
  WHERE prev_date IS NOT NULL
),
cur AS (SELECT p.snapshot_date, s.student_id FROM pairs AS p JOIN s USING (snapshot_date)),
prv AS (SELECT p.snapshot_date, s.student_id FROM pairs AS p JOIN s ON s.snapshot_date = p.prev_date)
SELECT
  COALESCE(cur.snapshot_date, prv.snapshot_date)  AS snapshot_date,
  COUNT(cur.student_id)                           AS students,
  COUNTIF(prv.student_id IS NULL)                 AS added,
  COUNTIF(cur.student_id IS NULL)                 AS dropped
FROM cur
FULL JOIN prv
  ON cur.snapshot_date = prv.snapshot_date AND cur.student_id = prv.student_id
GROUP BY 1
ORDER BY 1;

-- 5. Students whose transportation recommendation set changed between the two latest snapshots
WITH r AS (
  SELECT snapshot_date, student_id,
         STRING_AGG(DISTINCT transportation_recommendation, ' | ' ORDER BY transportation_recommendation) AS recs
  FROM `school-district-work.special_transportation.special_transportation_listing_synthetic`
  WHERE student_id IS NOT NULL
  GROUP BY snapshot_date, student_id
)
SELECT cur.student_id, prev.recs AS previous_recommendations, cur.recs AS current_recommendations
FROM r AS cur
JOIN r AS prev USING (student_id)
WHERE cur.snapshot_date  = DATE '2026-09-08'
  AND prev.snapshot_date = DATE '2026-08-25'
  AND cur.recs != prev.recs
ORDER BY cur.student_id;

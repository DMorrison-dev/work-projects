-- Port of NYSSISID_Update.py. Load the export into schooltool_replica.stg_nyssis (StudentID STRING, NYSSISID STRING).
CREATE OR REPLACE TABLE schooltool_replica.stg_nyssis AS
SELECT * FROM UNNEST([STRUCT('0002000001' AS StudentID,'900000099999' AS NYSSISID),
                      STRUCT('2000002', NULL),
                      STRUCT('2999999','900000011111')]);

-- Preview: what would change, and unmatched IDs
SELECT g.StudentID, s.UniversalStudentID AS old_id, g.NYSSISID AS new_id,
       IF(s.ID IS NULL,'NO MATCH','UPDATE') AS action
FROM schooltool_replica.stg_nyssis g
LEFT JOIN schooltool_replica.Student s ON s.StudentID = LTRIM(CAST(g.StudentID AS STRING), '0')
WHERE g.NYSSISID IS NOT NULL AND (s.ID IS NULL OR s.UniversalStudentID <> g.NYSSISID);

UPDATE schooltool_replica.Student s
SET UniversalStudentID = g.NYSSISID
FROM schooltool_replica.stg_nyssis g
WHERE s.StudentID = LTRIM(CAST(g.StudentID AS STRING), '0')
  AND g.NYSSISID IS NOT NULL AND s.UniversalStudentID <> g.NYSSISID;

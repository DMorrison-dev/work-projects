-- Port of FreeReducedLunch.py (create = new programs). Load rows into schooltool_replica.stg_frl
-- (StudentID STRING, Code STRING, StartDate DATE, EligibilityCode STRING). BigQuery has no row-level
-- transaction here, so run as one script: BEGIN TRANSACTION makes it all-or-nothing (fixes the half-applied crash risk).
CREATE OR REPLACE TABLE schooltool_replica.stg_frl AS
SELECT * FROM UNNEST([STRUCT('2000010' AS StudentID,'P016' AS Code,DATE '2026-11-01' AS StartDate,'SNAP' AS EligibilityCode),
                      STRUCT('2000011','P017',DATE '2026-11-01',NULL)]);

BEGIN TRANSACTION;

CREATE TEMP TABLE frl AS
SELECT s.ID AS Student_ID, p.ID AS OtherProgram_ID, g.StartDate, e.Value AS EligibilityTypeValue
FROM schooltool_replica.stg_frl g
JOIN schooltool_replica.Student s ON s.StudentID = g.StudentID
JOIN schooltool_replica.OtherProgram p ON p.STEPProgramServiceCode = g.Code
LEFT JOIN schooltool_replica.OtherProgramEligibilityType e ON e.Code = g.EligibilityCode;

-- 1) end old FRL (16,17) and poverty (2) records the day before the new start
UPDATE schooltool_replica.StudentOtherProgram o
SET EndDate = DATE_SUB(f.StartDate, INTERVAL 1 DAY), ModifiedOn = CURRENT_TIMESTAMP(), ModifiedByPerson_ID = 141247
FROM frl f
WHERE o.Student_ID = f.Student_ID AND o.OtherProgram_ID IN (2,16,17) AND o.EndDate IS NULL;

-- 2) insert the new FRL record
INSERT INTO schooltool_replica.StudentOtherProgram
 (ID, Student_ID, OtherProgram_ID, StartDate, BEDSProvider_ID, CreatedOn, CreatedByPerson_ID, ModifiedOn, ModifiedByPerson_ID, EligibilityTypeValue)
SELECT (SELECT IFNULL(MAX(ID),0) FROM schooltool_replica.StudentOtherProgram) + ROW_NUMBER() OVER (ORDER BY Student_ID),
       Student_ID, OtherProgram_ID, StartDate, 5038, CURRENT_TIMESTAMP(), 141247, CURRENT_TIMESTAMP(), 141247, EligibilityTypeValue
FROM frl;

-- 3) everyone gets a poverty record (program 2)
INSERT INTO schooltool_replica.StudentOtherProgram
 (ID, Student_ID, OtherProgram_ID, StartDate, BEDSProvider_ID, CreatedOn, CreatedByPerson_ID, ModifiedOn, ModifiedByPerson_ID)
SELECT (SELECT MAX(ID) FROM schooltool_replica.StudentOtherProgram) + ROW_NUMBER() OVER (ORDER BY Student_ID),
       Student_ID, 2, StartDate, 5038, CURRENT_TIMESTAMP(), 141247, CURRENT_TIMESTAMP(), 141247
FROM frl f
WHERE NOT EXISTS (SELECT 1 FROM schooltool_replica.StudentOtherProgram o
                  WHERE o.Student_ID = f.Student_ID AND o.OtherProgram_ID = 2 AND o.EndDate IS NULL);

COMMIT TRANSACTION;

-- Port of DisabilityEndDates.py. Load the spreadsheet into schooltool_replica.stg_disability_end
-- (StudentID STRING, Code STRING, StartDate DATE, EndDate DATE) first.
-- Exit reason: 2 (912) if the student has a later disability record, else 1 (901).
CREATE OR REPLACE TABLE schooltool_replica.stg_disability_end AS
SELECT * FROM UNNEST([STRUCT('2000003' AS StudentID,'S002' AS Code,DATE '2025-09-02' AS StartDate,DATE '2026-06-30' AS EndDate),
                      STRUCT('2000003','S004',DATE '2023-09-05',DATE '2025-09-01')]);

-- Rows that match nothing (report these instead of crashing):
SELECT g.* FROM schooltool_replica.stg_disability_end g
LEFT JOIN schooltool_replica.Student s ON s.StudentID = g.StudentID
LEFT JOIN schooltool_replica.Disability d ON d.STEPProgramServiceCode = g.Code
LEFT JOIN schooltool_replica.StudentDisability sd
  ON sd.Student_ID = s.ID AND sd.Disability_ID = d.ID AND sd.StartDate = g.StartDate
WHERE sd.ID IS NULL;

UPDATE schooltool_replica.StudentDisability sd
SET EndDate = m.EndDate,
    DisabilityExitReason_ID = IF(EXISTS (SELECT 1 FROM schooltool_replica.StudentDisability n
                                         WHERE n.Student_ID = sd.Student_ID AND n.StartDate > m.EndDate), 2, 1),
    ModifiedByPerson_ID = 141247, ModifiedOn = CURRENT_TIMESTAMP()
FROM (
  SELECT s.ID AS Student_ID, d.ID AS Disability_ID, g.StartDate, g.EndDate
  FROM schooltool_replica.stg_disability_end g
  JOIN schooltool_replica.Student s ON s.StudentID = g.StudentID
  JOIN schooltool_replica.Disability d ON d.STEPProgramServiceCode = g.Code
) m
WHERE sd.Student_ID = m.Student_ID AND sd.Disability_ID = m.Disability_ID AND sd.StartDate = m.StartDate;

-- Port of OutOfDistrictProjections.py. Set target_year_id / start_date for the rollover year.
DECLARE target_year_id INT64 DEFAULT 40;
DECLARE start_date DATE DEFAULT DATE '2027-07-01';

INSERT INTO schooltool_replica.StudentProjectedEnrollment
 (ID, Student_ID, Grade_ID, SchoolYear_ID, CreatedOn, CreatedByPerson_ID, ModifiedOn, ModifiedByPerson_ID, ProjectedStartDate, IsPlaced,
  Enrollment_ID, EnrollmentReason_ID, BEDSProvider_ID, ExpectedGraduationSchoolYear_ID, IncludeInRank, EntryNote, AutoStart)
SELECT (SELECT IFNULL(MAX(ID),0) FROM schooltool_replica.StudentProjectedEnrollment) + ROW_NUMBER() OVER (ORDER BY e.Student_ID),
  e.Student_ID,
  CASE WHEN e.Grade_ID IN (2,3) THEN 5 WHEN e.Grade_ID = 16 THEN 21 WHEN e.Grade_ID BETWEEN 40 AND 45 THEN 4
       WHEN e.Grade_ID IN (20,21) THEN e.Grade_ID ELSE e.Grade_ID + 1 END,
  target_year_id, CURRENT_TIMESTAMP(), 141247, CURRENT_TIMESTAMP(), 141247, start_date, 1,
  e.Enrollment_ID, 28, e.BEDSProvider_ID, e.ExpectedGraduationSchoolYear_ID, 1, 'Rollover Process', 0
FROM schooltool_replica.StudentBuildingSchoolLevelEnrollment e
WHERE e.EndDate IS NULL AND e.Building_SchoolLevel_ID IS NULL
  AND NOT EXISTS (SELECT 1 FROM schooltool_replica.StudentProjectedEnrollment p
                  WHERE p.Student_ID = e.Student_ID AND p.SchoolYear_ID = target_year_id);  -- year-aware, unlike the original

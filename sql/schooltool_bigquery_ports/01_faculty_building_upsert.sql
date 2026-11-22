-- Port of SummerSchoolBuildings.py: give a list of faculty access to a building/school level (default 80 = ESY).
-- IDs: 75 ES, 76 UPK, 77 MS, 78 HS, 80 ESY. Replace the faculty list and bsl_id below.
DECLARE bsl_id INT64 DEFAULT 80;
DECLARE faculty_ids ARRAY<STRING> DEFAULT ['30001','30002','30003','99999'];

MERGE schooltool_replica.FacultyBuildingSchoolLevel t
USING (
  SELECT f.ID AS Faculty_ID
  FROM UNNEST(faculty_ids) fid
  JOIN schooltool_replica.Faculty f ON f.FacultyID = fid   -- unknown IDs drop out instead of inserting NULL
) s
ON t.Faculty_ID = s.Faculty_ID AND t.Building_SchoolLevel_ID = bsl_id
WHEN MATCHED THEN UPDATE SET EndDate = NULL, LastModifiedOn = CURRENT_TIMESTAMP()
WHEN NOT MATCHED THEN INSERT (ID, Faculty_ID, Building_SchoolLevel_ID, StartDate, EndDate, LastModifiedOn)
  VALUES ((SELECT IFNULL(MAX(ID),0)+1 FROM schooltool_replica.FacultyBuildingSchoolLevel) , s.Faculty_ID, bsl_id, CURRENT_DATE(), NULL, CURRENT_TIMESTAMP());

-- Report IDs that were not found:
SELECT fid AS faculty_id_not_found
FROM UNNEST(faculty_ids) fid
LEFT JOIN schooltool_replica.Faculty f ON f.FacultyID = fid
WHERE f.ID IS NULL;

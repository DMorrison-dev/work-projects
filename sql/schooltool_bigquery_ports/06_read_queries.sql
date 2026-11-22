-- ===== AddressCleanup.py: fake (blank-street) addresses tied to exited students, with phone numbers to send to Geoffrey =====
SELECT DISTINCT a.ID AS AddressID, a.PhoneNumber
FROM schooltool_replica.Student s
JOIN schooltool_replica.StudentBuildingSchoolLevelEnrollment se ON s.ID = se.Student_ID
JOIN schooltool_replica.Person_Relationships pr ON s.Person_ID = pr.Person1_ID
JOIN schooltool_replica.Person_HomeInfo phi ON pr.Person2_ID = phi.Person_ID
JOIN schooltool_replica.Address a ON phi.Address_ID = a.ID
WHERE se.EndDate IS NOT NULL AND a.StreetName = '';
-- Delete (only after sending the phone numbers):
-- DELETE FROM schooltool_replica.Address WHERE ID IN (<ids above>);

-- ===== FacultyIDAssignment.py: faculty with blank FacultyID and an open building assignment, with their logins =====
SELECT DISTINCT f.ID AS Faculty_ID, u.Login
FROM schooltool_replica.Faculty f
JOIN schooltool_replica.FacultyBuildingSchoolLevel fbsl ON f.ID = fbsl.Faculty_ID AND fbsl.EndDate IS NULL
LEFT JOIN schooltool_replica.`User` u ON u.Person_ID = f.Person_ID
WHERE f.FacultyID = '';

-- ===== Securly.py: homeroom teacher + student IDs (HR courses, buildings 006/305) in one query =====
SELECT f.FacultyID, pf.LastName, ARRAY_AGG(DISTINCT s.StudentID ORDER BY s.StudentID) AS StudentIDs
FROM schooltool_replica.Student s
JOIN schooltool_replica.Sch_StudentSchedule sss ON s.ID = sss.Student_ID
JOIN schooltool_replica.CourseSchedule cs ON sss.Sch_Course_ID = cs.ID
JOIN schooltool_replica.Sch_Course sc ON cs.Sch_Course_ID = sc.ID
JOIN schooltool_replica.FacultySchedule fs ON cs.ID = fs.Sch_CourseSchedule_ID AND (fs.EndDate IS NULL OR fs.EndDate >= CURRENT_DATE())
JOIN schooltool_replica.Faculty f ON fs.Faculty_ID = f.ID
JOIN schooltool_replica.Person pf ON f.Person_ID = pf.ID
JOIN schooltool_replica.StudentBuildingSchoolLevelEnrollment se ON s.ID = se.Student_ID AND se.EnrollmentCategory_ID = 1 AND se.EndDate IS NULL
JOIN schooltool_replica.SchoolYear sy ON se.SchoolYear_ID = sy.ID AND sy.Active
JOIN schooltool_replica.BEDSProvider bp ON se.BEDSProvider_ID = bp.ID
WHERE sc.CourseID LIKE '%HR%' AND bp.ExportLocationCode IN ('006','305')
GROUP BY f.FacultyID, pf.LastName;

-- ===== BullseyeExport.py: one row per staff member per active building (replaces the per-row loop) =====
SELECT f.FacultyID AS EMPLID, p.FirstName, p.LastName, p.Email, b.Text AS Building, d.CleverID
FROM schooltool_replica.FacultyBuilding_SchoolLevel fb
JOIN schooltool_replica.Building_SchoolLevel bsl ON fb.Building_SchoolLevel_ID = bsl.ID
JOIN schooltool_replica.Building b ON bsl.Building_ID = b.ID AND b.Active AND b.ID NOT IN (36,45,46,47,52,50)
JOIN schooltool_replica.Faculty f ON fb.Faculty_ID = f.ID AND f.FacultyID <> '88888'
JOIN schooltool_replica.Person p ON f.Person_ID = p.ID AND p.Email LIKE '%@scsd.example'   -- '@scsd.us' on the real data
LEFT JOIN schooltool_replica.DigitalContentID d ON d.STIDTYPE = 'Faculty' AND d.STID = f.FacultyID  -- LEFT: missing Clever IDs show as NULL instead of crashing
WHERE fb.EndDate IS NULL
GROUP BY 1,2,3,4,5,6;

-- ===== AttendanceData.py (final draft, simplified to per-course/per-marking-period absence counts) =====
SELECT s.StudentID, sy.Text AS SchoolYear, sc.CourseID, cs.SectionNumber, sc.Name AS CourseName, d.Text AS Department,
  COUNTIF(i.ReportColumn = 1) AS MP1Absences, COUNTIF(i.ReportColumn = 2) AS MP2Absences,
  COUNTIF(i.ReportColumn = 3) AS MP3Absences, COUNTIF(i.ReportColumn = 4) AS MP4Absences
FROM schooltool_replica.Student_Period_Absense spa
JOIN schooltool_replica.Student s ON spa.Student_Id = s.ID
JOIN schooltool_replica.CourseSchedule cs ON spa.Sch_CourseSchedule_Id = cs.ID
JOIN schooltool_replica.Sch_Course sc ON cs.Sch_Course_ID = sc.ID
LEFT JOIN schooltool_replica.Department d ON sc.Department_ID = d.ID
JOIN schooltool_replica.SchoolYear sy ON sy.Active
JOIN schooltool_replica.StudentBuildingSchoolLevelEnrollment se
  ON se.Student_ID = s.ID AND se.SchoolYear_ID = sy.ID AND se.EnrollmentCategory_ID = 1 AND se.EndDate IS NULL
LEFT JOIN schooltool_replica.Grade_GradingInterval i
  ON i.Building_SchoolLevel_ID = se.Building_SchoolLevel_ID AND i.SchoolYear_ID = sy.ID
 AND spa.DateAbsent BETWEEN i.StartDate AND i.EndDate
WHERE spa.DateAbsent >= sy.StartDate AND sc.Department_ID <> 25
GROUP BY 1,2,3,4,5,6;

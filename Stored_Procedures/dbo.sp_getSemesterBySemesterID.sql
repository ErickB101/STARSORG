CREATE PROCEDURE [dbo].sp_getSemesterBySemesterID
	@semesterID nvarchar(4)
AS
	Select * From Semester
	where SemesterID=@semesterID

RETURN 0
CREATE PROCEDURE [dbo].sp_getCourseByCourseID
	@courseID nvarchar(15)
AS
	Select * From Course
	where CourseID=@courseID

RETURN 0
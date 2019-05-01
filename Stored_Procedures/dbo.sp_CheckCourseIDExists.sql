CREATE PROCEDURE [dbo].sp_CheckCourseIDExists
	@courseID nvarchar(15)
	,@recCount int =  0 OUTPUT	

AS
	BEGIN
		SET	@recCount=(Select count(0)From Course Where CourseID=@courseID)
		select @recCount as RecordCount
	End	
RETURN 0
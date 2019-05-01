CREATE PROCEDURE [dbo].sp_CheckSemesterIDExists
	@semesterID nvarchar(4)
	,@recCount int =  0 OUTPUT	

AS
	BEGIN
		SET	@recCount=(Select count(0)From Semester Where SemesterID=@semesterID)
		select @recCount as RecordCount
	End	
RETURN 0
CREATE PROCEDURE [dbo].sp_saveSemester
	@semesterID nvarchar(4)
	,@semesterDescription nvarchar(100)
AS
	Declare @countExists int
	Select @countExists =COUNT(0) from Semester where @semesterID=SemesterID

	if (@countExists=0)
	Begin
		insert into [dbo].Semester
			(
			SemesterID
			,SemesterDescription
			)

		Values
			(
			@semesterID
			,@semesterDescription
			)
		End
		Else
		Begin
			Update[dbo].Semester
			set	
				[SemesterDescription]=@semesterDescription
			where SemesterID=@semesterID
		END

RETURN @@error
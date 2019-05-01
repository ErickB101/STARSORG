CREATE PROCEDURE [dbo].sp_saveCourse
	@courseID nvarchar(10)
	,@courseName nvarchar(50)
AS
	Declare @countExists int
	Select @countExists =COUNT(0) from Course where @courseID=CourseID

	if (@countExists=0)
	Begin
		insert into [dbo].COURSE
			(
			CourseID
			,CourseName
			)

		Values
			(
			@courseID
			,@courseName
			)
		End
		Else
		Begin
			Update[dbo].COURSE
			set	
				[CourseName]=@courseName
			where CourseID=@courseID
		END

RETURN @@error
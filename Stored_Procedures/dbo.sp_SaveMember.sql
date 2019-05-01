CREATE PROCEDURE [dbo].sp_SaveMember
	@PID nvarchar(7)
	,@FName nvarchar(50)
	,@LName nvarchar(75)
	,@MI nvarchar(1)
	,@Email nvarchar(50)
	,@Phone nvarchar(13)
	,@PhotoPath nvarchar(300)
	,@RoleID nvarchar(15)
	,@SemesterID nvarchar(4)

AS
	Declare @countExists int
	SELECT @countExists=count(0) FROM MEMBER WHERE @PID=PID
	IF (@countExists=0)
	BEGIN
		INSERT INTO [dbo].MEMBER
			(
			PID
			,FName
			,LName
			,MI
			,Email
			,Phone
			,PhotoPath
			)
		VALUES
		(
		@PID
		,@FName
		,@LName
		,@MI
		,@Email
		,@Phone
		,@PhotoPath
		)
		INSERT into [dbo].MEMBER_ROLE
			(
			PID
			,RoleID
			,SemesterID
			)
		VALUES
		(
		@PID
		,@RoleID
		,@SemesterID
		)
	END
	ELSE
	BEGIN
		UPDATE [dbo].MEMBER
		SET
			[FName]=@FName
			,[LName]=@LName
			,[MI]=@MI
			,[Email]=@Email
			,[Phone]=@Phone
			,[PhotoPath]=@PhotoPath
		WHERE PID=@PID
	END
RETURN @@error
ALTER PROCEDURE [dbo].sp_saveSecurityProfile
	@pID nvarchar (7)
	,@userID nvarchar(15)
	,@password nvarchar(8)
	,@secRole nvarchar(10)
AS
	Declare @countExists int
	SELECT @countExists=count(0) FROM SECURITY WHERE @pID=PID
	IF (@countExists=0)
	BEGIN
		INSERT INTO [dbo].SECURITY
		(
		PID
		,UserID
		,Password
		,SecRole
		)
		VALUES
		(
		@pID
		,@userID
		,@password
		,@secRole
		)
	END
	ELSE
	BEGIN
		UPDATE [dbo].SECURITY
		SET
		[Password]=@password
		WHERE PID = @pID
	END
RETURN @@Error
CREATE PROCEDURE [dbo].sp_saveSecurityRole
    @pID NVARCHAR (7)
    ,@userID NVARCHAR (15)
    ,@password NVARCHAR (8) 
    ,@secRole  NVARCHAR (10)
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
		[SecRole]=@secRole
		WHERE PID = @pID
	END
RETURN @@Error
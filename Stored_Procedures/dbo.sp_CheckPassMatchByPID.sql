ALTER PROCEDURE [dbo].sp_CheckPassMatchByPID
	@pID nvarchar (7)
	,@password nvarchar (8)
AS
	BEGIN 
		IF EXISTS(SELECT * FROM SECURITY WHERE PID = @pID AND password = @password)
			SELECT 0 AS UserExists
		ELSE
			SELECT 1 AS UserExists
	END

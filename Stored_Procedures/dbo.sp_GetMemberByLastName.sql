CREATE PROCEDURE [dbo].sp_GetMemberByLastName
	@LName nvarchar(75)
AS
	SELECT LName, FName, PID
	FROM MEMBER
	WHERE LName LIKE @LName + '%'
RETURN 0
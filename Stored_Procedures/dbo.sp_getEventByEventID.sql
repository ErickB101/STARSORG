CREATE PROCEDURE [dbo].sp_getEventByEventID
	@eventID nvarchar(15)
AS
	SELECT *
	FROM EVENT
	WHERE EventID = @eventID
RETURN 0
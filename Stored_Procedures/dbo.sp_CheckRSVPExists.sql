CREATE PROCEDURE [dbo].sp_CheckRSVPExists
	@email nvarchar(50)
	,@eventID nvarchar(15)
	,@recCount int = 0 OUTPUT
	
AS
	BEGIN
	  SET @recCount = (Select count(0) FROM EVENT_RSVP WHERE Email=@email AND EventID=@eventID)
	  SELECT @recCount as RecordCount
	END
RETURN 0
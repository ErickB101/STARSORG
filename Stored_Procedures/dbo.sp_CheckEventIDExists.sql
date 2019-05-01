CREATE PROCEDURE [dbo].sp_CheckEventIDExists
	@eventID nvarchar(15)
	,@recCount int = 0 OUTPUT
	
AS
	BEGIN
	  SET @recCount = (Select count(0) FROM EVENT WHERE EventID=@eventID)
	  SELECT @recCount as RecordCount
	END
RETURN 0
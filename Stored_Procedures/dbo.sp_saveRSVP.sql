CREATE PROCEDURE [dbo].sp_saveRSVP
	@eventID nvarchar(15)
	,@firstName nvarchar(50)
	,@lastName nvarchar(75)
	,@email nvarchar(50)
AS
	Declare @countExists Int
	SELECT @countExists = count(0) FROM EVENT_RSVP WHERE @eventID = EventID AND @email = Email
	If (@countExists = 0)
	BEGIN
	  INSERT INTO [dbo].EVENT_RSVP
	    (EventID
	    ,FName
		,LName
		,Email
	    )
      VALUES
	    (@eventID
		,@firstName
		,@lastName
		,@email
		)

	  END
RETURN @@error
CREATE PROCEDURE [dbo].sp_saveEvent
	@eventID nvarchar(15)
	,@eventDescription nvarchar(500)
	,@eventTypeID nvarchar(15)
	,@semesterID nvarchar(4)
	,@startDate nvarchar(25)
	,@endDate nvarchar(25)
	,@location nvarchar(25)
AS
	Declare @countExists Int
	SELECT @countExists = count(0) FROM EVENT WHERE @eventID = EventID
	If (@countExists = 0)
	BEGIN
	  INSERT INTO [dbo].EVENT
	    (EventID
	    ,EventDescription
		,EventTypeID
		,SemesterID
		,StartDate
		,EndDate
		,Location
	    )
      VALUES
	    (@eventID
		,@eventDescription
		,@eventTypeID
		,@semesterID
		,@startDate
		,@endDate
		,@location
		)

	  END
	  ELSE
	  BEGIN
	    UPDATE [dbo].EVENT
		SET
          [EventDescription] = @eventDescription
		  ,[EventTypeID] = @eventTypeID
		  ,[SemesterID] = @semesterID
		  ,[StartDate] = @startDate
		  ,[EndDate] = @endDate
		  ,[Location] = @location
		WHERE EventID = @eventID
	  END

RETURN @@error
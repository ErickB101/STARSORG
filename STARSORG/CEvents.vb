'A second Event class, that connects the frmEvent and the CEvent class

Imports System.Data.SqlClient
Public Class CEvents
    Private _Event As CEvent
    
    
    'All of these fucntions and properties are shortcuts for the Event
    Public Sub New()
        'instantiate the CRole Object
        _Event = New CEvent
    End Sub

    Public ReadOnly Property CurrentObject() As CEvent
        Get
            Return _Event
        End Get
    End Property

    Public Sub Clear()
        _Event = New CEvent
    End Sub

    Public Sub CreateNewEvent()
        Clear()
        _Event.IsNewEvent = True
    End Sub

    Public Function Save() As Integer
        Return _Event.Save()
    End Function

    Public Function GetAllEvents() As SqlDataReader
        Dim objDR As SqlDataReader
        objDR = myDB.GetDataReaderBySP("sp_getAllEvents", Nothing)
        Return objDR
    End Function

    Public Function GetEventByEventID(strID As String) As CEvent
        Dim params As New ArrayList
        params.Add(New SqlParameter("eventID", strID))
        FillObject(myDB.GetDataReaderBySP("sp_getEventByEventID", params))
        Return _Event
    End Function

    Private Function FillObject(objDR As SqlDataReader) As CEvent
        If objDR.Read Then
            With _Event
                .EventID = objDR.Item("EventID")
                .EventDescription = objDR.Item("EventDescription")
                .TypeID = objDR.Item("EventTypeID")
                .SemesterID = objDR.Item("SemesterID")
                .StartDate = objDR.Item("StartDate")
                .EndDate = objDR.Item("EndDate")
                .Location = objDR.Item("Location")
            End With
        Else 'no record was returned
            'nothing to do here
        End If
        objDR.Close()
        Return _Event
    End Function
End Class

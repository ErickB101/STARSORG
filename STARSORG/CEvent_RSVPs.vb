Imports System.Data.SqlClient
Public Class CEvent_RSVPs
    Private _RSVP As CEvent_RSVP

    Public Sub New()
        'instantiate the CRole Object
        _RSVP = New CEvent_RSVP
    End Sub

    Public ReadOnly Property CurrentObject() As CEvent_RSVP
        Get
            Return _RSVP
        End Get
    End Property

    Public Sub Clear()
        _RSVP = New CEvent_RSVP
    End Sub

    Public Sub CreateNewRSVP()
        'call this when creating the edit portion of the screen to add a new role
        Clear()
        _RSVP.IsNewRSVP = True
    End Sub

    Public Function Save() As Integer
        Return _RSVP.Save()
    End Function

    Public Function GetAllEvents() As SqlDataReader
        Dim objDR As SqlDataReader
        objDR = myDB.GetDataReaderBySP("sp_getAllEvents", Nothing)
        Return objDR
    End Function

    Public Function GetEventByEventID(strID As String) As CEvent_RSVP
        Dim params As New ArrayList
        params.Add(New SqlParameter("eventID", strID))
        FillObject(myDB.GetDataReaderBySP("sp_getEventByEventID", params))
        Return _RSVP
    End Function

    Private Function FillObject(objDR As SqlDataReader) As CEvent_RSVP
        If objDR.Read Then
            With _RSVP
                .EventID = objDR.Item("EventID")
                .StartDate = objDR.Item("StartDate")
                .EndDate = objDR.Item("EndDate")
                .Location = objDR.Item("Location")
            End With
        Else 'no record was returned
            'nothing to do here
        End If
        objDR.Close()
        Return _RSVP
    End Function
End Class

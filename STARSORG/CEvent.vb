'This is code for an event form

Imports System.Data.SqlClient
Public Class CEvent
    'Represents a single record in the ROLE table
    Private _mstrEventID As String
    Private _mstrEventDescription As String
    Private _mstrTypeID As String
    Private _mstrSemesterID As String
    Private _mstrStartDate As String
    Private _mstrEndDate As String
    Private _mstrLocation As String

    Private _isNewEvent As Boolean
    'constructor
    Public Sub New()
        _mstrEventID = ""
        _mstrEventDescription = ""
        _mstrTypeID = ""
        _mstrSemesterID = ""
        _mstrStartDate = ""
        _mstrEndDate = ""
        _mstrLocation = ""

    End Sub
    
    'these properties help with storing and retrieving info from the database or Event form

#Region "Exposed Properties"
    Public Property EventID As String
        Get
            Return _mstrEventID
        End Get
        Set(strVal As String)
            _mstrEventID = strVal
        End Set
    End Property

    Public Property EventDescription As String
        Get
            Return _mstrEventDescription
        End Get
        Set(strVal As String)
            _mstrEventDescription = strVal
        End Set
    End Property

    Public Property TypeID As String
        Get
            Return _mstrTypeID
        End Get
        Set(strVal As String)
            _mstrTypeID = strVal
        End Set
    End Property

    Public Property SemesterID As String
        Get
            Return _mstrSemesterID
        End Get
        Set(strVal As String)
            _mstrSemesterID = strVal
        End Set
    End Property

    Public Property StartDate As String
        Get
            Return _mstrStartDate
        End Get
        Set(strVal As String)
            _mstrStartDate = strVal
        End Set
    End Property

    Public Property EndDate As String
        Get
            Return _mstrEndDate
        End Get
        Set(strVal As String)
            _mstrEndDate = strVal
        End Set
    End Property

    Public Property Location As String
        Get
            Return _mstrLocation
        End Get
        Set(strVal As String)
            _mstrLocation = strVal
        End Set
    End Property

    Public Property IsNewEvent As Boolean
        Get
            Return _isNewEvent
        End Get
        Set(blnVal As Boolean)
            _isNewEvent = blnVal
        End Set
    End Property
    
    'these properties assign the parameters for the saveEvent stored procedure and indicate whether its an update or an insert

    Public ReadOnly Property GetSaveParameters() As ArrayList
        'this property's code will create the parameters for the store procedure to save a record
        Get
            Dim params As New ArrayList
            params.Add(New SqlParameter("eventID", _mstrEventID))
            params.Add(New SqlParameter("eventDescription", _mstrEventDescription))
            params.Add(New SqlParameter("eventTypeID", _mstrTypeID))
            params.Add(New SqlParameter("semesterID", _mstrSemesterID))
            params.Add(New SqlParameter("startDate", _mstrStartDate))
            params.Add(New SqlParameter("endDate", _mstrEndDate))
            params.Add(New SqlParameter("location", _mstrLocation))
            Return params
        End Get
    End Property
#End Region



    Public Function Save() As Integer
        'return -1 if the ID already exits (and we cannot create a new record with duplicate ID)

        If IsNewEvent Then

            Dim params As New ArrayList

            params.Add(New SqlParameter("eventID", _mstrEventID))

            Dim strResult As String = myDB.GetSingleValueFromSP("sp_CheckEventIDExists", params)

            If Not strResult = 0 Then
                Return -1 'Not unique
            End If
        End If
        'if not a new role, or its new and has a unique ID, then do the save (update or insert)

        Return myDB.ExecSP("sp_saveEvent", GetSaveParameters())
    End Function
End Class

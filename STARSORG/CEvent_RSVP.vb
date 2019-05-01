Imports System.Data.SqlClient
Public Class CEvent_RSVP
    'Represents a single record in the ROLE table
    Private _mstrEventID As String
    Private _mstrFirstName As String
    Private _mstrLastName As String
    Private _mstrEmail As String
    Private _isNewRSVP As Boolean
    Private _mstrStartDate As String
    Private _mstrEndDate As String
    Private _mstrLocation As String
    'constructor
    Public Sub New()
        _mstrFirstName = ""
        _mstrLastName = ""
        _mstrEmail = ""
        _mstrEventID = ""
    End Sub

#Region "Exposed Properties"

    Public Property EventID As String
        Get
            Return _mstrEventID
        End Get
        Set(strVal As String)
            _mstrEventID = strVal
        End Set
    End Property

    Public Property FirstName As String
        Get
            Return _mstrFirstName
        End Get
        Set(strVal As String)
            _mstrFirstName = strVal
        End Set
    End Property

    Public Property LastName As String
        Get
            Return _mstrLastName
        End Get
        Set(strVal As String)
            _mstrLastName = strVal
        End Set
    End Property

    Public Property Email As String
        Get
            Return _mstrEmail
        End Get
        Set(strVal As String)
            _mstrEmail = strVal
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

    Public Property IsNewRSVP As Boolean
        Get
            Return _isNewRSVP
        End Get
        Set(blnVal As Boolean)
            _isNewRSVP = blnVal
        End Set
    End Property

    Public ReadOnly Property GetSaveParameters() As ArrayList
        'this property's code will create the parameters for the store procedure to save a record
        Get
            Dim params As New ArrayList
            params.Add(New SqlParameter("eventID", _mstrEventID))
            params.Add(New SqlParameter("firstName", _mstrFirstName))
            params.Add(New SqlParameter("lastName", _mstrLastName))
            params.Add(New SqlParameter("email", _mstrEmail))
            Return params
        End Get
    End Property
#End Region

    Public Function Save() As Integer
        'return -1 if the ID already exits (and we cannot create a new record with duplicate ID)

        If IsNewRSVP Then

            Dim params As New ArrayList

            params.Add(New SqlParameter("eventID", _mstrEventID))
            params.Add(New SqlParameter("email", _mstrEmail))

            Dim strResult As String = myDB.GetSingleValueFromSP("sp_CheckRSVPExists", params)

            If Not strResult = 0 Then
                Return -1 'Not unique
            End If
        End If
        'if not a new role, or its new and has a unique ID, then do the save (update or insert)

        Return myDB.ExecSP("sp_saveRSVP", GetSaveParameters())
    End Function

    Public Function GetReportData() As SqlDataAdapter
        Return myDB.GetDataAdapterBySP("dbo.sp_getAllRSVPs", Nothing)
    End Function

End Class

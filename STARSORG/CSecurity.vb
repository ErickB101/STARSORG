Imports System.Data.SqlClient
Public Class CSecurity
    Private _mintPID As Integer
    Private _mstrUserID As String
    Private _mstrPassword As String
    Private _mstrSecRole As String
    Private _isSecRole As Boolean

#Region "Exposed Properties"
    Public Sub New()
        _mstrUserID = ""
        _mstrPassword = ""
        _mstrSecRole = ""
    End Sub

    Public Property PID As Integer
        Get
            Return _mintPID
        End Get
        Set(intVal As Integer)
            _mintPID = intVal
        End Set
    End Property

    Public Property UserID As String
        Get
            Return _mstrUserID
        End Get
        Set(strVal As String)
            _mstrUserID = strVal
        End Set
    End Property

    Public Property Password As String
        Get
            Return _mstrPassword
        End Get
        Set(strVal As String)
            _mstrPassword = strVal
        End Set
    End Property

    Public Property SecRole As String
        Get
            Return _mstrSecRole
        End Get
        Set(strVal As String)
            _mstrSecRole = strVal
        End Set
    End Property

    Public Property IsNewSecRole As Boolean
        Get
            Return _isSecRole
        End Get
        Set(blnVal As Boolean)
            _isSecRole = blnVal
        End Set
    End Property

    Public ReadOnly Property GetSaveParameters() As ArrayList
        Get
            Dim params As New ArrayList
            params.Add(New SqlParameter("pID", _mintPID))
            params.Add(New SqlParameter("userID", _mstrUserID))
            params.Add(New SqlParameter("password", _mstrPassword))
            params.Add(New SqlParameter("secRole", _mstrSecRole))
            Return params
        End Get
    End Property
#End Region

    Public Function CheckUsernameAndPass(strUsername As String, strPass As String)
        Dim params As New ArrayList
        params.Add(New SqlParameter("UserID", strUsername))
        params.Add(New SqlParameter("Password", strPass))
        Return myDB.GetSingleValueFromSP("sp_CheckUsernameAndPass2", params)
    End Function

    Public Function Save() As Integer
        'return -1 if the PID already exists in SECURITY Table
        If IsNewSecRole Then
            Dim params As New ArrayList
            params.Add(New SqlParameter("pID", _mintPID))
            Dim strResult As String = myDB.GetSingleValueFromSP("sp_CheckPIDExists", params)
            If Not strResult = 0 Then
                Return -1 'not UNIQUE!!
            End If
        End If
        'if not new security profile. or it is new and has a unique ID, then update or save new
        Return myDB.ExecSP("sp_saveSecurityProfile", GetSaveParameters)
    End Function

    Public Function CheckPIDExists(strPID As String) As Integer
        'return -1 if the PID already exists in SECURITY Table
        Dim params As New ArrayList
        params.Add(New SqlParameter("pID", strPID))
        Dim strResult As String = myDB.GetSingleValueFromSP("sp_CheckPIDExists", params)
        If Not strResult = 0 Then
            Return -1 'not UNIQUE!!
        End If
        Return 0
    End Function


End Class

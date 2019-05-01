Imports System.Data.SqlClient
Public Class CMemberRole

    Private _mstrFName As String
    Private _mstrLName As String
    Private _mstrSemesterID As String
    Private _mstrRoleID As String
    Private _isNewRole As Boolean
    Private _mstrPID As String


    Public Sub New()
        _mstrRoleID = ""
        _mstrPID = ""
        _mstrFName = ""
        _mstrLName = ""
    End Sub
#Region "Exposed Properties"
    Public Property RoleId As String
        Get
            Return _mstrRoleID
        End Get
        Set(strVal As String)
            _mstrRoleID = strVal
        End Set
    End Property
    Public Property PID As String
        Get
            Return _mstrPID
        End Get
        Set(strVal As String)
            _mstrPID = strVal
        End Set
    End Property
    Public Property FName As String
        Get
            Return _mstrFName
        End Get
        Set(strVal As String)
            _mstrFName = strVal
        End Set
    End Property
    Public Property LName As String
        Get
            Return _mstrLName
        End Get
        Set(strVal As String)
            _mstrLName = strVal
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

    Public Property isNewRole As Boolean
        Get
            Return _isNewRole
        End Get
        Set(blnVal As Boolean)
            _isNewRole = blnVal
        End Set
    End Property
    Public ReadOnly Property getsaveParameters() As ArrayList

        Get
            Dim params As New ArrayList
            params.Add(New SqlParameter("PID", _mstrPID))
            params.Add(New SqlParameter("RoleID", _mstrRoleID))
            params.Add(New SqlParameter("SemesterID", _mstrSemesterID))
            Return params
        End Get

    End Property

#End Region

    Public Function Save() As Integer
        Return myDB.ExecSP("sp_saveMemberRoles", getsaveParameters())
    End Function

    'Public Function Delete() As Integer
    '    If isNewRole Then
    '        Dim params As New ArrayList
    '        params.Add(New SqlParameter("roleID", _mstrRoleID))
    '        Dim strResult As String = myDB.GetSingleValueFromSP("sp_CheckMemberRoleIDExists", params)
    '        If Not strResult = 0 Then
    '            Return -1
    '        End If
    '    End If
    '    Return myDB.ExecSP("sp_deleteMemberRoles", getdeleteParameters())
    'End Function

    Public Function GETREPORTDATA() As SqlDataAdapter
        Return myDB.GetDataAdapterBySP("dbo.sp_getMembersForReport", Nothing)

    End Function

End Class




Imports System.Data.SqlClient
Public Class CMembersRoles
    'Represents the MEMEBERS_ROLE table and its associated business rules

    Private _MemberRoles As CMemberRole


    Public Sub New()
        'INSTANTIATE THE CmemberRole object
        _MemberRoles = New CMemberRole
    End Sub
    Public ReadOnly Property currentObject() As CMemberRole
        Get
            Return _MemberRoles

        End Get
    End Property

    Public Sub clear()
        _MemberRoles = New CMemberRole
    End Sub
    Public Sub CreateNewMember()
        clear()
        _MemberRoles.isNewRole = True
    End Sub

    Public Function save() As Integer
        Return _MemberRoles.Save()
    End Function

    Public Function sp_getAllMembersBySemesterID() As SqlDataReader
        Dim objDR As SqlDataReader
        objDR = myDB.GetDataReaderBySP("sp_getMembersBySemesterID", Nothing)

        Return objDR
    End Function

    Public Function sp_getAllMembersForAdmin() As SqlDataReader
        Dim objDR As SqlDataReader
        objDR = myDB.GetDataReaderBySP("sp_getAllMembersForAdmin", Nothing)

        Return objDR
    End Function

    Public Function sp_GetAllSemesters() As SqlDataReader
        Dim objDR As SqlDataReader
        objDR = myDB.GetDataReaderBySP("sp_getAllSemesters", Nothing)

        Return objDR
    End Function
    Public Function sp_GetAllRoles() As SqlDataReader
        Dim objDR As SqlDataReader
        objDR = myDB.GetDataReaderBySP("sp_getAllRoles", Nothing)

        Return objDR
    End Function
    Public Function sp_GetAllMemberRoles() As SqlDataReader
        Dim objDR As SqlDataReader
        objDR = myDB.GetDataReaderBySP("sp_getAllMemberRoles", Nothing)

        Return objDR
    End Function
    Public Sub sp_DeleteRecord(strPID As String, strRoleID As String, strsemesterID As String)
        Dim params As New ArrayList
        params.Add(New SqlParameter("PID", strPID))
        params.Add(New SqlParameter("roleID", strRoleID))
        params.Add(New SqlParameter("semesterID", strsemesterID))
        myDB.ExecSP("sp_deleteMemberRoles", params)
    End Sub
    Public Function sp_getRoleIDByPID(strID As String, strSemesterID As String) As SqlDataReader
        Dim params As New ArrayList
        params.Add(New SqlParameter("PID", strID))
        params.Add(New SqlParameter("semesterID", strSemesterID))
        Return myDB.GetDataReaderBySP("sp_getRoleIDBypid", params)
    End Function

    Private Function FillObject(objDR As SqlDataReader) As CMemberRole

        If objDR.Read Then
            With _MemberRoles
                .PID = objDR.Item("PID")
                .RoleId = objDR.Item("RoleID")
                .SemesterID = objDR.Item("SemesterID")
            End With
        Else
        End If
        objDR.Close()
        Return _MemberRoles
    End Function

End Class

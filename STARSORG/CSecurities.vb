Imports System.Data.SqlClient
Public Class CSecurities
    Private _Security As CSecurity

    'constructor
    Public Sub New()
        'instantiate the CSecurity object
        _Security = New CSecurity
    End Sub

    Public ReadOnly Property CurrentObject() As CSecurity
        Get
            Return _Security
        End Get
    End Property

    Public Sub Clear()
        _Security = New CSecurity
    End Sub

    Public Function Save() As Integer
        Return _Security.Save()
    End Function

    Public Function CheckUsernameAndPass(strUsername As String, strPass As String)
        Return _Security.CheckUsernameAndPass(strUsername, strPass)
    End Function

    Public Function CheckPIDExists(strPID As String) As Integer
        Return _Security.CheckPIDExists(strPID)
    End Function
    Public Function GetAllSecMembers() As SqlDataReader
        Dim objDR As SqlDataReader
        objDR = myDB.GetDataReaderBySP("sp_getAllSecurities", Nothing)
        Return objDR
    End Function

    Public Function GetAllMembers() As SqlDataReader
        Dim objDR As SqlDataReader
        objDR = myDB.GetDataReaderBySP("sp_getAllMembersForAdmin", Nothing)
        Return objDR
    End Function

    Public Sub CreateNewSecRole()
        Clear()
        _Security.IsNewSecRole = True
    End Sub

    Private Function FillObject(objDR) As CSecurity
        If objDR.read Then
            With _Security
                .PID = objDR.Item("PID")
                .UserID = objDR.Item("UserID")
                .Password = objDR.Item("Password")
                .SecRole = objDR.Item("SecRole")
            End With
        Else 'no record was returned
            'nothing to do here
        End If
        objDR.Close()
        Return _Security
    End Function

    Public Function GetSecByUserID(strID As String) As CSecurity
        Dim params As New ArrayList
        params.Add(New SqlParameter("userID", strID))
        FillObject(myDB.GetDataReaderBySP("sp_getSecByUserID", params))
        Return _Security
    End Function

    Public Function GetSecByPID(strID As String) As CSecurity
        Dim params As New ArrayList
        params.Add(New SqlParameter("pID", strID))
        FillObject(myDB.GetDataReaderBySP("sp_getSecByPID", params))
        Return _Security
    End Function



End Class


Imports System.Data.SqlClient
Public Class CAudits
    'Represents the ROLE table and its associated business rules
    Private _Audit As CAudit

    'constructor
    Public Sub New()
        'instantiate the CAudit object
        _Audit = New CAudit
    End Sub

    Public ReadOnly Property CurrentObject() As CAudit
        Get
            Return _Audit
        End Get
    End Property

    Public Sub Clear()
        _Audit = New CAudit
    End Sub

    Public Sub isSuccess()
        'If UserID and Password match then give access to user
        _Audit.SUCCESS = True
    End Sub

    Public Function Save() As Integer
        Return _Audit.Save()
    End Function

    Public Function GetAllAudits() As SqlDataReader
        Dim objDR As SqlDataReader
        objDR = myDB.GetDataReaderBySP("sp_getAllAudits", Nothing)
        Return objDR
    End Function

    Public Function GetAuditByPID(intPID As Integer) As CAudit
        Dim params As New ArrayList
        params.Add(New SqlParameter("pID", intPID))
        FillObject(myDB.GetDataReaderBySP("sp_getAuditByPID", params))
        Return _Audit
    End Function

    Private Function FillObject(objDR) As CAudit
        If objDR.read Then
            With _Audit
                .PID = objDR.Item("PID")
                .ACCESSTIMESTAMP = objDR.Item("ACCESSTIMESTAMP")
                .SUCCESS = objDR.Item("SUCCESS")
            End With
        Else 'no record was returned
            'nothing to do here
        End If
        objDR.Close()
        Return _Audit
    End Function

End Class

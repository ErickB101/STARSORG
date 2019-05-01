Imports System.Data.SqlClient
Public Class CAudit
    Private _mintPID As Integer
    Private _mdtmAccessTimeStamp As DateTime
    Private _mblnSuccess As Boolean

#Region "Exposed Properties"
    Public Sub New()
    End Sub

    Public Property PID As Integer
        Get
            Return _mintPID
        End Get
        Set(intVal As Integer)
            _mintPID = intVal
        End Set
    End Property

    Public Property ACCESSTIMESTAMP As DateTime
        Get
            Return _mdtmAccessTimeStamp
        End Get
        Set(dtmVal As DateTime)
            _mdtmAccessTimeStamp = dtmVal
        End Set
    End Property

    Public Property SUCCESS As Boolean
        Get
            Return _mblnSuccess
        End Get
        Set(blnVal As Boolean)
            _mblnSuccess = blnVal
        End Set
    End Property

    Public ReadOnly Property GetSaveParameters() As ArrayList
        Get
            Dim params As New ArrayList
            params.Add(New SqlParameter("pID", _mintPID))
            params.Add(New SqlParameter("accessTimeStamp", _mdtmAccessTimeStamp))
            params.Add(New SqlParameter("success", _mblnSuccess))
            Return params
        End Get
    End Property
#End Region

    Public Function Save()
        Return myDB.GetSingleValueFromSP("sp_saveAuditGuest", Nothing)
    End Function

End Class

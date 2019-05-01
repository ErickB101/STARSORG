Imports System.Data.SqlClient
Public Class CDB
    Public Function OpenDB() As Boolean
        objSQLCommand = New SqlCommand 'instantiate the command object
        Dim blnResult As Boolean
        If objSQLConn Is Nothing Then 'we need to instantiate the connection object
            Try
                objSQLConn = New SqlConnection(gstrConn)
                objSQLConn.Open()
                blnResult = True
            Catch exOpenConnError As Exception
                MessageBox.Show("Cannot open database: " & exOpenConnError.ToString, "Connection error", MessageBoxButtons.OK, MessageBoxIcon.Error) 'fatal error
                blnResult = False
                'you should also log the error
            End Try
        Else
            If objSQLConn.State = ConnectionState.Open Then
                Return True
            Else
                Return False
            End If
        End If
        Return blnResult
    End Function

    Public Sub CloseDB()
        Try
            objSQLConn.Close()
        Catch ex As Exception
            MessageBox.Show("Error attempting to close database: " & ex.ToString, "Database error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Function GetDataReaderBySP(ByVal strSP As String, ByRef params As ArrayList) As SqlDataReader
        If Not OpenDB() Then
            'error log this
            Return Nothing
        End If
        Dim sqlComm As New SqlCommand(strSP, objSQLConn)
        sqlComm.CommandType = CommandType.StoredProcedure   'very important
        If Not params Is Nothing Then 'add each parameter into the command object
            For Each p As SqlParameter In params
                sqlComm.Parameters.Add(p)
            Next
        End If
        Try
            Return sqlComm.ExecuteReader()
        Catch ex As Exception
            MessageBox.Show("Failed to get datareader: " & ex.ToString, "Database error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try
    End Function

    Public Function GetDataAdapterBySP(ByVal strSP As String, ByRef params As ArrayList) As SqlDataAdapter
        Dim sqlComm As New SqlCommand(strSP, objSQLConn)
        Dim sqlDA As SqlDataAdapter
        sqlComm.CommandType = CommandType.StoredProcedure
        If Not params Is Nothing Then
            For Each p As SqlParameter In params
                sqlComm.Parameters.Add(p)
            Next
        End If
        Try
            sqlDA = New SqlDataAdapter(sqlComm)
            Return sqlDA
        Catch ex As Exception
            MessageBox.Show("Failed to get data adapter: " & ex.ToString, "Database error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try
    End Function

    Public Function ExecSP(ByVal strSP As String, ByRef params As ArrayList) As Integer
        If Not OpenDB() Then
            Return -1
        End If

        objSQLCommand = New SqlCommand(strSP, objSQLConn)
        objSQLCommand.CommandType = CommandType.StoredProcedure

        Try
            If Not params Is Nothing Then
                For Each p As SqlParameter In params
                    objSQLCommand.Parameters.Add(p)
                Next
            End If
            Return objSQLCommand.ExecuteNonQuery()

        Catch ex As Exception
            MessageBox.Show("Error executing SP:" & ex.ToString, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return -1
        End Try
    End Function

    Public Function GetSingleValueFromSP(ByVal strSP As String, ByRef params As ArrayList) As String
        Dim sqlDR As SqlDataReader = GetDataReaderBySP(strSP, params)
        Dim strResult As String
        If Not sqlDR Is Nothing Then 'we go a data reader
            If sqlDR.Read Then 'we got some data
                strResult = sqlDR.GetValue(0).ToString
                sqlDR.Close()
                Return strResult
            End If
            sqlDR.Close()
            Return "-1"
        End If
        Return "-2" 'failed to connect or even get a datareader
    End Function
End Class

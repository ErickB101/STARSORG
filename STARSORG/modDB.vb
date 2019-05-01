Imports System.Data.SqlClient
Module modDB
    Public gstrConn As String = "Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=D:\STARSORG\STARSDB.mdf;Integrated Security=True"
    Public objSQLConn As SqlConnection
    Public objSQLCommand As SqlCommand
End Module

Imports System.Data.SqlClient

Public Class CSemesters
    'Represents the Semester table and its associated Business rules

    Private _Semester As CSemester

    Public Sub New()
        'instantiate the CSemester Object 
        _Semester = New CSemester

    End Sub

    Public ReadOnly Property CurrentObject() As CSemester
        Get
            Return _Semester
        End Get

    End Property

    Public Sub Clear()
        _Semester = New CSemester
    End Sub

    Public Sub CreateNewSemester()
        'call this when clearing the edit portion of the screen to add a new Semester
        Clear()
        _Semester.IsNewSemester = True

    End Sub

    Public Function Save() As Integer

        'Put checks to be able to get the specific data that you want such as 
        'The txtSemesterID.Text has to be at least 3 Characters long (This is an ***EXAMPLE***)


        Return _Semester.Save()
    End Function

    Public Function GetAllSemesters() As SqlDataReader
        Dim objDR As SqlDataReader
        objDR = myDB.GetDataReaderBySP("sp_getAllSemesters", Nothing)
        Return objDR
    End Function

    Public Function GetSemesterBySemesterID(strID As String) As CSemester
        Dim params As New ArrayList
        'Dim objDR As SqlDataReader
        params.Add(New SqlParameter("semesterID", strID))

        'objDR = myDB.GetDataReaderBySP("sp_getSemesterBySemesterID", params)
        FillObject(myDB.GetDataReaderBySP("sp_getSemesterBySemesterID", params))
        'Return objDR
        Return _Semester

    End Function

    Private Function FillObject(objDR As SqlDataReader) As CSemester

        If objDR.Read Then

            With _Semester
                .SemesterID = objDR.Item("SemesterID")
                .SemesterDescription = objDR.Item("SemesterDescription")
            End With

        Else 'no record 

        End If

        objDR.Close()
        Return _Semester

    End Function

End Class

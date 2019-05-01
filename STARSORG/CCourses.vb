Imports System.Data.SqlClient

Public Class CCourses
    'Represents the Course table and its associated Business rules

    Private _Course As CCourse

    Public Sub New()
        'instantiate the CCourse Object 
        _Course = New CCourse

    End Sub

    Public ReadOnly Property CurrentObject() As CCourse
        Get
            Return _Course
        End Get

    End Property

    Public Sub Clear()
        _Course = New CCourse
    End Sub

    Public Sub CreateNewCourse()
        'call this when clearing the edit portion of the screen to add a new Course
        Clear()
        _Course.IsNewCourse = True

    End Sub

    Public Function Save() As Integer

        'Put checks to be able to get the specific data that you want such as 
        'The txtCourseID.Text has to be at least 3 Characters long (This is an ***EXAMPLE***)


        Return _Course.Save()
    End Function

    Public Function GetAllCourses() As SqlDataReader
        Dim objDR As SqlDataReader
        objDR = myDB.GetDataReaderBySP("sp_getAllCourses", Nothing)
        Return objDR
    End Function


    Public Function GetCourseByCourseID(strID As String) As CCourse
        Dim params As New ArrayList
        'Dim objDR As SqlDataReader
        params.Add(New SqlParameter("courseID", strID))

        'objDR = myDB.GetDataReaderBySP("sp_getCourseByCourseID", params)
        FillObject(myDB.GetDataReaderBySP("sp_getCourseByCourseID", params))
        'Return objDR
        Return _Course

    End Function

    Private Function FillObject(objDR As SqlDataReader) As CCourse

        If objDR.Read Then

            With _Course
                .CourseID = objDR.Item("CourseID")
                .CourseName = objDR.Item("CourseName")
            End With

        Else 'no record 

        End If

        objDR.Close()
        Return _Course

    End Function

End Class

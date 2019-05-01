Imports Microsoft.Reporting.WinForms
Imports System.Data.SqlClient

Public Class frmCourseReport

    Private ds As DataSet
    Private da As SqlDataAdapter
    Private Course As CCourse

    Private Sub frmCourseReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.rpvReport.RefreshReport()
    End Sub


    Public Sub Display()
        Course = New CCourse
        rpvReport.LocalReport.ReportPath = AppDomain.CurrentDomain.BaseDirectory & "Reports\rptCourseReport.rdlc"

        ds = New DataSet
        da = Course.GetReportData()
        da.Fill(ds)

        rpvReport.LocalReport.DataSources.Add(New ReportDataSource("dsCourses", ds.Tables(0)))

        rpvReport.SetDisplayMode(DisplayMode.PrintLayout)

        rpvReport.RefreshReport()
        Me.Cursor = Cursors.Default

        Me.ShowDialog()

    End Sub

End Class
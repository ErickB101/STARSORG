Imports Microsoft.Reporting.WinForms
Imports System.Data.SqlClient

Public Class frmSemesterReport

    Private ds As DataSet
    Private da As SqlDataAdapter
    Private Semester As CSemester


    Private Sub frmSemesterReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.rpvReport.RefreshReport()
    End Sub

    Public Sub Display()
        Semester = New CSemester
        rpvReport.LocalReport.ReportPath = AppDomain.CurrentDomain.BaseDirectory & "Reports\rptSemesterReport.rdlc"

        ds = New DataSet
        da = Semester.GetReportData()
        da.Fill(ds)

        rpvReport.LocalReport.DataSources.Add(New ReportDataSource("dsSemesters", ds.Tables(0)))

        rpvReport.SetDisplayMode(DisplayMode.PrintLayout)

        rpvReport.RefreshReport()
        Me.Cursor = Cursors.Default

        Me.ShowDialog()

    End Sub


End Class
Imports Microsoft.Reporting.WinForms
Imports System.Data.SqlClient
Public Class frmMemberReport
    Private ds As DataSet
    Private da As SqlDataAdapter
    Private Member As CMember

    Private Sub frmMemberReport_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.rpvReport.RefreshReport()
    End Sub
    Public Sub Display()
        Member = New CMember
        rpvReport.LocalReport.ReportPath = AppDomain.CurrentDomain.BaseDirectory & "Reports\rptMemberReport.rdlc"
        ds = New DataSet
        da = Member.GetReportData()
        da.Fill(ds)
        rpvReport.LocalReport.DataSources.Add(New ReportDataSource("dsMembers", ds.Tables(0)))
        rpvReport.SetDisplayMode(DisplayMode.PrintLayout)
        rpvReport.RefreshReport()
        Me.Cursor = Cursors.Default
        Me.ShowDialog()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) 
        Me.Close()
    End Sub
End Class
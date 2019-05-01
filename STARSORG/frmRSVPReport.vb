Imports Microsoft.Reporting.WinForms
Imports System.Data.SqlClient
Public Class frmRSVPReport
    Private ds As DataSet
    Private da As SqlDataAdapter
    Private RSVP As CEvent_RSVP
    Private Sub frmRSVPReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.rpvReport.RefreshReport()
    End Sub

    Public Sub Display()
        RSVP = New CEvent_RSVP
        rpvReport.LocalReport.ReportPath = AppDomain.CurrentDomain.BaseDirectory & "Reports\rptRSVPReport.rdlc"
        ds = New DataSet
        da = RSVP.GetReportData()
        da.Fill(ds)
        rpvReport.LocalReport.DataSources.Add(New ReportDataSource("dsRSVPs", ds.Tables(0)))
        rpvReport.SetDisplayMode(DisplayMode.PrintLayout)
        rpvReport.RefreshReport()
        Me.Cursor = Cursors.Default
        Me.ShowDialog()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) 
        Me.Close()
    End Sub
End Class
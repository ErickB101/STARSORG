Imports Microsoft.Reporting.WinForms
Imports System.Data.SqlClient

Public Class frmRoleReport
    Private ds As DataSet
    Private da As SqlDataAdapter
    Private Role As CRole

    Private Sub frmRoleReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.rpvReport.RefreshReport()
    End Sub

    Public Sub Display()
        Role = New CRole
        rpvReport.LocalReport.ReportPath = AppDomain.CurrentDomain.BaseDirectory & "Reports\rptRoleReport.rdlc"
        ds = New DataSet
        da = Role.GetReportData()
        da.Fill(ds)
        'rpvReport.LocalReport.DataSources.Add(New ReportDataSource("dsRoles", dsTables(0)))
        rpvReport.SetDisplayMode(DisplayMode.PrintLayout)
        rpvReport.RefreshReport()
        Me.Cursor = Cursors.Default
        Me.ShowDialog()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs)
        Dim RoleReport As New frmRoleReport
        Me.Cursor = Cursors.WaitCursor
        RoleReport.Display()
        Me.Cursor = Cursors.Default
    End Sub
End Class
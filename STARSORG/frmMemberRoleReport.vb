Imports Microsoft.Reporting.WinForms
Imports System.Data.SqlClient
Public Class frmMemberRoleReport
    Private Ds As DataSet
    Private Da As SqlDataAdapter
    Private MemberRole As CMemberRole

    Private Sub frmMemberRoleReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.rpvReport.RefreshReport()
    End Sub
    Public Sub Display()
        MemberRole = New CMemberRole
        rpvReport.LocalReport.ReportPath = AppDomain.CurrentDomain.BaseDirectory & "Reports\rptMemberRoleReport.rdlc"
        Ds = New DataSet
        Da = MemberRole.GETREPORTDATA()
        Da.Fill(Ds)
        rpvReport.LocalReport.DataSources.Add(New ReportDataSource("dsMemberRoles", Ds.Tables(0)))
        rpvReport.SetDisplayMode(DisplayMode.PrintLayout)
        rpvReport.RefreshReport()
        Me.Cursor = Cursors.Default
        Me.ShowDialog()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs)
        Close()
    End Sub
End Class
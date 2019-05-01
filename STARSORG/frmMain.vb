Public Class frmMain
    Private Login As frmLogin
    Private Admin As frmAdmin
    Private Course As frmCourse
    Private EventInfo As frmEvent
    Private Member As frmMember
    Private RoleInfo As frmRole
    Private RSVP As frmRSVP
    Private Semester As frmSemester
    Private MemberRoles As frmMemberRoles
    Private objCurrentUser As CSecurities

#Region "Toolbar Stuff"
    Private Sub tsbProxy_MouseEnter(sender As Object, e As EventArgs) Handles tsbCourse.MouseEnter, tsbEvent.MouseEnter, tsbHelp.MouseEnter, tsbHome.MouseEnter, tsbLogOut.MouseEnter, tsbMember.MouseEnter, tsbRole.MouseEnter, tsbRSVP.MouseEnter, tsbSemester.MouseEnter, tsbTutor.MouseEnter, tsbAdmin.MouseEnter, tsbMemberRoles.MouseEnter
        'we need to do this only because we are not putting the images in our Image property of the toolbar buttons
        Dim tsbProxy As ToolStripButton
        tsbProxy = DirectCast(sender, ToolStripButton)
        tsbProxy.DisplayStyle = ToolStripItemDisplayStyle.Text
    End Sub
    Private Sub tsbProxy_MouseLeave(sender As Object, e As EventArgs) Handles tsbCourse.MouseLeave, tsbEvent.MouseLeave, tsbHelp.MouseLeave, tsbHome.MouseLeave, tsbLogOut.MouseLeave, tsbMember.MouseLeave, tsbRole.MouseLeave, tsbRSVP.MouseLeave, tsbSemester.MouseLeave, tsbTutor.MouseLeave, tsbAdmin.MouseLeave, tsbMemberRoles.MouseLeave
        Dim tsbProxy As ToolStripButton
        tsbProxy = DirectCast(sender, ToolStripButton)
        tsbProxy.DisplayStyle = ToolStripItemDisplayStyle.Image
    End Sub

    Private Sub tsbCourse_Click(sender As Object, e As EventArgs) Handles tsbCourse.Click
        Me.Hide()
        Course.ShowDialog()
        Me.Show()
        PerformNextAction()
    End Sub

    Private Sub tsbEvent_Click(sender As Object, e As EventArgs) Handles tsbEvent.Click
        Me.Hide()
        EventInfo.ShowDialog()
        Me.Show()
        PerformNextAction()
    End Sub

    Private Sub tsbMember_Click(sender As Object, e As EventArgs) Handles tsbMember.Click
        Me.Hide()
        Member.ShowDialog()
        Me.Show()
        PerformNextAction()
    End Sub

    Private Sub tsbRole_Click(sender As Object, e As EventArgs) Handles tsbRole.Click
        Me.Hide()
        RoleInfo.ShowDialog()
        Me.Show()
        PerformNextAction()
    End Sub
    Private Sub tsbRSVP_Click(sender As Object, e As EventArgs) Handles tsbRSVP.Click
        Me.Hide()
        RSVP.ShowDialog()
        Me.Show()
        PerformNextAction()
    End Sub

    Private Sub tsbSemester_Click(sender As Object, e As EventArgs) Handles tsbSemester.Click
        Me.Hide()
        Semester.ShowDialog()
        Me.Show()
        PerformNextAction()
    End Sub

    Private Sub tsbAdmin_Click(sender As Object, e As EventArgs) Handles tsbAdmin.Click
        Me.Hide()
        Admin.ShowDialog()
        Me.Show()
        PerformNextAction()
    End Sub

    Private Sub tsbLogOut_Click(sender As Object, e As EventArgs) Handles tsbLogOut.Click
        EndProgram()
    End Sub

    Private Sub tsbMemberRoles_Click(sender As Object, e As EventArgs) Handles tsbMemberRoles.Click
        Me.Hide()
        MemberRoles.ShowDialog()
        Me.Show()
        PerformNextAction()
    End Sub
#End Region

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        Login = New frmLogin
        Admin = New frmAdmin
        Course = New frmCourse
        EventInfo = New frmEvent
        Member = New frmMember
        RoleInfo = New frmRole
        RSVP = New frmRSVP
        Semester = New frmSemester
        MemberRoles = New frmMemberRoles
        objCurrentUser = New CSecurities
        Try
            myDB.OpenDB()
        Catch ex As Exception
            MessageBox.Show("Unable to open database. Connection string = " & gstrConn, "DB Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            EndProgram()
        End Try

        'If user is not logged in, show login page
        ShowLogin()

        'Retrieve the current user after login
        Dim userSecRole As String
        Try
            objCurrentUser = Login.CurrentUser
            userSecRole = objCurrentUser.CurrentObject.SecRole.ToString
        Catch ex As Exception
            MessageBox.Show("Unable to retrieve current user " & ex.ToString, "Database error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try

        'Apply security restrictions
        If userSecRole = "Guest" Or userSecRole = "Member" Then 'can only access RSVP
            tsbMember.Enabled = False
            tsbSemester.Enabled = False
            tsbCourse.Enabled = False
            tsbEvent.Enabled = False
            tsbMemberRoles.Enabled = False
            tsbRole.Enabled = False
            tsbAdmin.Enabled = False
        ElseIf userSecRole = "Officer" Then 'can access everything accept ADMIN form
            tsbAdmin.Enabled = False
        End If
    End Sub

    Public Sub EndProgram()
        'close each form except main
        Dim f As Form
        Me.Cursor = Cursors.WaitCursor
        For Each f In Application.OpenForms
            If f.Name <> Me.Name Then
                If Not f Is Nothing Then
                    f.Close()
                End If
            End If
        Next
        'close db connection
        If Not objSQLConn Is Nothing Then
            objSQLConn.Close()
            objSQLConn.Dispose()
        End If
        Me.Cursor = Cursors.Default
        Application.Exit()
    End Sub

    Private Sub PerformNextAction()
        'get the next action specified on the child form, and then simulate the click of that button here
        Select Case intNextAction
            Case ACTION_COURSE
                tsbCourse.PerformClick()
            Case ACTION_EVENT
                tsbEvent.PerformClick()
            Case ACTION_HELP
                tsbHelp.PerformClick()
            Case ACTION_HOME
                tsbHome.PerformClick()
            Case ACTION_LOGOUT
                tsbLogOut.PerformClick()
            Case ACTION_MEMBER
                tsbMember.PerformClick()
            Case ACTION_NONE
                'nothing to do here
            Case ACTION_ROLE
                tsbRole.PerformClick()
            Case ACTION_RSVP
                tsbRSVP.PerformClick()
            Case ACTION_SEMESTER
                tsbSemester.PerformClick()
            Case ACTION_TUTOR
                tsbTutor.PerformClick()
            Case ACTION_ADMIN
                tsbAdmin.PerformClick()
            Case ACTION_MEMBERROLES
                tsbMemberRoles.PerformClick()
            Case Else
                MessageBox.Show("Unexpected case value in frmMain:PerformNextAction", "Program Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Select
    End Sub

    Private Sub ShowLogin()
        If Login.IsValidLogin = False Then
            Me.Hide()
            Login.ShowDialog()
            Exit Sub
        End If
    End Sub


End Class

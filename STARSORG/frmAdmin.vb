Imports System.Data.SqlClient

Public Class frmAdmin
    Private objSecurities As CSecurities
    Private blnClearing As Boolean
    Private blnReloading As Boolean

#Region "Toolbar stuff"
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
        intNextAction = ACTION_COURSE
        Me.Hide()
    End Sub

    Private Sub tsbEvent_Click(sender As Object, e As EventArgs) Handles tsbEvent.Click
        intNextAction = ACTION_EVENT
        Me.Hide()
    End Sub

    Private Sub tsbHelp_Click(sender As Object, e As EventArgs) Handles tsbHelp.Click
        intNextAction = ACTION_HELP
        Me.Hide()
    End Sub

    Private Sub tsbHome_Click(sender As Object, e As EventArgs) Handles tsbHome.Click
        intNextAction = ACTION_HOME
        Me.Hide()
    End Sub

    Private Sub tsbLogOut_Click(sender As Object, e As EventArgs) Handles tsbLogOut.Click
        intNextAction = ACTION_LOGOUT
        Me.Hide()
    End Sub

    Private Sub tsbMember_Click(sender As Object, e As EventArgs) Handles tsbMember.Click
        intNextAction = ACTION_MEMBER
        Me.Hide()
    End Sub

    Private Sub tsbRole_Click(sender As Object, e As EventArgs) Handles tsbRole.Click
        intNextAction = ACTION_ROLE
        Me.Hide()
    End Sub

    Private Sub tsbRSVP_Click(sender As Object, e As EventArgs) Handles tsbRSVP.Click
        intNextAction = ACTION_RSVP
        Me.Hide()
    End Sub

    Private Sub tsbSemester_Click(sender As Object, e As EventArgs) Handles tsbSemester.Click
        intNextAction = ACTION_SEMESTER
        Me.Hide()
    End Sub

    Private Sub tsbTutor_Click(sender As Object, e As EventArgs) Handles tsbTutor.Click
        intNextAction = ACTION_TUTOR
        Me.Hide()
    End Sub
    Private Sub tsbAdmin_Click(sender As Object, e As EventArgs) Handles tsbAdmin.Click
        'already here
    End Sub

    Private Sub tsbMemberRole_Click(sender As Object, e As EventArgs) Handles tsbMemberRoles.Click
        intNextAction = ACTION_MEMBERROLES
        Me.Hide()
    End Sub
#End Region
#Region "Textboxes"
    Private Sub txtAll_GotFocus(sender As Object, e As EventArgs)
        Dim txtBox As TextBox
        txtBox = DirectCast(sender, TextBox)
        txtBox.SelectAll()
    End Sub

    Private Sub txtAll_LostFocus(sender As Object, e As EventArgs)
        Dim txtBox As TextBox
        txtBox = DirectCast(sender, TextBox)
        txtBox.DeselectAll()
    End Sub
#End Region
    Private Sub LoadMembersByPID()
        Dim objReader As SqlDataReader
        lstMembers.Items.Clear()
        Try
            objReader = objSecurities.GetAllMembers()
            Do While objReader.Read
                lstMembers.Items.Add(objReader.Item("PID"))
            Loop
            objReader.Close()
        Catch ex As Exception
            'already handled in CDB
        End Try
        If objSecurities.CurrentObject.PID <> Nothing Then
            lstMembers.SelectedIndex = lstMembers.FindStringExact(objSecurities.CurrentObject.PID)
        End If
        blnReloading = False
    End Sub

    Private Sub frmAdmin_Load(sender As Object, e As EventArgs) Handles Me.Load
        objSecurities = New CSecurities
    End Sub

    Private Sub frmAdmin_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        blnClearing = True
        ClearScreenControls(Me)
        blnReloading = True
        LoadMembersByPID()
        grpNew.Enabled = False
        grpReset.Enabled = False
        blnClearing = False
        blnReloading = False
    End Sub

    Private Sub lstMembers_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstMembers.SelectedIndexChanged
        If blnClearing Then
            Exit Sub
        End If
        If blnReloading Then
            tslStatus.Text = ""
            Exit Sub
        End If
        If lstMembers.SelectedIndex = -1 Then 'nothing to do
            Exit Sub
        End If
        If radReset.Checked Then
            LoadResetRecord()
        End If
        If radNew.Checked Then
            LoadSecurityRecord()
        End If
    End Sub

    Private Sub LoadResetRecord()
        Try
            objSecurities.GetSecByPID(lstMembers.SelectedItem.ToString)
            With objSecurities.CurrentObject
                txtUserID.Text = .UserID
            End With
        Catch ex As Exception
            MessageBox.Show("Error loading Role record: " & ex.ToString, "Program Error", MessageBoxButtons.OK, MessageBoxIcon.Error)  'ex.ToString for debug purposes
        End Try
    End Sub

    Private Sub LoadSecurityRecord()
        Dim strPID As String
        Dim intResult As Integer
        strPID = lstMembers.SelectedItem.ToString
        intResult = objSecurities.CheckPIDExists(strPID)
        Try
            If intResult = -1 Then
                objSecurities.GetSecByPID(strPID)
            Else
                objSecurities.Clear()
            End If

            With objSecurities.CurrentObject
                txtInitUserID.Text = .UserID
                txtInitPass.Text = .Password
                txtInitSecRole.Text = .SecRole
            End With
        Catch ex As Exception
            tslStatus.Text = "New User"
        End Try
    End Sub

    Private Sub radAll_CheckedChanged(sender As Object, e As EventArgs) Handles radNew.CheckedChanged, radReset.CheckedChanged
        If blnClearing Then
            Exit Sub
        End If
        Clear()

        If radReset.Checked Then
            grpReset.Enabled = True
            grpNew.Enabled = False
            LoadResetRecord()
            txtUserID.Focus()
            txtNewUserPass.Focus()
            objSecurities.CurrentObject.IsNewSecRole = False
        End If

        If radNew.Checked Then
            grpNew.Enabled = True
            grpReset.Enabled = False
            LoadSecurityRecord()
            objSecurities.CreateNewSecRole()
            objSecurities.CurrentObject.PID = lstMembers.SelectedItem.ToString
            txtInitPass.Focus()
            txtInitSecRole.Focus()
            txtInitUserID.Focus()
        End If

    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        blnClearing = True
        Clear()
        radNew.Checked = False
        radReset.Checked = False
        objSecurities = New CSecurities
        errP.Clear()
        If lstMembers.SelectedIndex <> -1 Then
            If radNew.Checked Then
                LoadResetRecord() 'reload what was not selected in case user messed up the form
            End If
            If radReset.Checked Then
                LoadSecurityRecord()
            End If
        Else
            grpNew.Enabled = False
            grpReset.Enabled = False
        End If
        blnClearing = False
        objSecurities.CurrentObject.IsNewSecRole = False
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim intResult As Integer
        Dim blnErrors As Boolean
        '-----add your validation code here----'
        If radReset.Checked Then 'validate grpReset
            If Not ValidateTextBoxLength(txtUserID, errP) Then
                blnErrors = True
            End If
            If Not ValidateTextBoxLength(txtNewUserPass, errP) Then
                blnErrors = True
            End If
            With objSecurities.CurrentObject
                .Password = txtNewUserPass.Text
            End With
        ElseIf radNew.Checked Then 'validate grpNew
            If Not ValidateTextBoxLength(txtInitUserID, errP) Then
                blnErrors = True
            End If
            If Not ValidateTextBoxLength(txtInitPass, errP) Then
                blnErrors = True
            End If
            If Not ValidateTextBoxLength(txtInitSecRole, errP) Then
                blnErrors = True
            End If
            Dim SecRole As String = txtInitSecRole.Text
            If Not SecRole = "Admin" Or Not SecRole = "Officer" Or Not SecRole = "Member" Or Not SecRole = "Guest" Then
                MessageBox.Show("Security role but be valid. Unable to save Security record", "Database error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                tslStatus.Text = "Error"
                blnErrors = True
            End If
            With objSecurities.CurrentObject
                .UserID = txtInitUserID.Text
                .Password = txtInitPass.Text
                .SecRole = txtInitSecRole.Text
            End With
        End If
        If blnErrors = True Then
            btnCancel.PerformClick()
            Exit Sub
        End If

        Try
            Me.Cursor = Cursors.WaitCursor
            objSecurities.CurrentObject.PID =
            intResult = objSecurities.Save
            If intResult = 1 Then
                tslStatus.Text = "Security record saved successfully"
            End If
            If intResult = -1 Then
                MessageBox.Show("PID must be unique. Unable to save Security record", "Database error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                tslStatus.Text = "Error"
            End If
        Catch ex As Exception
            MessageBox.Show("Unable to save Security record: " & ex.ToString, "Database error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            tslStatus.Text = "Error"
        End Try
        Me.Cursor = Cursors.Default

        blnReloading = True
        LoadMembersByPID()
        radNew.Checked = False
        radReset.Checked = False
        objSecurities.CurrentObject.IsNewSecRole = False
    End Sub

    Private Sub Clear()
        txtUserID.Text = ""
        txtNewUserPass.Text = ""
        txtInitUserID.Text = ""
        txtInitPass.Text = ""
        txtInitSecRole.Text = ""
        tslStatus.Text = ""
    End Sub

End Class
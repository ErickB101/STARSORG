Imports System.Data.SqlClient

Public Class frmRole
    Private objRoles As CRoles
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
        'already here
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
        intNextAction = ACTION_ADMIN
        Me.Hide()
    End Sub

    Private Sub tsbMemberRole_Click(sender As Object, e As EventArgs) Handles tsbMemberRoles.Click
        intNextAction = ACTION_MEMBERROLES
        Me.Hide()
    End Sub

#End Region
#Region "Textboxes"
    Private Sub txtAll_GotFocus(sender As Object, e As EventArgs) Handles txtRoleID.GotFocus, txtDesc.GotFocus
        Dim txtBox As TextBox
        txtBox = DirectCast(sender, TextBox)
        txtBox.SelectAll()
    End Sub

    Private Sub txtAll_LostFocus(sender As Object, e As EventArgs) Handles txtRoleID.LostFocus, txtDesc.LostFocus
        Dim txtBox As TextBox
        txtBox = DirectCast(sender, TextBox)
        txtBox.DeselectAll()
    End Sub
#End Region

    Private Sub LoadRoles()
        Dim objReader As SqlDataReader
        lstRoles.Items.Clear()
        Try
            objReader = objRoles.GetAllRoles
            Do While objReader.Read
                lstRoles.Items.Add(objReader.Item("RoleID"))
            Loop
            objReader.Close()
        Catch ex As Exception
            'already handled in CDB
        End Try
        If objRoles.CurrentObject.RoleID <> "" Then
            lstRoles.SelectedIndex = lstRoles.FindStringExact(objRoles.CurrentObject.RoleID)
        End If
        blnReloading = False
    End Sub

    Private Sub frmRole_Load(sender As Object, e As EventArgs) Handles Me.Load
        objRoles = New CRoles
    End Sub

    Private Sub frmRole_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        blnClearing = True
        ClearScreenControls(Me)
        blnReloading = True
        LoadRoles()
        grpEdit.Enabled = False
        blnClearing = False
        blnReloading = False
    End Sub

    Private Sub lstRoles_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstRoles.SelectedIndexChanged
        If blnClearing Then
            Exit Sub
        End If
        If blnReloading Then
            tslStatus.Text = ""
            Exit Sub
        End If
        If lstRoles.SelectedIndex = -1 Then 'nothing to do
            Exit Sub
        End If
        chkNew.Checked = False
        LoadSelectedRecord()
        grpEdit.Enabled = True
    End Sub

    Private Sub LoadSelectedRecord()
        Try
            objRoles.GetRoleByRoleID(lstRoles.SelectedItem.ToString)
            With objRoles.CurrentObject
                txtRoleID.Text = .RoleID
                txtDesc.Text = .RoleDescription
            End With
        Catch ex As Exception
            MessageBox.Show("Error loading Role record: " & ex.ToString, "Program Error", MessageBoxButtons.OK, MessageBoxIcon.Error)  'ex.ToString for debug purposes
        End Try
    End Sub

    Private Sub chkNew_CheckedChanged(sender As Object, e As EventArgs) Handles chkNew.CheckedChanged
        If blnClearing Then
            Exit Sub
        End If
        If chkNew.Checked Then
            tslStatus.Text = ""
            txtRoleID.Clear()
            txtDesc.Clear()
            lstRoles.SelectedIndex = -1
            grpRoles.Enabled = False
            grpEdit.Enabled = True
            objRoles.CreateNewRole()
            txtRoleID.Focus()
        Else
            grpRoles.Enabled = True
            grpEdit.Enabled = False
            objRoles.CurrentObject.IsNewRole = False
        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        blnClearing = True
        tslStatus.Text = ""
        chkNew.Checked = False
        errP.Clear()
        If lstRoles.SelectedIndex <> -1 Then
            LoadSelectedRecord() 'reload what was not selected in case user messed up the form
        Else 'disable edit area - nothing was currently selected
            grpEdit.Enabled = False
        End If
        blnClearing = False
        objRoles.CurrentObject.IsNewRole = False
        grpRoles.Enabled = True
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim intResult As Integer
        Dim blnErrors As Boolean
        tslStatus.Text = ""
        '-----add your validation code here----'
        If Not ValidateTextBoxLength(txtRoleID, errP) Then
            blnErrors = True
        End If
        If Not ValidateTextBoxLength(txtDesc, errP) Then
            blnErrors = True
        End If
        If blnErrors = True Then
            Exit Sub
        End If
        With objRoles.CurrentObject
            .RoleID = txtRoleID.Text
            .RoleDescription = txtDesc.Text
        End With
        Try
            Me.Cursor = Cursors.WaitCursor
            intResult = objRoles.Save
            If intResult = 1 Then
                tslStatus.Text = "Role record saved"
            End If
            If intResult = -1 Then 'roleID was not unique when adding a new record
                MessageBox.Show("RoleID must be unique. Unable to save Role record", "Database error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                tslStatus.Text = "Error"
            End If
        Catch ex As Exception
            MessageBox.Show("Unable to save Role record: " & ex.ToString, "Database error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            tslStatus.Text = "Error"
        End Try
        Me.Cursor = Cursors.Default
        blnReloading = True
        LoadRoles() 'load roles so that a newly saved record will appear in the list
        chkNew.Checked = False
        grpRoles.Enabled = True 'in case it was disabled for a new record
    End Sub
End Class
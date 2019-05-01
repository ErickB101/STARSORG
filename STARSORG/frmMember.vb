Imports System.Data.SqlClient
Imports System.IO
Public Class frmMember
    Private objMembers As CMembers
    Private blnClearing As Boolean
    Private blnReloading As Boolean
    Private strFileName As String
#Region "Toolbar Stuff"
    Private Sub tsbProxy_MouseEnter(sender As Object, e As EventArgs) Handles tsbCourse.MouseEnter, tsbEvent.MouseEnter, tsbHelp.MouseEnter, tsbHome.MouseEnter, tsbLogout.MouseEnter, tsbMember.MouseEnter, tsbRole.MouseEnter, tsbRSVP.MouseEnter, tsbSemester.MouseEnter, tsbTutor.MouseEnter, tsbAdmin.MouseEnter, tsbMemberRoles.MouseEnter
        'we need to do this only because we are not putting the images in our Image property of the toolbar buttons
        Dim tsbProxy As ToolStripButton
        tsbProxy = DirectCast(sender, ToolStripButton)
        tsbProxy.DisplayStyle = ToolStripItemDisplayStyle.Text
    End Sub
    Private Sub tsbProxy_MouseLeave(sender As Object, e As EventArgs) Handles tsbCourse.MouseLeave, tsbEvent.MouseLeave, tsbHelp.MouseLeave, tsbHome.MouseLeave, tsbLogout.MouseLeave, tsbMember.MouseLeave, tsbRole.MouseLeave, tsbRSVP.MouseLeave, tsbSemester.MouseLeave, tsbTutor.MouseLeave, tsbAdmin.MouseLeave, tsbMemberRoles.MouseLeave
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

    Private Sub tsbLogout_Click(sender As Object, e As EventArgs) Handles tsbLogout.Click
        intNextAction = ACTION_LOGOUT
        Me.Hide()
    End Sub

    Private Sub tsbMember_Click(sender As Object, e As EventArgs) Handles tsbMember.Click
        'nothing to do, already in frmRole
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
        intNextAction = ACTION_ADMIN
        Me.Hide()
    End Sub

    Private Sub tsbMemberRole_Click(sender As Object, e As EventArgs) Handles tsbMemberRoles.Click
        intNextAction = ACTION_MEMBERROLES
        Me.Hide()
    End Sub
#End Region

#Region "Text Boxes"
    Private Sub txtBoxes_GotFocus(sender As Object, e As EventArgs) Handles txtEmail.GotFocus, txtFName.GotFocus, txtLName.GotFocus, txtPID.GotFocus, txtMI.GotFocus
        Dim txtBox As TextBox
        txtBox = DirectCast(sender, TextBox)
        txtBox.SelectAll()
    End Sub
    Private Sub txtBoxes_LostFocus(sender As Object, e As EventArgs) Handles txtEmail.LostFocus, txtFName.LostFocus, txtLName.LostFocus, txtPID.LostFocus, txtMI.LostFocus
        Dim txtBox As TextBox
        txtBox = DirectCast(sender, TextBox)
        txtBox.DeselectAll()
    End Sub
#End Region
    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Dim blnErrors As Boolean
        Dim objDR As SqlDataReader
        lstMembers.Items.Clear()
        If txtMemberSearch.TextLength = 0 Then 'missing search value
            errP.SetError(txtMemberSearch, "You must make a selection here")
            blnErrors = True
        End If
        If blnErrors Then
            Exit Sub
        End If
        Try
            objDR = objMembers.GetMemberByLastName(txtMemberSearch.Text)
            Do While objDR.Read
                lstMembers.Items.Add(objDR.Item("LName") & ", " & objDR.Item("FName") & " #" & objDR.Item("PID"))
            Loop
            objDR.Close()
        Catch ex As Exception
            'could have CDB throw the error and trap it here
        End Try
        If objMembers.CurrentObject.PID <> "" Then
            lstMembers.SelectedIndex = lstMembers.FindStringExact(objMembers.CurrentObject.PID)
        End If
    End Sub
    Private Sub frmMember_Load(sender As Object, e As EventArgs) Handles Me.Load
        objMembers = New CMembers
    End Sub

    Private Sub frmMember_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        blnClearing = True
        ClearScreenControls(Me)
        blnReloading = True
        grpMemberDetails.Enabled = False
        blnClearing = False
        blnReloading = False
    End Sub

    Private Sub lstMembers_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstMembers.SelectedIndexChanged
        If blnClearing Then
            Exit Sub
        End If
        'If blnReloading Then
        '    tslStatus.Text = ""
        '    Exit Sub
        'End If
        If lstMembers.SelectedIndex = -1 Then 'nothing to do
            Exit Sub
        End If
        chkNewMember.Checked = False
        LoadSelectedMember()
        grpMemberDetails.Enabled = True
    End Sub
    Private Sub LoadSelectedMember()
        Try
            objMembers.GetMemberByMemberID(lstMembers.SelectedItem.ToString.Substring(lstMembers.SelectedItem.ToString.IndexOf("#") + 1))
            With objMembers.CurrentObject
                txtPID.Text = .PID
                txtFName.Text = .FName
                txtLName.Text = .LName
                txtMI.Text = .MI
                txtEmail.Text = .Email
                mskPhone.Text = .Phone
                picPhotoPath.Load(.PhotoPath)
                txtSemesterID.Text = .SemesterID
            End With
        Catch ex As Exception
            MessageBox.Show("Error loading Member records: " & ex.ToString, "Program error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub chkNewMember_CheckedChanged(sender As Object, e As EventArgs) Handles chkNewMember.CheckedChanged
        If blnClearing Then 'nothing to do
            Exit Sub
        End If
        If chkNewMember.Checked Then
            tslStatus.Text = ""
            txtPID.Clear()
            txtFName.Clear()
            txtLName.Clear()
            txtMI.Clear()
            txtEmail.Clear()
            mskPhone.Clear()
            txtSemesterID.Clear()
            picPhotoPath.Image = Nothing
            lstMembers.SelectedIndex = -1
            grpMembersList.Enabled = False
            grpSearchMember.Enabled = False
            grpMemberDetails.Enabled = True
            objMembers.CreateNewMember()
            txtPID.Focus()
        Else 'is not checked
            grpMembersList.Enabled = True
            grpSearchMember.Enabled = True
            grpMemberDetails.Enabled = False
            objMembers.CurrentObject.isNewMember = False
        End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim intResult As Integer
        Dim blnErrors As Boolean
        tslStatus.Text = ""
        '----add validation code here----'
        If Not ValidateTextBoxNumeric(txtPID, errP) Then
            blnErrors = True
        End If
        'If Not ValidateTextBoxLength(txtPID, errP) Then
        '    blnErrors = True
        'End If
        If Not ValidateTextBoxLength(txtFName, errP) Then
            blnErrors = True
        End If
        If Not ValidateTextBoxLength(txtLName, errP) Then
            blnErrors = True
        End If
        If Not ValidateTextBoxLength(txtEmail, errP) Then
            blnErrors = True
        End If
        If Not ValidateMaskedtextBoxLength(mskPhone, errP) Then
            blnErrors = True
        End If
        If Not ValidateTextBoxLength(txtSemesterID, errP) Then
            blnErrors = True
        End If
        If blnErrors Then
            Exit Sub
        End If
        With objMembers.CurrentObject
            .PID = txtPID.Text
            .FName = txtFName.Text
            .LName = txtLName.Text
            .MI = txtMI.Text
            .Email = txtEmail.Text
            .Phone = mskPhone.Text
            .PhotoPath = picPhotoPath.ImageLocation.ToString
            .SemesterID = txtSemesterID.Text
        End With
        Try
            Me.Cursor = Cursors.WaitCursor
            intResult = objMembers.Save
            If intResult = 1 Then
                tslStatus.Text = "Member record saved"
            End If
            If intResult = -1 Then 'PID is not unique when adding a new record
                MessageBox.Show("PID must be unique. Unable to save Member record", "Database error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                tslStatus.Text = "Error"
            End If
        Catch ex As Exception
            MessageBox.Show("Unable to save Member record", "Database error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            tslStatus.Text = "Error"
        End Try
        Me.Cursor = Cursors.Default
        blnReloading = True
        chkNewMember.Checked = False
        grpMembersList.Enabled = True
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        blnClearing = True
        tslStatus.Text = ""
        chkNewMember.Checked = False
        errP.Clear()
        If lstMembers.SelectedIndex <> -1 Then
            LoadSelectedMember() 'reload what was selected in case user had messed up the form
        Else 'disable member details are, nothing was currently selected
            grpMemberDetails.Enabled = False
        End If
        blnClearing = False
        objMembers.CurrentObject.isNewMember = False
        grpMemberDetails.Enabled = False
        grpMembersList.Enabled = True
        grpSearchMember.Enabled = True
    End Sub
    Private Sub btnBrowsePhoto_Click(sender As Object, e As EventArgs) Handles btnBrowsePhoto.Click
        Dim intResult As Integer
        ofdOpen.InitialDirectory = Application.StartupPath
        ofdOpen.Filter = "All File (*.*)|*.*|Image Files (*.gif;*.jpg;*.jpeg;*.bmp;*.wmf;*.png)|*.gif;*.jpg;*.jpeg;*.bmp;*.wmf;*.png"
        ofdOpen.FilterIndex = 1
        intResult = ofdOpen.ShowDialog
        If intResult = DialogResult.Cancel Then 'user cancelled the browse
            Exit Sub
        End If
        strFileName = ofdOpen.FileName
        Try
            picPhotoPath.Load(strFileName)
        Catch exNotFound As FileNotFoundException
            MessageBox.Show("File was not found" & exNotFound.ToString, "Program error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch exIOError As IOException
            MessageBox.Show("IO error" & exIOError.ToString, "Program error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch exOther As Exception 'anything else that might go wrong
            MessageBox.Show("Unexpected error" & exOther.ToString, "Program error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnReport_Click(sender As Object, e As EventArgs) Handles btnReport.Click
        Dim MemberReport As New frmMemberReport
        Me.Cursor = Cursors.WaitCursor
        MemberReport.Display()
        Me.Cursor = Cursors.Default
    End Sub


End Class
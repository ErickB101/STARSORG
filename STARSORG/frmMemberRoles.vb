Imports System.Data.SqlClient
Public Class frmMemberRoles

    Private objMembersRoles As CMembersRoles
    Private blnClearing As Boolean
    Private blnReloading As Boolean


#Region "toolbar"
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
    Private Sub tsbMemberRole_Click(sender As Object, e As EventArgs) Handles tsbMemberRoles.Click
        'intNextAction = ACTION_MEMBERROLE
        'Me.Hide()
        ' COMMENTED OUT SO THE USER CANNOT GO TO THE SAME FORM THAT THEY ARE ALREADY IN THIS FORCES THEM TO STAY IN THERE FORM OR SWITCH FORMS
    End Sub
    Private Sub tsbAdmin_Click(sender As Object, e As EventArgs) Handles tsbAdmin.Click
        intNextAction = ACTION_ADMIN
        Me.Hide()
    End Sub
#End Region

#Region "textboxes"
    Private Sub txtAll_GotFocus(Sender As Object, e As EventArgs)
        Dim txtBox As TextBox
        txtBox = DirectCast(Sender, TextBox)
        txtBox.SelectAll()

    End Sub
    Private Sub txtAll_lostFocus(Sender As Object, e As EventArgs)
        Dim txtBox As TextBox
        txtBox = DirectCast(Sender, TextBox)
        txtBox.SelectAll()
    End Sub
#End Region
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


    Private Sub frmMember_Load(sender As Object, e As EventArgs) Handles Me.Load
        objMembersRoles = New CMembersRoles
        cboSemesters.Items.Clear()
        cboRoles.Items.Clear()
        LoadcboSemester()
        LoadMembers()       'Have to load Roles here cause need lstMembers select index parameters to be choosen then passed (otherwise error)
        LoadcboRoles()

    End Sub
    Private Sub LoadMembers()
        Dim objDR As SqlDataReader
        lstMemberRoles.Items.Clear()
        Try
            objDR = objMembersRoles.sp_getAllMembersForAdmin()
            Do While objDR.Read
                lstMemberRoles.Items.Add(objDR.Item("FName") & " " & objDR.Item("LName") & "," & objDR.Item("PID"))
            Loop
            objDR.Close()
        Catch ex As Exception
            'already handled in CDB(CBD)
            Throw ex 'should throw it if not handling here
        End Try
        If objMembersRoles.currentObject.PID <> "" Then
            lstMemberRoles.SelectedIndex = lstMemberRoles.FindStringExact(objMembersRoles.currentObject.PID)
        End If
        blnReloading = False

    End Sub
    Private Sub LoadRoles()
        lstRoles.Items.Clear()
        Try
            Dim objDR As SqlDataReader
            Dim strPID = Trim(lstMemberRoles.SelectedItem.ToString.Split(",")(1))
            Dim strSemesterID = cboSemesters.SelectedItem
            objDR = objMembersRoles.sp_getRoleIDByPID(strPID, strSemesterID)
            Do While objDR.Read

                lstRoles.Items.Add(objDR.Item("RoleID") & " | " & objDR.Item("semesterID"))
            Loop
            objDR.Close()
        Catch ex As Exception
            MessageBox.Show("Error loading event record" & ex.ToString, "Program Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        blnReloading = False

    End Sub
    Private Sub LoadcboRoles()
        Dim objDR As SqlDataReader
        Try
            objDR = objMembersRoles.sp_GetAllRoles()
            Do While objDR.Read
                cboRoles.Items.Add(objDR.Item("RoleID"))
            Loop
            objDR.Close()
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    Private Sub LoadcboSemester()

        Dim objDR As SqlDataReader
        Try
            objDR = objMembersRoles.sp_GetAllSemesters()
            Do While objDR.Read
                cboSemesters.Items.Add(objDR.Item("SemesterID"))
            Loop
            objDR.Close()
        Catch ex As Exception
            Throw ex
        End Try

    End Sub


    Private Sub lstMemberRoles_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstMemberRoles.SelectedIndexChanged
        lstRoles.Items.Clear()
        If blnClearing Then
            Exit Sub
        End If
        If blnReloading Then
            tslStatus.Text = ""
            Exit Sub
        End If
        If lstMemberRoles.SelectedIndex = -1 Then
            Exit Sub
        End If

        LoadRoles()
        grpEdit.Enabled = True
        grpRoles.Enabled = True

    End Sub
    Private Sub cboSemesters_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSemesters.SelectedIndexChanged
        grpMembers.Enabled = True
        grpRoles.Enabled = True

        If cboSemesters.SelectedIndex <> -1 Then
            ' there is semester selected
            If lstMemberRoles.SelectedIndex <> -1 Then
                ' there is a member selected
                LoadRoles()
            End If
        End If
    End Sub


    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click

        Dim intResult As Integer
        Dim blnErrors As Boolean
        tslStatus.Text = ""

        If Not ValidateCombo(cboRoles, errP) Then
            blnErrors = True
        End If
        If blnErrors Then
            Exit Sub
        End If
        If cboRoles.SelectedItem = "" Then
            MessageBox.Show("Please Select a Role", "Program Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        If cboRoles.SelectedItem.ToString = "Guest" Then
            MessageBox.Show("You Can Not Pick Guest as a Role", "Program Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        With objMembersRoles.currentObject
            .RoleId = cboRoles.SelectedItem.ToString
            .PID = Trim(lstMemberRoles.SelectedItem.ToString.Split(",")(1))
            .SemesterID = cboSemesters.SelectedItem.ToString
        End With

        Try
            Me.Cursor = Cursors.WaitCursor
            intResult = objMembersRoles.save
            If intResult = 1 Then  'Update Confirmation for StatusBox
                tslStatus.Text = "Member Role Record added"
            End If
            If intResult = -1 Then
                MessageBox.Show("OOPS! You've Already Picked This Role For Selected Semester", "database error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MessageBox.Show("unable to save Member Role record" & ex.ToString, "database error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            tslStatus.Text = "Error"
        End Try

        Me.Cursor = Cursors.Default
        blnReloading = True
        grpMembers.Enabled = True
        LoadRoles()

    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        'ERROR CHECKING
        If cboRoles.SelectedIndex <> -1 Then
            MessageBox.Show("You did not select a Role to delete", "database error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        ElseIf lstMemberRoles.SelectedIndex <> -1 Then
            MessageBox.Show("You Need to Select A Member to Delete", "database error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        ElseIf cboSemesters.SelectedIndex <> -1 Then
            MessageBox.Show("You did not select a Semesterto delete", "database error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If


        Dim RoleId = cboRoles.SelectedItem.ToString
        Dim PID = Trim(lstMemberRoles.SelectedItem.ToString.Split(",")(1))
        Dim semesterID = cboSemesters.SelectedItem.ToString

        objMembersRoles.sp_DeleteRecord(PID, RoleId, semesterID)

        LoadRoles()
    End Sub


    Private Sub btnReport_Click(sender As Object, e As EventArgs) Handles btnReport.Click

        Dim memberRoleReport As New frmMemberRoleReport
        Me.Cursor = Cursors.WaitCursor
        memberRoleReport.Display()
        Me.Cursor = Cursors.Default

    End Sub


End Class
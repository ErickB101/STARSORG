Imports System.Data.SqlClient

Public Class frmSemester

    Private objSemesters As CSemesters
    Private blnClearing As Boolean
    Private blnReloading As Boolean

#Region "ToolBar Stuff"
    Private Sub tsbMember_Click(sender As Object, e As EventArgs) Handles tsbMember.Click
        intNextAction = ACTION_MEMBER
        Me.Hide()
    End Sub

    Private Sub tsbLogOut_Click(sender As Object, e As EventArgs) Handles tsbLogOut.Click
        intNextAction = ACTION_LOGOUT
        Me.Hide()
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

    Private Sub tsbRole_Click(sender As Object, e As EventArgs) Handles tsbRole.Click
        intNextAction = ACTION_ROLE
        Me.Hide()
    End Sub


    Private Sub tsbHome_Click(sender As Object, e As EventArgs) Handles tsbHome.Click
        intNextAction = ACTION_HOME
        Me.Hide()
    End Sub

    Private Sub tsbSemester_Click(sender As Object, e As EventArgs) Handles tsbSemester.Click
        'Do nothing since you are here
    End Sub

    Private Sub tsbRSVP_Click(sender As Object, e As EventArgs) Handles tsbRSVP.Click
        intNextAction = ACTION_RSVP
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

#End Region

#Region "TextBoxes"

    Private Sub txtBoxes_GotFocus(sender As Object, e As EventArgs) Handles txtSemesterID.GotFocus, txtDesc.GotFocus
        Dim txtBox As TextBox
        txtBox = DirectCast(sender, TextBox)
        txtBox.SelectAll()


    End Sub

    Private Sub txtBoxes_LostFocus(sender As Object, e As EventArgs) Handles txtSemesterID.LostFocus, txtDesc.LostFocus

        Dim txtBox As TextBox
        txtBox = DirectCast(sender, TextBox)
        txtBox.DeselectAll()

    End Sub

#End Region

    Private Sub LoadSemesters()
        Dim objDR As SqlDataReader

        lstSemesters.Items.Clear()

        Try
            objDR = objSemesters.GetAllSemesters()
            Do While objDR.Read
                lstSemesters.Items.Add(objDR.Item("SemesterID"))
            Loop

            objDR.Close()

        Catch ex As Exception

            'Already handled in CDB

        End Try

        If objSemesters.CurrentObject.SemesterID <> "" Then

            lstSemesters.SelectedIndex = lstSemesters.FindStringExact(objSemesters.CurrentObject.SemesterID)

        End If

        blnReloading = False

    End Sub

    Private Sub frmSemester_Load(sender As Object, e As EventArgs) Handles Me.Load
        objSemesters = New CSemesters

    End Sub


    Private Sub frmSemester_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        ClearScreenControls(Me)
        LoadSemesters()
        grpEdit.Enabled = False

        blnReloading = True
        blnClearing = True
        blnReloading = False
        blnClearing = False
    End Sub

    Private Sub lstSemesters_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstSemesters.SelectedIndexChanged

        If blnClearing Then

            Exit Sub

        End If

        If blnReloading Then
            Exit Sub
        End If

        If lstSemesters.SelectedIndex = -1 Then 'nothing to do
            Exit Sub
        End If

        chkNew.Checked = False
        LoadSelectedRecord()
        grpEdit.Enabled = True


    End Sub


    Private Sub LoadSelectedRecord()

        Try

            objSemesters.GetSemesterBySemesterID(lstSemesters.SelectedItem.ToString)
            With objSemesters.CurrentObject
                txtSemesterID.Text = .SemesterID
                txtDesc.Text = .SemesterDescription

            End With

        Catch ex As Exception

            MessageBox.Show("Error loading Semester Record: " & ex.ToString, "Program Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub chkNew_CheckedChanged(sender As Object, e As EventArgs) Handles chkNew.CheckedChanged

        If blnClearing Then

            Exit Sub

        End If

        If chkNew.Checked Then

            tslStatus.Text = ""
            txtSemesterID.Clear()
            txtDesc.Clear()
            lstSemesters.SelectedIndex = -1
            grpSemesters.Enabled = False
            grpEdit.Enabled = True
            txtSemesterID.Focus()
            objSemesters.CreateNewSemester()
        Else

            grpSemesters.Enabled = True
            grpEdit.Enabled = False
            objSemesters.CurrentObject.IsNewSemester = False

        End If
    End Sub



    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        blnClearing = True
        tslStatus.Text = ""
        chkNew.Checked = False
        errP.Clear()

        If lstSemesters.SelectedIndex <> -1 Then
            LoadSelectedRecord() 'reload what was selected in case user had messed up the form

        Else ' Disable edit area nothing was currently selected

            grpEdit.Enabled = False

        End If

        blnClearing = False
        objSemesters.CurrentObject.IsNewSemester = False

        grpSemesters.Enabled = True


    End Sub







    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click

        Dim intResult As Integer
        Dim blnErrors As Boolean

        tslStatus.Text = ""

        '----------- add your validations code here --------------
        'modErrHandler Check that they are adding the correct data to the data base

        If Not ValidateTextBoxLength(txtSemesterID, errP) Then
            blnErrors = True
        End If

        If Not ValidateTextBoxLength(txtDesc, errP) Then

            blnErrors = True

        End If

        If blnErrors Then
            Exit Sub

        End If

        With objSemesters.CurrentObject
            .SemesterID = txtSemesterID.Text
            .SemesterDescription = txtDesc.Text
        End With

        Try
            Me.Cursor = Cursors.WaitCursor
            intResult = objSemesters.Save

            If intResult = 1 Then
                tslStatus.Text = "Semester Record Saved"
            End If

            If intResult = -1 Then
                MessageBox.Show("Semester ID must be unique. Unable to save Semester Record", "DataBase Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                tslStatus.Text = "Error"

            End If


        Catch ex As Exception
            MessageBox.Show("Unable to save Semester Record" & ex.ToString, "DataBase Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            tslStatus.Text = "Error"

        End Try

        Me.Cursor = Cursors.Default
        blnReloading = True
        LoadSemesters() ' Reload so that a newly saved record will appear in the list

        grpSemesters.Enabled = True 'In case it was disabled for a new record




    End Sub

    Private Sub btnReport_Click(sender As Object, e As EventArgs) Handles btnReport.Click
        Dim SemesterReport As New frmSemesterReport
        Me.Cursor = Cursors.WaitCursor
        SemesterReport.Display()
        Me.Cursor = Cursors.Default
    End Sub

End Class
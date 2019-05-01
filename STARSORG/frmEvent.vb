'this is code for the Event form displayed to the user

Imports System.Data.SqlClient
Public Class frmEvent
    Private objEvents As CEvents
    Private blnClearing As Boolean
    Private blnReloading As Boolean

    'created focus functions that is determined by whether the textbox is selected or not
#Region "Textboxes"
    Private Sub txtBoxes_GotFocus(sender As Object, e As EventArgs) Handles txtEventID.GotFocus, txtDesc.GotFocus, txtLocation.GotFocus
        Dim txtBox As TextBox
        txtBox = DirectCast(sender, TextBox)
        txtBox.SelectAll()
    End Sub

    Private Sub txtBoxes_LostFocus(sender As Object, e As EventArgs) Handles txtEventID.LostFocus, txtDesc.LostFocus, txtLocation.LostFocus
        Dim txtBox As TextBox
        txtBox = DirectCast(sender, TextBox)
        txtBox.DeselectAll()
    End Sub
#End Region

'this region controls the tool bar on the STARSORG main page

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

    Private Sub tsbMember_Click(sender As Object, e As EventArgs) Handles tsbMember.Click
        intNextAction = ACTION_MEMBER
        Me.Hide()
    End Sub

    Private Sub tsbLogOut_Click(sender As Object, e As EventArgs) Handles tsbLogOut.Click
        intNextAction = ACTION_LOGOUT
        Me.Hide()
    End Sub

    Private Sub tsbRole_Click(sender As Object, e As EventArgs) Handles tsbRole.Click
        intNextAction = ACTION_ROLE
        Me.Hide()
    End Sub

    Private Sub tsbCourse_Click(sender As Object, e As EventArgs) Handles tsbCourse.Click
        intNextAction = ACTION_COURSE
        Me.Hide()
    End Sub

    Private Sub tsbHelp_Click(sender As Object, e As EventArgs) Handles tsbHelp.Click
        intNextAction = ACTION_HELP
        Me.Hide()
    End Sub

    Private Sub tsbEvent_Click(sender As Object, e As EventArgs) Handles tsbEvent.Click
        'nothing to do here
    End Sub

    Private Sub tsbHome_Click(sender As Object, e As EventArgs) Handles tsbHome.Click
        intNextAction = ACTION_HOME
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

    Private Sub tsbRSVP_Click(sender As Object, e As EventArgs) Handles tsbRSVP.Click
        intNextAction = ACTION_RSVP
        Me.Hide()
    End Sub

    Private Sub tsbMemberRole_Click(sender As Object, e As EventArgs) Handles tsbMemberRoles.Click
        intNextAction = ACTION_MEMBERROLES
        Me.Hide()
    End Sub
#End Region

'this loads all the events on a list box so the user can select it and view its contents

    Private Sub LoadEvents()
        Dim objReader As SqlDataReader
        lstEvents.Items.Clear()
        Try
            objReader = objEvents.GetAllEvents()
            Do While objReader.Read
                lstEvents.Items.Add(objReader.Item("EventID"))
            Loop
            objReader.Close()
        Catch ex As Exception
            'already handled in CDB
        End Try
        If objEvents.CurrentObject.EventID <> "" Then
            lstEvents.SelectedIndex = lstEvents.FindStringExact(objEvents.CurrentObject.EventID)
        End If
        blnReloading = False
    End Sub

'this loads all the Event Types on a combo box so the user can select it 

    Private Sub LoadTypes()
        Dim objReader As SqlDataReader
        cboTypeID.Items.Clear()
        Try
            objReader = myDB.GetDataReaderBySP("sp_getAllTypes", Nothing)
            Do While objReader.Read
                cboTypeID.Items.Add(objReader.Item("EventTypeID"))
            Loop
            objReader.Close()
        Catch ex As Exception
            'already handled in CDB
        End Try
    End Sub

'this loads all the Semesters on a combo box so the user can select it 

    Private Sub LoadSemesters()
        Dim objReader As SqlDataReader
        cboSemesterID.Items.Clear()
        Try
            objReader = myDB.GetDataReaderBySP("sp_getAllSemesters", Nothing)
            Do While objReader.Read
                cboSemesterID.Items.Add(objReader.Item("SemesterID"))
            Loop
            objReader.Close()
        Catch ex As Exception
            'already handled in CDB
        End Try
    End Sub

'these 2 subs basically loads every function that needs to be loaded when the page starts up
    Private Sub frmEvent_Load(sender As Object, e As EventArgs) Handles Me.Load
        objEvents = New CEvents
        LoadTypes()
        LoadSemesters()
    End Sub

    Private Sub frmEvent_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        blnClearing = True
        ClearScreenControls(Me)
        blnReloading = True
        LoadEvents()
        grpEdit.Enabled = False
        blnClearing = False
        blnReloading = False
    End Sub

'these 2 subs have to do with loading the Event records information from the database

    Private Sub lstEvents_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstEvents.SelectedIndexChanged
        If blnClearing Then
            Exit Sub
        End If

        If blnReloading Then
            Exit Sub
        End If

        If lstEvents.SelectedIndex = -1 Then 'nothing to do
            Exit Sub
        End If
        chkNew.Checked = False
        LoadSelectedRecord()
        grpEdit.Enabled = True
        If (dtpStartDate.Value.Date < DateTime.Now.Date) Then
            txtEventID.Enabled = False
            cboTypeID.Enabled = True
            txtDesc.Enabled = False
            cboSemesterID.Enabled = False
            dtpStartDate.Enabled = False
            dtpEndDate.Enabled = False
            txtLocation.Enabled = False
        Else
            txtEventID.Enabled = True
            txtDesc.Enabled = True
            cboSemesterID.Enabled = True
            dtpStartDate.Enabled = True
            dtpEndDate.Enabled = True
            txtLocation.Enabled = True
        End If
    End Sub

    Private Sub LoadSelectedRecord()
        Try
            objEvents.GetEventByEventID(lstEvents.SelectedItem.ToString)
            With objEvents.CurrentObject
                txtEventID.Text = .EventID
                txtDesc.Text = .EventDescription
                cboTypeID.Text = .TypeID
                cboSemesterID.Text = .SemesterID
                dtpStartDate.Value = .StartDate
                dtpEndDate.Value = .EndDate
                txtLocation.Text = .Location
            End With


        Catch ex As Exception
            MessageBox.Show("Error loading event record" & ex.ToString, "Program Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

'this controls what happens when the user checks the new event box

    Private Sub chkNew_CheckedChanged(sender As Object, e As EventArgs) Handles chkNew.CheckedChanged
        If blnClearing Then
            Exit Sub
        End If

        If chkNew.Checked Then


            tslStatus.Text = ""
            txtEventID.Clear()
            txtDesc.Clear()
            LoadTypes()
            LoadSemesters()
            dtpStartDate.ResetText()
            dtpEndDate.ResetText()
            txtLocation.Clear()
            lstEvents.SelectedIndex = -1
            grpEvents.Enabled = False
            grpEdit.Enabled = True
            objEvents.CreateNewEvent()
            txtEventID.Focus()



        Else
            grpEvents.Enabled = True
            grpEdit.Enabled = False
            objEvents.CurrentObject.IsNewEvent = False
        End If
    End Sub

'this sub controls what happens when the user clicks the cancel button

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        blnClearing = True
        tslStatus.Text = ""
        chkNew.Checked = False
        errP.Clear()

        If lstEvents.SelectedIndex <> -1 Then
            LoadSelectedRecord() 'reload what wa selected in case user had messed up the form
        Else 'disable edit area = nothing was currently selected
            grpEdit.Enabled = False
        End If
        blnClearing = False
        objEvents.CurrentObject.IsNewEvent = False
        grpEvents.Enabled = True

    End Sub

'this sub controls what happens when the user clicks the save button

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click

        Dim intResult As Integer
        Dim blnErrors As Boolean

        tslStatus.Text = ""

        '----------- add your validations code here --------------
        'modErrHandler Check that they are adding the correct data to the data base

        If Not ValidateTextBoxLength(txtEventID, errP) Then
            blnErrors = True
        End If

        If Not ValidateTextBoxLength(txtDesc, errP) Then

            blnErrors = True

        End If

        If Not ValidateTextBoxLength(txtLocation, errP) Then
            blnErrors = True
        End If


        If blnErrors Then
            Exit Sub

        End If

        With objEvents.CurrentObject
            .EventID = txtEventID.Text
            .EventDescription = txtDesc.Text
            .TypeID = cboTypeID.Text
            .SemesterID = cboSemesterID.Text
            .StartDate = dtpStartDate.Value
            .EndDate = dtpEndDate.Value
            .Location = txtLocation.Text
        End With

        Try
            Me.Cursor = Cursors.WaitCursor
            intResult = objEvents.Save

            If intResult = 1 Then
                tslStatus.Text = "Event Record Saved"
            End If

            If intResult = -1 Then
                MessageBox.Show("Event ID must be unique. Unable to save Role Record", "DataBase Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                tslStatus.Text = "Error"

            End If


        Catch ex As Exception
            MessageBox.Show("Unable to save Event Record" & ex.ToString, "DataBase Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            tslStatus.Text = "Error"

        End Try

        Me.Cursor = Cursors.Default
        blnReloading = True
        LoadEvents() ' Reload so that a newly saved record will appear in the list

        grpEvents.Enabled = True 'In case it was disabled for a new record

    End Sub

    Private Sub tsbTutor_Click_1(sender As Object, e As EventArgs) Handles tsbTutor.Click

    End Sub

    'Private Sub btnReport_Click(sender As Object, e As EventArgs) Handles btnReport.Click
    '    Dim RoleReport As New frmRoleReport
    '    Me.Cursor = Cursors.WaitCursor
    '    RoleReport.Display()
    '    Me.Cursor = Cursors.Default
    'End Sub

End Class

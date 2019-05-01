Imports System.Data.SqlClient
Public Class frmRSVP
    Private objRSVPs As CEvent_RSVPs
    Private blnClearing As Boolean
    Private blnReloading As Boolean


#Region "Toolbar stuff"
    Private Sub tsbMember_Click(sender As Object, e As EventArgs) Handles tsbMember.Click
        intNextAction = ACTION_MEMBER
        Me.Hide()
    End Sub

    Private Sub tsbLogOut_Click(sender As Object, e As EventArgs) Handles tsbRole.Click
        'nothing to do here = already on the role form
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
        intNextAction = ACTION_EVENT
        Me.Hide()
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
        'already here
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

    Private Sub LoadEvents()
        Dim objReader As SqlDataReader
        lstEvents.Items.Clear()
        Try
            objReader = objRSVPs.GetAllEvents
            Do While objReader.Read
                lstEvents.Items.Add(objReader.Item("EventID"))

            Loop
            objReader.Close()
        Catch ex As Exception
            'already handled in CDB
        End Try
        If objRSVPs.CurrentObject.EventID <> "" Then
            lstEvents.SelectedIndex = lstEvents.FindStringExact(objRSVPs.CurrentObject.EventID)
        End If
        blnReloading = False
    End Sub

    Private Sub frmRSVP_Load(sender As Object, e As EventArgs) Handles Me.Load
        objRSVPs = New CEvent_RSVPs
    End Sub

    Private Sub frmEvent_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        blnClearing = True
        ClearScreenControls(Me)
        blnReloading = True
        LoadEvents()
        blnClearing = False
        blnReloading = False
    End Sub

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
        LoadSelectedRecord()
    End Sub

    Private Sub LoadSelectedRecord()
        Try
            objRSVPs.GetEventByEventID(lstEvents.SelectedItem.ToString)
            With objRSVPs.CurrentObject
                lblStartDate.Text = .StartDate
                lblEndDate.Text = .EndDate
                lblLocation.Text = .Location
            End With
        Catch ex As Exception
            MessageBox.Show("Error loading event record" & ex.ToString, "Program Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        blnClearing = False
        tslStatus.Text = ""
        errP.Clear()
        txtFirstName.Clear()
        txtLastName.Clear()
        txtEmail.Clear()
        objRSVPs.CurrentObject.IsNewRSVP = False
    End Sub


    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnRSVP.Click
        Dim intResult As Integer
        Dim blnErrors As Boolean
        tslStatus.Text = ""
        'add your validation code here
        If Not ValidateTextBoxLength(txtFirstName, errP) Then
            blnErrors = True
        End If

        If Not ValidateTextBoxLength(txtLastName, errP) Then
            blnErrors = True
        End If

        If Not ValidateTextBoxLength(txtEmail, errP) Then
            blnErrors = True
        End If



        If blnErrors Then
            Exit Sub
        End If

        If (CDate(lblStartDate.Text) <= DateTime.Now.Date) Then
            MessageBox.Show("This event has already passed", "Database error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            tslStatus.Text = "Error"
            Exit Sub
        End If



        With objRSVPs.CurrentObject
            .EventID = objRSVPs.CurrentObject.EventID
            .FirstName = txtFirstName.Text
            .LastName = txtLastName.Text
            .Email = txtEmail.Text
        End With

        Try
            Me.Cursor = Cursors.WaitCursor
            intResult = objRSVPs.Save
            If intResult = 1 Then
                tslStatus.Text = "RSVP record saved"
            End If

            If intResult = -1 Then 'role ID was not unique when adding a new record
                MessageBox.Show("Email must be unique. Unable to save event record", "Database error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                tslStatus.Text = "Error"
            End If
        Catch ex As Exception
            MessageBox.Show("Unable to save RSVP record" & ex.ToString, "Database error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            tslStatus.Text = "Error"
        End Try
        Me.Cursor = Cursors.Default
        blnReloading = True
        LoadEvents() 'reload so that a newly saved record will appear on the list
    End Sub

    Private Sub btnReport_Click(sender As Object, e As EventArgs) Handles btnReport.Click
        Dim RSVPReport As New frmRSVPReport
        Me.Cursor = Cursors.WaitCursor
        RSVPReport.Display()
        Me.Cursor = Cursors.Default
    End Sub

End Class
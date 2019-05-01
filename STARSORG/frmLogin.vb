Public Class frmLogin
    Private Main As frmMain
    Private objSecurities As CSecurities
    Private objAudits As CAudits
    Private blnValidLogin As Boolean

#Region "Constants"
    Public Const GUESTID = "0000001"
#End Region

#Region "Textboxes"
    Private Sub txtAll_GotFocus(sender As Object, e As EventArgs) Handles txtUserID.GotFocus, txtUserPass.GotFocus, txtNewPass.GotFocus
        Dim txtBox As TextBox
        txtBox = DirectCast(sender, TextBox)
        txtBox.SelectAll()
    End Sub

    Private Sub txtAll_LostFocus(sender As Object, e As EventArgs) Handles txtUserID.LostFocus, txtUserPass.LostFocus, txtNewPass.LostFocus
        Dim txtBox As TextBox
        txtBox = DirectCast(sender, TextBox)
        txtBox.DeselectAll()
    End Sub
#End Region

    Private Sub Clear()
        objSecurities = New CSecurities
        objAudits = New CAudits
        tslStatus.Text = ""
        errP.Clear()
        blnValidLogin = False
    End Sub

    Private Function ValidateCredentials() As Integer
        Clear()
        Dim intResult As Integer
        Dim blnResult As Boolean
        'simple errors, returns error code 1: required fields are empty
        If Not ValidateTextBoxLength(txtUserID, errP) Then 'missing user ID
            blnResult = True
        End If
        If Not ValidateTextBoxLength(txtUserPass, errP) Then 'missing password
            blnResult = True
        End If
        If blnResult Then
            tslStatus.Text = "Error: missing credentials"
            Return 1
        End If

        'text inputs valid
        Dim username = txtUserID.Text
        Dim pass = txtUserPass.Text
        Try
            intResult = objSecurities.CheckUsernameAndPass(username, pass)
            If intResult = 1 Then
                'success
                Return 0
            Else
                ' error
                Return 2
            End If
        Catch ex As Exception
            MessageBox.Show("Unable to process login attempt " & ex.ToString, "Database error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            tslStatus.Text = "Error: exception"
            Return 2
        End Try

        Return 0 'if no errors occured
    End Function

    Private Sub frmLogin_Load(sender As Object, e As EventArgs) Handles Me.Load
        Clear()
        Main = New frmMain
        txtUserID.Text = ""
        txtUserPass.Text = ""
        txtNewPass.Text = ""
        objSecurities.CurrentObject.IsNewSecRole = False
    End Sub

    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        Dim intValResult As Integer
        intValResult = ValidateCredentials()
        If intValResult = 1 Then
            Exit Sub
        End If
        If intValResult = 2 Then 'if errors occured
            tslStatus.Text = "Invalid credentials"
            MessageBox.Show(tslStatus.Text & " Unable to login, please try again.", "Login Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        'if no errors were encountered then proceed to login & save data to Audit
        tslStatus.Text = "Login successful"
        blnValidLogin = True

        objSecurities.GetSecByUserID(txtUserID.Text)

        'Access Granted
        ExitLogin()
    End Sub


    Private Sub btnGuestLogin_Click(sender As Object, e As EventArgs) Handles btnGuestLogin.Click
        'if this button is clicked - user has very limited restrictions

        Try
            objAudits.Save() 'save the guest login attempt
        Catch ex As Exception
            MessageBox.Show("Unable to process login attempt " & ex.ToString, "Database error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            tslStatus.Text = "Error: Exception"
            Exit Sub
        End Try

        'allow access to frmMain
        tslStatus.Text = "Guest login successful"
        blnValidLogin = True

        objSecurities.GetSecByPID(GUESTID)
        'Access Main Page
        ExitLogin()
    End Sub

    Private Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click
        Dim blnErrors As Boolean
        Dim intResult As Integer

        If Not ValidateTextBoxLength(txtNewPass, errP) Then 'missing password
            tslStatus.Text = "Error: missing new password."
            Exit Sub
        End If
        intResult = ValidateCredentials()
        If intResult = 1 Then 'if missing userID or password
            blnErrors = True
        End If
        If intResult = 2 Then 'if credentials invalid, error
            blnErrors = True
        End If
        If blnErrors Then
            tslStatus.Text = "Invalid Credentials"
            MessageBox.Show(tslStatus.Text & " Please try again.", "Login Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        'if no errors were encountered then proceed to save new password
        Try
            Me.Cursor = Cursors.WaitCursor
            objSecurities.GetSecByUserID(txtUserID.Text)
            With objSecurities.CurrentObject
                .Password = txtNewPass.Text
            End With
            intResult = objSecurities.Save
            If intResult = 1 Then
                tslStatus.Text = "New password saved successfully"
            End If
            If intResult = -1 Then
                MessageBox.Show("PID must be unique. Unable to save Security record", "Database error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                tslStatus.Text = "Error"
            End If
        Catch ex As Exception
            MessageBox.Show("Unable to save new password: " & ex.ToString, "Database error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            tslStatus.Text = "Error"
            Exit Sub
        End Try
        Me.Cursor = Cursors.Default

        btnCancel.PerformClick()  'return to login format
    End Sub


    Private Sub btnChangePass_Click(sender As Object, e As EventArgs) Handles btnChangePass.Click
        'toggle new password fields on
        lblNewPass.Visible = True
        txtNewPass.Visible = True
        btnSubmit.Visible = True
        btnCancel.Visible = True
        btnLogin.Visible = False
        btnChangePass.Visible = False
        txtNewPass.Text = ""
    End Sub

    Private Sub ExitLogin()

        Me.Hide()
        If blnValidLogin Then
            intNextAction = ACTION_HOME
        End If
    End Sub

    Public Function CurrentUser()
        Return objSecurities
    End Function

    Public Function IsValidLogin() As Boolean
        If blnValidLogin = True Then
            Return True
        End If
        Return False
    End Function

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        lblNewPass.Visible = False
        txtNewPass.Visible = False
        btnSubmit.Visible = False
        btnLogin.Visible = True
        btnChangePass.Visible = True
        txtUserPass.Text = ""
        txtNewPass.Text = ""
        txtUserID.Text = ""
        Clear()
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Main.EndProgram()
    End Sub
End Class
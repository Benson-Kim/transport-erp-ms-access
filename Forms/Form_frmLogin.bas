
Option Compare Database
Option Explicit

Private Sub cmdLogin_Click()
    On Error GoTo ErrHandler
        
    If Trim(Me.txtUsername & "") = "" Then
        Me.txtUsername.SetFocus
        Me.txtUsername.BorderColor = RGB(242, 80, 34)
        Me.lblStatus.Caption = "Username is required"
        Exit Sub
    End If
    Me.txtUsername.BorderColor = RGB(166, 166, 166)
    
    If Trim(Me.txtPassword & "") = "" Then
        Me.txtPassword.SetFocus
        Me.txtPassword.BorderColor = RGB(242, 80, 34)
        Me.lblStatus.Caption = "Password is required"
        Exit Sub
    End If
    Me.txtPassword.BorderColor = RGB(166, 166, 166)
    
    modAuthentication.AuthenticateUser Me.txtUsername, Me.txtPassword
    
    If modAuthentication.AuthenticateUser(Me.txtUsername, Me.txtPassword) Then
        DoCmd.Close acForm, Me.Name
        DoCmd.OpenForm "frmMainMenu"  ' Open main form on success
    End If
    
    Exit Sub
ErrHandler:
    modUtilities.LogError "frmLogin.cmdLogin", Err.number, Err.Description
    Me.lblStatus.Caption = "Login error: " & Err.Description
End Sub

Private Sub cmdResetPassword_Click()
    On Error GoTo ErrHandler
    If Not HasPermission("ResetPassword") Then
        MsgBox "Insufficient permissions to reset password.", vbExclamation
        Exit Sub
    End If
    
    Dim Username As String: Username = InputBox("Enter username to reset:")
    If Username = "" Then Exit Sub
    Dim NewPassword As String: NewPassword = InputBox("Enter new password:")
    If NewPassword = "" Then Exit Sub
    
    Dim UserID As Long: UserID = Nz(DLookup("UserID", "Users", "Username = '" & Replace(Username, "'", "''") & "'"), 0)
    If UserID = 0 Then
        MsgBox "User not found.", vbExclamation
        Exit Sub
    End If
    
    UpdateUser UserID, Username, DLookup("Role", "Users", "UserID = " & UserID), NewPassword, DLookup("CanCreateItems", "Users", "UserID = " & UserID), DLookup("CanEditItems", "Users", "UserID = " & UserID), DLookup("IsActive", "Users", "UserID = " & UserID)
    MsgBox "Password reset successfully.", vbInformation
    Exit Sub
ErrHandler:
    modUtilities.LogError "frmLogin_cmdResetPassword", Err.number, Err.Description
    MsgBox "Reset password error: " & Err.Description, vbCritical, modGlobals.APP_NAME & " - Reset Password"
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandler
    
    Me.txtUsername.BorderColor = RGB(166, 166, 166)
    Me.txtPassword.BorderColor = RGB(166, 166, 166)
    Me.lblStatus.Caption = ""
    Me.txtPassword.InputMask = "Password"  ' Hide password input
    
    Exit Sub
    
ErrHandler:
    modUtilities.LogError "frmLogin.Form_Load", Err.number, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrHandler
    Me.txtUsername.BorderColor = RGB(166, 166, 166)
    Me.txtPassword.BorderColor = RGB(166, 166, 166)
    Me.lblStatus.Caption = ""
    Exit Sub
ErrHandler:
     modUtilities.LogError "frmLogin.Form_Unload", Err.number, Err.Description
End Sub



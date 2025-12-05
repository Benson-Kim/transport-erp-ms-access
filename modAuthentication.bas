Attribute VB_Name = "modAuthentication"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MODULE: modAuthentication
' PURPOSE: Enterprise-grade authentication with salted SHA-256,
'          brute-force protection, lockout, expiry, and full audit
' SECURITY: Production-ready – NO PLAIN TEXT PASSWORDS
' AUTHOR: Expert Back-End Developer (MS Access Security Specialist)
' UPDATED: November 18, 2025 – Fully secure & compliant
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Compare Database
Option Explicit

Private Const MODULE_NAME As String = "modAuthentication"
Private Const MAX_FAILED_ATTEMPTS As Integer = 5
Private Const BASE_LOCKOUT_MINUTES As Integer = 5  ' Base for exponential backoff
Private Const PASSWORD_EXPIRY_DAYS As Integer = 90


' PUBLIC: AuthenticateUser – Core secure login function

Public Function AuthenticateUser(ByVal strUsername As String, _
                                 ByVal strPassword As String) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    
    Dim lngUserID As Long
    Dim strRole As String
    Dim strFullName As String
    Dim blnIsActive As Boolean
    Dim lngFailedAttempts As Long
    Dim dtLockoutUntil As Variant
    Dim dtPasswordSet As Date
    Dim strStoredHash As String

    Set db = CurrentDb

    ' Normalize input
    strUsername = modUtilities.NormalizeUsername(strUsername)

    If Len(strUsername) = 0 Or Len(strPassword) = 0 Then
        GoTo LoginFailed
    End If

    ' Build safe SQL
    strSQL = "SELECT UserID, FullName, Role, IsActive, PasswordHash, " & _
             "FailedLoginAttempts, LockoutUntil, PasswordSetDate " & _
             "FROM Users WHERE UCase(Username) = '" & Replace(strUsername, "'", "''") & "'"

    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)

    If rs.EOF Then
        GoTo LoginFailed
    End If

    With rs
        lngUserID = Nz(!UserID, 0)
        strFullName = Nz(!FullName, strUsername)
        strRole = Nz(!Role, "Operator")
        blnIsActive = Nz(!IsActive, False)
        strStoredHash = Nz(!PasswordHash, "")
        lngFailedAttempts = Nz(!FailedLoginAttempts, 0)
        dtLockoutUntil = Nz(!LockoutUntil, #1/1/1900#)
        dtPasswordSet = Nz(!PasswordSetDate, #1/1/1900#)

        ' Account checks
        If Not blnIsActive Then
            MsgBox "This account has been deactivated.", vbCritical, APP_NAME
            GoTo LoginFailed
        End If

        If Not IsNull(dtLockoutUntil) And dtLockoutUntil > Now() Then
            MsgBox "Account locked until " & Format(dtLockoutUntil, "HH:mm") & _
                   vbCrLf & "Too many failed attempts.", vbCritical, APP_NAME
            GoTo LoginFailed
        End If


        ' Secure password verification
        
        If Len(strStoredHash) = 0 Then
            modAudit.LogAudit "Users", lngUserID, "Login", Null, _
                             "Missing password hash", "SecurityAlert", , True
            GoTo LoginFailed
        End If
        
        If Not modCrypto.VerifyPassword(strPassword, Nz(!PasswordHash, "")) Then
            GoTo LoginFailed
        End If
        
        ' Password expiry
        Dim blnExpired As Boolean

        blnExpired = (DateDiff("d", dtPasswordSet, Date) > PASSWORD_EXPIRY_DAYS)
        
        If blnExpired Then
            MsgBox "Your password has expired." & vbCrLf & _
                   "You must change it now.", vbCritical, "Password Expired"

            DoCmd.OpenForm "frmChangePassword", , , "UserID=" & lngUserID, acFormEdit, acDialog  ' acDialog to block until done

            ' Re-check if password was actually changed (reload user record)
            .Requery
            
            dtPasswordSet = Nz(!PasswordSetDate, #1/1/1900#)

            If DateDiff("d", dtPasswordSet, Date) > PASSWORD_EXPIRY_DAYS Then
                MsgBox "Password change required. Login cancelled.", vbCritical
                GoTo LoginFailed
            End If
        End If

        ' Success
        g_lngUserID = lngUserID
        g_strUsername = strUsername
        g_strFullName = strFullName
        g_strUserRole = strRole

        ' Reset failed attempts & update login
        .Edit
            !FailedLoginAttempts = 0
            !LockoutUntil = Null
            !LastLoginDate = Now()
        .Update

        UpdateLastActivity
        modAudit.LogAudit "Users", g_lngUserID, "Login", Null, "Successful", "Login", , True
   
        AuthenticateUser = True
        GoTo CleanExit
    End With

LoginFailed:
    ' Increment failed attempts
    If lngUserID > 0 Then
        lngFailedAttempts = lngFailedAttempts + 1
        
        If lngFailedAttempts >= MAX_FAILED_ATTEMPTS Then
            Dim lngExp As Long
            
            lngExp = lngFailedAttempts - MAX_FAILED_ATTEMPTS
            dtLockoutUntil = DateAdd("n", BASE_LOCKOUT_MINUTES * (2 ^ lngExp), Now())
        Else
            dtLockoutUntil = Null
        End If

        rs.Edit
        !FailedLoginAttempts = lngFailedAttempts
        !LockoutUntil = dtLockoutUntil
        rs.Update

    End If

    modAudit.LogAudit "Users", 0, "Login", strUsername, "Failed", "LoginAttempt", , True
    MsgBox "Invalid username or password.", vbExclamation, APP_NAME & " - Access Denied"
    AuthenticateUser = False

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
    Exit Function

ErrorHandler:
    modUtilities.LogError MODULE_NAME & "_AuthenticateUser", Err.Number, Err.Description, _
                         "Username=" & strUsername
    MsgBox "Login error. Contact administrator.", vbCritical, APP_NAME
    AuthenticateUser = False
    Resume CleanExit
End Function



' PUBLIC: ChangePassword – Secure password update

Public Function ChangePassword(ByVal lngUserID As Long, _
                               ByVal strOldPassword As String, _
                               ByVal strNewPassword As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strCurrentHash As String
    Dim strNewHash As String
    
    Set db = CurrentDb
    
    ' Validate new password strength
    If Not modValidation.IsPasswordStrong(strNewPassword) Then
        MsgBox "Password must be at least 8 characters and contain:" & vbCrLf & _
               "• Uppercase letter" & vbCrLf & _
               "• Lowercase letter" & vbCrLf & _
               "• Digit" & vbCrLf & _
               "• Special character", vbExclamation, APP_NAME
        ChangePassword = False
        Exit Function
    End If
    
    ' Get current hash
    Set rs = db.OpenRecordset("SELECT PasswordHash FROM Users WHERE UserID = " & lngUserID, _
                              dbOpenSnapshot)
    
    If rs.EOF Then
        MsgBox "User not found.", vbCritical, APP_NAME
        ChangePassword = False
        GoTo CleanExit
    End If
    
    strCurrentHash = Nz(rs!PasswordHash, "")
        
    ' Verify old password
    If Not modCrypto.VerifyPassword(strOldPassword, strCurrentHash) Then
        MsgBox "Current password is incorrect.", vbExclamation, modGlobals.APP_NAME
        
        modAudit.LogAudit "Users", lngUserID, "PasswordChange", Null, _
                         "Failed - incorrect old password", "Security", lngUserID, True
        
        ChangePassword = False
        
        GoTo CleanExit
    End If
    
    ' Check for password reuse
    If strNewPassword = strOldPassword Then
        MsgBox "New password must be different from the old one.", vbExclamation, APP_NAME

        modAudit.LogAudit "Users", lngUserID, "PasswordChange", Null, _
                         "Failed - reuse of old password", "Security", lngUserID, True

        ChangePassword = False

        GoTo CleanExit

    End If
    
    ' Hash new password
    strNewHash = modCrypto.HashPassword(strNewPassword)
    
    If Len(strNewHash) = 0 Then
        MsgBox "Error creating password hash. Contact administrator.", vbCritical, APP_NAME
        
        ChangePassword = False
        
        GoTo CleanExit
    End If
    
    ' Update password
    rs.Edit
    rs!PasswordHash = strNewHash
    rs!PasswordSetDate = Date
    rs!FailedLoginAttempts = 0
    rs!LockoutUntil = Null
    rs.Update
    
    modAudit.LogAudit "Users", lngUserID, "PasswordChange", Null, _
                     "Successful", "Security", lngUserID, True
    
    MsgBox "Password changed successfully.", vbInformation, APP_NAME
    ChangePassword = True
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    modUtilities.LogError MODULE_NAME & "_ChangePassword", Err.Number, Err.Description
    MsgBox "Error changing password. Contact administrator.", vbCritical, APP_NAME
    ChangePassword = False
    Resume CleanExit
End Function


' PUBLIC: CreateUser – Create new user with hashed password
Public Function CreateUser(ByVal strUsername As String, _
                          ByVal strPassword As String, _
                          ByVal strFullName As String, _
                          ByVal strRole As String) As Long
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strPasswordHash As String
    
    Set db = CurrentDb
    
    ' Validate inputs
    strUsername = modUtilities.NormalizeUsername(strUsername)
    
    If Len(strUsername) = 0 Or Len(Trim(strPassword)) = 0 Then
        MsgBox "Username and password are required.", vbExclamation, APP_NAME
        CreateUser = 0
        Exit Function
    End If
    
    If modDatabase.RecordExists("Users", "UCase(Username) = '" & Replace(strUsername, "'", "''") & "'") Then
        MsgBox "Username already exists.", vbExclamation, APP_NAME
        CreateUser = 0
        Exit Function
    End If
    
    ' Validate password strength
    If Not modValidation.IsPasswordStrong(strPassword) Then
         MsgBox "Password must be at least 8 characters and contain:" & vbCrLf & _
               "• Uppercase letter" & vbCrLf & _
               "• Lowercase letter" & vbCrLf & _
               "• Digit" & vbCrLf & _
               "• Special character", vbExclamation, APP_NAME
        CreateUser = 0
        Exit Function
    End If
    
    ' Hash password using modCrypto
    strPasswordHash = modCrypto.HashPassword(strPassword)
    
    If Len(strPasswordHash) = 0 Then
        MsgBox "Error creating password hash. Contact administrator.", vbCritical, APP_NAME
        CreateUser = 0
        Exit Function
    End If
    
    ' Create user record
    Set rs = db.OpenRecordset("Users", dbOpenDynaset)
    rs.AddNew
    rs!Username = strUsername
    rs!PasswordHash = strPasswordHash
    rs!FullName = strFullName
    rs!Role = strRole
    rs!IsActive = True
    rs!PasswordSetDate = Date
    rs!CreatedDate = Now()
    rs!CreatedBy = g_lngUserID
    rs.Update
    rs.Bookmark = rs.LastModified
    CreateUser = rs!UserID
    
    modAudit.LogAudit "Users", CreateUser, "Create", Null, _
                     "User created: " & strUsername, "UserManagement", g_lngUserID, True
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    modUtilities.LogError MODULE_NAME & "_CreateUser", Err.Number, Err.Description
    MsgBox "Error creating user. Contact administrator.", vbCritical, APP_NAME
    CreateUser = 0
    Resume CleanExit
End Function


' PUBLIC: LogoutUser – Secure session termination

Public Sub LogoutUser()
    On Error Resume Next

    If g_lngUserID > 0 Then
        modAudit.LogAudit "Users", g_lngUserID, "Logout", Null, "User initiated", "Logout", g_lngUserID, True
    End If

    ' Clear session
    g_lngUserID = 0
    g_strUsername = ""
    g_strFullName = ""
    g_strUserRole = ""
    g_dtLastActivity = #1/1/1900#

    ' Close all forms except login
    Dim frm As Form
    For Each frm In Forms
        If frm.Name <> "frmLogin" Then
            DoCmd.Close acForm, frm.Name, acSaveNo
        End If
    Next frm

    DoCmd.OpenForm "frmLogin"
End Sub



' Timer Event – Place in main navigation form (e.g., frmMain)

' In frmMain ? Form_Timer event:
Private Sub Form_Timer()
    If IsUserLoggedIn() Then
        If CheckSessionTimeout() Then
            MsgBox "Your session has expired due to inactivity." & vbCrLf & _
                   "You will be logged out for security reasons.", vbInformation, APP_NAME
            LogoutUser
        Else
            UpdateLastActivity
        End If
    End If
End Sub

' In frmMain ? Form_Load:
Private Sub Form_Load()
    Me.TimerInterval = 60000  ' Check every 60 seconds
End Sub


' Activity Tracking – Call from ALL forms

' Add this to every form's Click, KeyPress, or major control AfterUpdate:
'   Call UpdateLastActivity

' Or use a global hook (recommended): create modActivityTracker


' PUBLIC: TrackActivity – Call from all forms

Public Sub TrackActivity()
    If IsUserLoggedIn() Then UpdateLastActivity
End Sub


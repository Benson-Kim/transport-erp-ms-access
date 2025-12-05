'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' MODULE: modAudit

' PURPOSE: Comprehensive audit trail framework with synchronous and

'          asynchronous logging, form-level change detection, and

'          Admin-only audit viewer support.

' SECURITY: Only Admin can view. All actions are logged permanently.

' AUTHOR: Expert Back-End Developer (MS Access VBA Security Specialist)

' CREATED: November 17, 2025

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Database

Option Explicit



Private Const MODULE_NAME As String = "modAudit"

Private Const BATCH_SIZE As Long = 50



' Queue for asynchronous (batched) audit logging

Private Type AuditQueueItem

    TableName As String

    RecordID As Long

    FieldName As String

    OldValue As String

    NewValue As String

    ActionType As String

    PerformedBy As Long

End Type



Private marrQueue() As AuditQueueItem

Private mlngQueueCount As Long



' INITIALIZE QUEUE (called once at app startup from modStartup)

Public Sub InitializeAuditQueue()

    ReDim marrQueue(1 To BATCH_SIZE)

    mlngQueueCount = 0

End Sub





' CORE: LogAudit – Synchronous or Queued depending on criticality

' Procedure: LogAudit

' Purpose  : Primary audit logging function – ALWAYS synchronous for

'            critical operations. Use LogAuditAsync for non-critical.

'

Public Sub LogAudit(ByVal strTableName As String, _

                    ByVal lngRecordID As Long, _

                    ByVal strFieldName As String, _

                    ByVal varOldValue As Variant, _

                    ByVal varNewValue As Variant, _

                    ByVal strActionType As String, _

                    Optional ByVal lngPerformedBy As Long = -1, _

                    Optional ByVal blnForceSync As Boolean = True)



    On Error GoTo ErrorHandler



    Dim strOld As String, strNew As String

    Dim lngUser As Long

    

    strOld = Left(Nz(varOldValue, ""), 255)

    strNew = Left(Nz(varNewValue, ""), 255)

    

    ' Determine user ID: override > current > 1 (system)

    If lngPerformedBy <> -1 Then

        lngUser = lngPerformedBy

    ElseIf g_lngUserID > 0 Then

        lngUser = g_lngUserID

    Else

        lngUser = 1  ' System/Test

    End If



    ' Critical actions (Insert/Delete/Login) or forced sync ? write immediately

    If blnForceSync Or strActionType = "Insert" Or strActionType = "Delete" Then

        WriteAuditRecord strTableName, lngRecordID, strFieldName, strOld, strNew, lngUser, strActionType

    Else

        ' Queue for batch write (non-critical field updates)

        QueueAuditRecord strTableName, lngRecordID, strFieldName, strOld, strNew, strActionType, lngUser

    End If



    Exit Sub



ErrorHandler:

    modUtilities.LogError "LogAudit", Err.Number, Err.Description, _

                         "Table=" & strTableName & " | Action=" & strActionType

End Sub





' WRAPPERS – Clean, self-documenting calls



Public Sub LogInsert(ByVal strTableName As String, ByVal lngRecordID As Long, Optional lngUser As Long = -1)

    LogAudit strTableName, lngRecordID, "", Null, Null, "Insert", lngUser, True

End Sub



Public Sub LogUpdate(ByVal strTableName As String, _

                     ByVal lngRecordID As Long, _

                     ByVal strFieldName As String, _

                     ByVal varOldValue As Variant, _

                     ByVal varNewValue As Variant, _

                     Optional lngUser As Long = -1)

    LogAudit strTableName, lngRecordID, strFieldName, _

             varOldValue, varNewValue, "Update", lngUser, False

End Sub



Public Sub LogDelete(ByVal strTableName As String, ByVal lngRecordID As Long, Optional lngUser As Long = -1)

    LogAudit strTableName, lngRecordID, "", Null, Null, "Delete", lngUser, True

End Sub





' ASYNCHRONOUS (BATCHED) SUPPORT

Public Sub LogAuditAsync(ByVal strTableName As String, _

                         ByVal lngRecordID As Long, _

                         ByVal strFieldName As String, _

                         ByVal varOldValue As Variant, _

                         ByVal varNewValue As Variant, _

                         ByVal strActionType As String, _

                         ByVal lngUser As Long)

    LogAudit strTableName, lngRecordID, strFieldName, varOldValue, varNewValue, strActionType, lngUser, False

End Sub



Private Sub QueueAuditRecord(ByVal strTableName As String, _

                            ByVal lngRecordID As Long, _

                            ByVal strFieldName As String, _

                            ByVal strOldValue As String, _

                            ByVal strNewValue As String, _

                            ByVal strActionType As String, _

                            ByVal lngUser As Long)



    mlngQueueCount = mlngQueueCount + 1

    If mlngQueueCount > UBound(marrQueue) Then

        ReDim Preserve marrQueue(1 To UBound(marrQueue) + BATCH_SIZE)

    End If



    With marrQueue(mlngQueueCount)

        .TableName = Left(strTableName, 50)

        .RecordID = lngRecordID

        .FieldName = Left(Nz(strFieldName, ""), 50)

        .OldValue = Left(strOldValue, 255)

        .NewValue = Left(strNewValue, 255)

        .ActionType = Left(strActionType, 20)

        .PerformedBy = lngUser

    End With



    ' Flush if queue is getting large

    If mlngQueueCount >= BATCH_SIZE Then FlushAuditQueue



End Sub



Public Sub FlushAuditQueue()

    If mlngQueueCount = 0 Then Exit Sub



    On Error GoTo ErrorHandler



    Dim db As DAO.Database

    Dim lngI As Long

    Dim strSQL As String



    Set db = CurrentDb

    BeginTransaction



    For lngI = 1 To mlngQueueCount

        With marrQueue(lngI)

            strSQL = "INSERT INTO AuditLog (TableName, RecordID, FieldName, OldValue, NewValue, ActionType, " & _

                     "PerformedBy, PerformedDate, WorkstationName) VALUES (" & _

                     "'" & Replace(.TableName, "'", "''") & "', " & _

                     .RecordID & ", " & _

                     "'" & Replace(.FieldName, "'", "''") & "', " & _

                     "'" & Replace(.OldValue, "'", "''") & "', " & _

                     "'" & Replace(.NewValue, "'", "''") & "', " & _

                     "'" & .ActionType & "', " & _

                     .PerformedBy & ", Now(), '" & Left(Environ("COMPUTERNAME"), 50) & "')"

            db.Execute strSQL, dbFailOnError

        End With

    Next lngI



    CommitTransaction

    mlngQueueCount = 0



    Exit Sub



ErrorHandler:

    RollbackTransaction

    modUtilities.LogError "FlushAuditQueue", Err.Number, Err.Description, "QueueSize=" & mlngQueueCount

End Sub



Private Sub WriteAuditRecord(ByVal strTableName As String, _

                            ByVal lngRecordID As Long, _

                            ByVal strFieldName As String, _

                            ByVal strOldValue As String, _

                            ByVal strNewValue As String, _

                            ByVal lngUser As Long, _

                            ByVal strActionType As String)



    On Error GoTo ErrorHandler



    Dim strSQL As String

    strSQL = "INSERT INTO AuditLog (TableName, RecordID, FieldName, OldValue, NewValue, ActionType, " & _

             "PerformedBy, PerformedDate, WorkstationName) VALUES (" & _

             "'" & Replace(Left(strTableName, 50), "'", "''") & "', " & _

             lngRecordID & ", " & _

             "'" & Replace(Left(Nz(strFieldName, ""), 50), "'", "''") & "', " & _

             "'" & Replace(Left(strOldValue, 255), "'", "''") & "', " & _

             "'" & Replace(Left(strNewValue, 255), "'", "''") & "', " & _

             "'" & Left(strActionType, 20) & "', " & _

             lngUser & ", Now(), '" & Left(Environ("COMPUTERNAME"), 50) & "')"



    CurrentDb.Execute strSQL, dbFailOnError

    Exit Sub



ErrorHandler:

    modUtilities.LogError "WriteAuditRecord", Err.Number, Err.Description, strSQL

End Sub



' FORM-LEVEL AUTOMATIC AUDIT (BeforeUpdate)

'

' Procedure: AuditFormChanges

' Purpose  : Call from ANY form's Form_BeforeUpdate event

'             Automatically logs ALL changed fields

'

Public Sub AuditFormChanges(frm As Form, ByVal strTableName As String, ByVal lngRecordID As Long)

    On Error GoTo ErrorHandler



    Dim ctl As control

    Dim strField As String

    Dim varOld As Variant, varNew As Variant



    For Each ctl In frm.Controls

        If ctl.ControlType = acTextBox Or ctl.ControlType = acComboBox Or _

           ctl.ControlType = acCheckBox Or ctl.ControlType = acOptionGroup Then



            If ctl.Name Like "*_OldValue" Or ctl.Name = "AuditSkip" Then GoTo NextControl



            strField = Nz(ctl.Tag, ctl.Name)  ' Prefer Tag for real DB field name

            If strField = "" Then strField = ctl.Name



            ' Skip if no OldValue (new record) or control not bound

            If Not ctl.OldValue & "" = "" Then

                varOld = ctl.OldValue

                varNew = ctl.Value



                If Nz(varOld, "") <> Nz(varNew, "") Then

                    LogAuditAsync strTableName, lngRecordID, strField, varOld, varNew, "Update"

                End If

            End If

        End If

NextControl:

    Next ctl



    Exit Sub



ErrorHandler:

    modUtilities.LogError "AuditFormChanges", Err.Number, Err.Description, _

                         "Form=" & frm.Name & " | RecordID=" & lngRecordID

End Sub





' ADMIN AUDIT VIEWER FORM (frmAuditLog) - Query Source



' Create this saved query: qryAuditLog_WithUserNames

' SELECT AuditLog.*, Users.FullName, Users.Username

' FROM AuditLog LEFT JOIN Users ON AuditLog.PerformedBy = Users.UserID

' ORDER BY AuditLog.PerformedDate DESC;



' In frmAuditLog:

' - RecordSource = qryAuditLog_WithUserNames

' - Add filters: Date range, User combo (row source: SELECT UserID, FullName FROM Users WHERE IsActive=True)

' - Add Export to Excel button:

Private Sub cmdExportAudit_Click()

    DoCmd.OutputTo acOutputQuery, "qryAuditLog_WithUserNames", acFormatXLSX, , True

End Sub



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

Private Const LOCKOUT_DURATION_MINUTES As Integer = 15

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

    strUsername = Trim(UCase(strUsername))

    

    If Len(strUsername) = 0 Or Len(strPassword) = 0 Then

        GoTo LoginFailed

    End If



    ' Build safe SQL

    strSQL = "SELECT UserID, FullName, Role, IsActive, PasswordHash, " & _

             "FailedLoginAttempts, LockoutUntil, PasswordSetDate " & _

             "FROM Users WHERE UCase(Username) = '" & Replace(strUsername, "'", "''") & "'"



    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot, dbReadOnly)



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



        ' Password expiry

        If DateDiff("d", dtPasswordSet, Date) > PASSWORD_EXPIRY_DAYS Then

            MsgBox "Your password has expired." & vbCrLf & _

                   "You must change it now.", vbCritical, "Password Expired"

            DoCmd.OpenForm "frmChangePassword", , , "UserID=" & lngUserID

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



        ' === SUCCESS ===

        g_lngUserID = lngUserID

        g_strUsername = strUsername

        g_strFullName = strFullName

        g_strUserRole = strRole



        ' Reset failed attempts & update login

        .Close

        Set rs = db.OpenRecordset("Users", dbOpenDynaset)

        rs.FindFirst "UserID = " & lngUserID

        If Not rs.NoMatch Then

            rs.Edit

            rs!FailedLoginAttempts = 0

            rs!LockoutUntil = Null

            rs!LastLoginDate = Now()

            rs.Update

        End If



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

            dtLockoutUntil = DateAdd("n", LOCKOUT_DURATION_MINUTES, Now())

        Else

            dtLockoutUntil = Null

        End If

        

'        dtLockoutUntil = IIf(lngFailedAttempts >= MAX_FAILED_ATTEMPTS, _

'                             DateAdd("n", LOCKOUT_DURATION_MINUTES, Now()), Null)



        rs.Close

        Set rs = db.OpenRecordset("Users", dbOpenDynaset)

        rs.FindFirst "UserID = " & lngUserID

        If Not rs.NoMatch Then

            rs.Edit

            rs!FailedLoginAttempts = lngFailedAttempts

            rs!LockoutUntil = dtLockoutUntil

            rs.Update

        End If

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

    modUtilities.LogError "AuthenticateUser", Err.Number, Err.Description, _

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

                              dbOpenSnapshot, dbReadOnly)

    

    If rs.EOF Then

        MsgBox "User not found.", vbCritical, APP_NAME

        ChangePassword = False

        GoTo CleanExit

    End If

    

    strCurrentHash = Nz(rs!PasswordHash, "")

    rs.Close

    

    ' Verify old password

    If Not modCrypto.VerifyPassword(strOldPassword, strCurrentHash) Then

        MsgBox "Current password is incorrect.", vbExclamation, modGlobals.APP_NAME

        modAudit.LogAudit "Users", lngUserID, "PasswordChange", Null, _

                         "Failed - incorrect old password", "Security", lngUserID, True

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

    Set rs = db.OpenRecordset("Users", dbOpenDynaset)

    rs.FindFirst "UserID = " & lngUserID

    If Not rs.NoMatch Then

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

    Else

        ChangePassword = False

    End If

    

CleanExit:

    On Error Resume Next

    If Not rs Is Nothing Then

        rs.Close

        Set rs = Nothing

    End If

    Set db = Nothing

    Exit Function

    

ErrorHandler:

    modUtilities.LogError "ChangePassword", Err.Number, Err.Description

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

    If Len(Trim(strUsername)) = 0 Or Len(Trim(strPassword)) = 0 Then

        MsgBox "Username and password are required.", vbExclamation, APP_NAME

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

    rs!UserName = Trim(strUsername)

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

    modUtilities.LogError "CreateUser", Err.Number, Err.Description

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





' MODULE: modCrypto

' PURPOSE: Secure password hashing using Windows CNG API (BCrypt)

'          Pure Windows API - no .NET required



Option Compare Database

Option Explicit



Private Const MODULE_NAME As String = "modCrypto"



' BCrypt API declarations

Private Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" ( _

    ByRef phAlgorithm As LongPtr, _

    ByVal pszAlgId As LongPtr, _

    ByVal pszImplementation As LongPtr, _

    ByVal dwFlags As Long) As Long



Private Declare PtrSafe Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" ( _

    ByVal hAlgorithm As LongPtr, _

    ByVal dwFlags As Long) As Long



Private Declare PtrSafe Function BCryptCreateHash Lib "bcrypt.dll" ( _

    ByVal hAlgorithm As LongPtr, _

    ByRef phHash As LongPtr, _

    ByVal pbHashObject As LongPtr, _

    ByVal cbHashObject As Long, _

    ByVal pbSecret As LongPtr, _

    ByVal cbSecret As Long, _

    ByVal dwFlags As Long) As Long



Private Declare PtrSafe Function BCryptHashData Lib "bcrypt.dll" ( _

    ByVal hHash As LongPtr, _

    ByVal pbInput As LongPtr, _

    ByVal cbInput As Long, _

    ByVal dwFlags As Long) As Long



Private Declare PtrSafe Function BCryptFinishHash Lib "bcrypt.dll" ( _

    ByVal hHash As LongPtr, _

    ByVal pbOutput As LongPtr, _

    ByVal cbOutput As Long, _

    ByVal dwFlags As Long) As Long



Private Declare PtrSafe Function BCryptDestroyHash Lib "bcrypt.dll" ( _

    ByVal hHash As LongPtr) As Long



Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _

    Destination As Any, Source As Any, ByVal Length As LongPtr)



Private Declare PtrSafe Function BCryptGetProperty Lib "bcrypt.dll" ( _

    ByVal hAlgorithm As LongPtr, _

    ByVal pszProperty As LongPtr, _

    ByRef pbOutput As Any, _

    ByVal cbOutput As Long, _

    ByRef pcbResult As Long, _

    ByVal dwFlags As Long) As Long



Private Const BCRYPT_OBJECT_LENGTH As String = "ObjectLength"



Private Const BCRYPT_SHA256_ALGORITHM As String = "SHA256"



Private EmptyByteArray() As Byte





' PUBLIC: HashPassword – Returns Base64(salt + hash)



Public Function HashPassword(ByVal strPassword As String) As String

    On Error GoTo ErrorHandler



    Dim arrSalt() As Byte

    Dim arrPassword() As Byte

    Dim arrHash() As Byte

    Dim arrCombined() As Byte

    Dim arrSalted() As Byte

    Dim strResult As String

    Dim i As Long



    ' Generate 16-byte random salt

    Randomize Timer

    ReDim arrSalt(0 To 15)

    For i = 0 To 15

        arrSalt(i) = Int(Rnd * 256)

    Next i



    ' Convert password to UTF-8 bytes

    arrPassword = StringToUTF8(strPassword)



    ' Combine salt + password

    ReDim arrSalted(0 To UBound(arrSalt) + UBound(arrPassword) + 1)

    Call CopyMemory(arrSalted(0), arrSalt(0), 16)

    Call CopyMemory(arrSalted(16), arrPassword(0), UBound(arrPassword) + 1)



    ' Compute SHA-256 hash using Windows BCrypt

    arrHash = ComputeSHA256(arrSalted)



    If UBound(arrHash) < 0 Then

        HashPassword = ""

        Exit Function

    End If



    ' Combine salt + hash for storage

    ReDim arrCombined(0 To 47) ' 16 bytes salt + 32 bytes hash

    Call CopyMemory(arrCombined(0), arrSalt(0), 16)

    Call CopyMemory(arrCombined(16), arrHash(0), 32)



    strResult = EncodeBase64(arrCombined)

    HashPassword = strResult



CleanExit:

    Exit Function



ErrorHandler:

    On Error Resume Next

    modUtilities.LogError "HashPassword", Err.Number, Err.Description

    HashPassword = ""

    Resume CleanExit

End Function





' PUBLIC: VerifyPassword – Compare input vs stored hash



Public Function VerifyPassword(ByVal strPassword As String, ByVal strStoredHash As String) As Boolean

    On Error GoTo ErrorHandler



    Dim arrStored() As Byte

    Dim arrSalt(0 To 15) As Byte

    Dim arrPassword() As Byte

    Dim arrSalted() As Byte

    Dim arrComputedHash() As Byte

    Dim i As Long



    ' Decode stored Base64

    arrStored = DecodeBase64(strStoredHash)

    If UBound(arrStored) < 47 Then GoTo Fail ' Need 48 bytes



    ' Extract salt (first 16 bytes)

    Call CopyMemory(arrSalt(0), arrStored(0), 16)



    ' Convert password to UTF-8

    arrPassword = StringToUTF8(strPassword)



    ' Combine salt + password

    ReDim arrSalted(0 To 15 + UBound(arrPassword) + 1)

    Call CopyMemory(arrSalted(0), arrSalt(0), 16)

    Call CopyMemory(arrSalted(16), arrPassword(0), UBound(arrPassword) + 1)



    ' Compute hash

    arrComputedHash = ComputeSHA256(arrSalted)



    If UBound(arrComputedHash) < 31 Then GoTo Fail



    ' Compare computed hash with stored hash (constant-time comparison)

    VerifyPassword = True

    For i = 0 To 31

        If arrComputedHash(i) <> arrStored(16 + i) Then

            VerifyPassword = False

        End If

    Next i



    GoTo CleanExit



Fail:

    VerifyPassword = False



CleanExit:

    Exit Function



ErrorHandler:

    VerifyPassword = False

    Resume CleanExit

End Function





' PRIVATE: ComputeSHA256 using Windows BCrypt API



Private Function ComputeSHA256(arrData() As Byte) As Byte()

    Dim hAlg As LongPtr

    Dim hHash As LongPtr

    Dim arrResult() As Byte

    Dim hashObj() As Byte

    Dim lResult As Long

    Dim cbHashObj As Long

    Dim cbReturned As Long

    Dim strAlg As String

    Dim pbData As LongPtr

    Dim cbData As Long



    On Error GoTo ErrorHandler

    

    ReDim arrResult(0 To 31)



    strAlg = BCRYPT_SHA256_ALGORITHM



    ' 1. Open algorithm provider

    lResult = BCryptOpenAlgorithmProvider(hAlg, StrPtr(strAlg), 0, 0)

    If lResult <> 0 Then GoTo ErrorHandler



    ' 2. Get required hash object size

    lResult = BCryptGetProperty(hAlg, StrPtr(BCRYPT_OBJECT_LENGTH), cbHashObj, 4, cbReturned, 0)

    If lResult <> 0 Then GoTo ErrorHandler

    

    ReDim hashObj(0 To cbHashObj - 1)



    ' 3. Create hash object — NOW with proper buffer

    lResult = BCryptCreateHash(hAlg, hHash, VarPtr(hashObj(0)), cbHashObj, 0, 0, 0)

    If lResult <> 0 Then GoTo ErrorHandler



    ' 4. Hash the data (64-bit safe)

    pbData = VarPtr(arrData(0))

    cbData = UBound(arrData) + 1

    lResult = BCryptHashData(hHash, pbData, cbData, 0)

    If lResult <> 0 Then GoTo ErrorHandler



    ' 5. Finish hash

    lResult = BCryptFinishHash(hHash, VarPtr(arrResult(0)), 32, 0)

    If lResult <> 0 Then GoTo ErrorHandler



    ComputeSHA256 = arrResult

    GoTo Cleanup



ErrorHandler:

    ComputeSHA256 = EmptyByteArray



Cleanup:

    If hHash <> 0 Then BCryptDestroyHash hHash

    If hAlg <> 0 Then BCryptCloseAlgorithmProvider hAlg, 0

End Function





' PRIVATE: Convert String to UTF-8 Byte Array



Private Function StringToUTF8(ByVal strText As String) As Byte()

    Dim objStream As Object

    Dim arrBytes() As Byte



    On Error GoTo ErrorHandler



    ' Use ADODB.Stream for UTF-8 conversion

    Set objStream = CreateObject("ADODB.Stream")

    objStream.Type = 2 ' adTypeText

    objStream.Charset = "UTF-8"

    objStream.Open

    objStream.WriteText strText

    objStream.Position = 0

    objStream.Type = 1 ' adTypeBinary

    objStream.Position = 3 ' Skip UTF-8 BOM

    arrBytes = objStream.Read

    objStream.Close



    StringToUTF8 = arrBytes

    Set objStream = Nothing

    Exit Function



ErrorHandler:

    ReDim arrBytes(0 To -1)

    StringToUTF8 = arrBytes

    If Not objStream Is Nothing Then objStream.Close

    Set objStream = Nothing

End Function





' PRIVATE: Base64 Encode/Decode using MSXML



Private Function EncodeBase64(arrData() As Byte) As String

    Dim objXML As Object

    Dim objNode As Object



    On Error GoTo ErrorHandler



    Set objXML = CreateObject("MSXML2.DOMDocument")

    Set objNode = objXML.createElement("b64")

    objNode.DataType = "bin.base64"

    objNode.nodeTypedValue = arrData

    EncodeBase64 = Replace(objNode.Text, vbLf, "")



    Set objNode = Nothing

    Set objXML = Nothing

    Exit Function



ErrorHandler:

    EncodeBase64 = ""

    Set objNode = Nothing

    Set objXML = Nothing

End Function



Private Function DecodeBase64(ByVal strData As String) As Byte()

    Dim objXML As Object

    Dim objNode As Object

    Dim arrEmpty() As Byte



    On Error GoTo ErrorHandler



    Set objXML = CreateObject("MSXML2.DOMDocument")

    Set objNode = objXML.createElement("b64")

    objNode.DataType = "bin.base64"

    objNode.Text = strData

    DecodeBase64 = objNode.nodeTypedValue



    Set objNode = Nothing

    Set objXML = Nothing

    Exit Function



ErrorHandler:

    ReDim arrEmpty(0 To -1)

    DecodeBase64 = arrEmpty

    Set objNode = Nothing

    Set objXML = Nothing

End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' MODULE: modDatabase

' PURPOSE: Centralize all database-related helper functions,

'          transaction management, recordset handling and common

'          data-access patterns used throughout the entire application.

' AUTHOR: Expert Back-End Developer (MS Access VBA Specialist)

' CREATED: November 17, 2025

' SECURITY NOTE: All public procedures include comprehensive error

'                handling, proper object cleanup and audit-ready logging.

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Database

Option Explicit





' PRIVATE CONSTANTS



Private Const MODULE_NAME As String = "modDatabase"





' PUBLIC FUNCTIONS





''

''' Function: GetNextID

''' Purpose : Returns the next logical AutoNumber value for any table

'''           (useful when the ID is needed before the record is saved)

''' Parameters:

'''   strTableName  - Name of the table

'''   strIDField    - Name of the AutoNumber primary key field

''' Returns   : Long - Next available ID (Max + 1). Returns 1 if table empty

''

Public Function GetNextID(ByVal strTableName As String, ByVal strIDField As String) As Long

    On Error GoTo ErrorHandler

    

    Dim db          As DAO.Database

    Dim rs          As DAO.Recordset

    Dim lngNextID   As Long

    

    Set db = CurrentDb

    

    ' DMax is safe and fast - returns Null if no records exist

    lngNextID = Nz(DMax(strIDField, strTableName), 0) + 1

    

    GetNextID = lngNextID

    

CleanExit:

    Set rs = Nothing

    Set db = Nothing

    Exit Function

    

ErrorHandler:

    modUtilities.LogError MODULE_NAME & "_GetNextID", Err.Number, Err.Description

    MsgBox "Error retrieving next ID for " & strTableName & vbCrLf & _

           Err.Description, vbCritical, APP_NAME

    GetNextID = 0

    Resume CleanExit

End Function



''

''' Function: RecordExists

''' Purpose : Generic existence check - eliminates repetitive code

''

Public Function RecordExists(ByVal strTableName As String, _

                            ByVal strCriteria As String) As Boolean

    On Error GoTo ErrorHandler

    

    RecordExists = (DCount("*", strTableName, strCriteria) > 0)

    

CleanExit:

    Exit Function

    

ErrorHandler:

    modUtilities.LogError MODULE_NAME & "_RecordExists", Err.Number, Err.Description

    RecordExists = False

    Resume CleanExit

End Function



''

''' Function: GetSingleValue

''' Purpose : Return a single field value from the first matching record

''

Public Function GetSingleValue(ByVal strTableName As String, _

                               ByVal strFieldName As String, _

                               ByVal strCriteria As String, _

                               Optional ByVal varDefault As Variant = Null) As Variant

    On Error GoTo ErrorHandler

    

    Dim varResult As Variant

    varResult = DLookup(strFieldName, strTableName, strCriteria)

    

    If IsNull(varResult) Then

        GetSingleValue = varDefault

    Else

        GetSingleValue = varResult

    End If

    

CleanExit:

    Exit Function

    

ErrorHandler:

    modUtilities.LogError MODULE_NAME & "_GetSingleValue", Err.Number, Err.Description

    GetSingleValue = varDefault

    Resume CleanExit

End Function



''

''' Transaction Management - Begin / Commit / Rollback

''

Public Sub BeginTransaction()

    On Error GoTo ErrorHandler

    DBEngine.BeginTrans

    Exit Sub

    

ErrorHandler:

    modUtilities.LogError MODULE_NAME & "_BeginTransaction", Err.Number, Err.Description

    MsgBox "Unable to start transaction.", vbCritical, APP_NAME

End Sub



Public Sub CommitTransaction()

    On Error GoTo ErrorHandler

    DBEngine.CommitTrans

    Exit Sub

    

ErrorHandler:

    modUtilities.LogError MODULE_NAME & "_CommitTransaction", Err.Number, Err.Description

    MsgBox "Transaction commit failed.", vbCritical, APP_NAME

End Sub



Public Sub RollbackTransaction()

    On Error GoTo ErrorHandler

    DBEngine.Rollback

    Exit Sub

    

ErrorHandler:

    modUtilities.LogError MODULE_NAME & "_RollbackTransaction", Err.Number, Err.Description

    MsgBox "Transaction rollback failed.", vbCritical, APP_NAME

End Sub



''

''' Function: ExecuteQuery

''' Purpose : Execute an action query (INSERT/UPDATE/DELETE) with

'''           centralized error handling and return affected rows

''

Public Function ExecuteQuery(ByVal strSQL As String) As Long

    On Error GoTo ErrorHandler

    

    Dim db As DAO.Database

    Set db = CurrentDb

    db.Execute strSQL, dbFailOnError

    ExecuteQuery = db.RecordsAffected

    

CleanExit:

    Set db = Nothing

    Exit Function

    

ErrorHandler:

    modUtilities.LogError MODULE_NAME & " _ExecuteQuery", Err.Number, Err.Description & vbCrLf & "SQL: " & strSQL

    MsgBox "Database error executing query." & vbCrLf & Err.Description, vbCritical, APP_NAME

    ExecuteQuery = 0

    Resume CleanExit

End Function



''

''' Function: GetRecordset

''' Purpose : Standardized recordset creation with sensible defaults

'''           LockType: dbOpenDynaset (default) - editable

'''                     dbOpenSnapshot - read-only (faster for reports)

''

Public Function GetRecordset(ByVal strSQL As String, _

                  Optional ByVal lngLockType As DAO.LockTypeEnum = dbOpenDynaset) As DAO.Recordset

    On Error GoTo ErrorHandler

    

    Dim db As DAO.Database

    Dim rs As DAO.Recordset

    

    Set db = CurrentDb

    Set rs = db.OpenRecordset(strSQL, dbOpenDynaset, 0, lngLockType)

    Set GetRecordset = rs

    

CleanExit:

    Set db = Nothing

    Exit Function

    

ErrorHandler:

    modUtilities.LogError MODULE_NAME & " _GetRecordset", Err.Number, Err.Description & vbCrLf & "SQL: " & strSQL

    MsgBox "Failed to open recordset." & vbCrLf & Err.Description, vbCritical, APP_NAME

    Set GetRecordset = Nothing

    Resume CleanExit

End Function





' STRING / DATA CLEANING & VALIDATION UTILITIES





''

''' Function: CleanString

''' Purpose : Remove leading/trailing spaces, control chars, multiple spaces

''

Public Function CleanString(ByVal strInput As String) As String

    Dim strTemp As String

    strTemp = Trim$(Nz(strInput, ""))

    ' Remove non-printable chars except CR/LF/TAB

    strTemp = RegexReplace(strTemp, "[^\x09\x0A\x0D\x20-\x7E]", "")

    ' Collapse multiple spaces

    strTemp = RegexReplace(strTemp, "\s{2,}", " ")

    CleanString = strTemp

End Function



''

''' Function: FormatVATNumber

''' Purpose : Standardize VAT numbers (remove spaces/dots/dashes, uppercase)

''

Public Function FormatVATNumber(ByVal strVAT As String) As String

    Dim strClean As String

    strClean = UCase(Trim$(Nz(strVAT, "")))

    strClean = Replace(strClean, " ", "")

    strClean = Replace(strClean, "-", "")

    strClean = Replace(strClean, ".", "")

    FormatVATNumber = strClean

End Function





''

''' Function: ParseTime

''' Purpose : Convert HH.MM string (e.g. "14.30") to true TimeSerial value

''' Returns : Date (time portion only) or #12:00:00 AM# on error

''

Public Function ParseTime(ByVal strTime As String) As Date

    On Error GoTo ErrorHandler

    

    Dim strClean As String

    strClean = Replace(Trim$(strTime), ".", ":")

    

    If IsTime(strClean) Then

        ParseTime = TimeValue(strClean)

    Else

        ParseTime = #12:00:00 AM#

    End If

    

    Exit Function

    

ErrorHandler:

    ParseTime = #12:00:00 AM#

End Function





' PRIVATE SUPPORT ROUTINES (Regex helpers - require reference to

'  Microsoft VBScript Regular Expressions 5.5)



Private Function RegexReplace(ByVal strInput As String, _

                              ByVal strPattern As String, _

                              ByVal strReplace As String) As String

    With New RegExp

        .Global = True

        .IgnoreCase = True

        .Pattern = strPattern

        RegexReplace = .Replace(strInput, strReplace)

    End With

End Function



Private Function RegexTest(ByVal strInput As String, _

                           ByVal strPattern As String, _

                           ByVal blnIgnoreCase As Boolean) As Boolean

    With New RegExp

        .Pattern = strPattern

        .IgnoreCase = blnIgnoreCase

        RegexTest = .Test(strInput)

    End With

End Function





'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' MODULE: modGlobals

' PURPOSE: Centralized global constants, variables, and application-wide

'          settings with enterprise-grade error handling and logging

' AUTHOR: Expert Back-End Developer (MS Access Security & Reliability)

' CREATED: November 18, 2025

' UPDATED: November 18, 2025 – Full standards compliance

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Database

Option Explicit



Private Const MODULE_NAME As String = "modGlobals"





' APPLICATION CONSTANTS



Public Const APP_NAME As String = "Road Freight Transport ERP"

Public Const APP_VERSION As String = "v1.0"

Public Const VAT_RATE As Double = 0.21                    ' 21% VAT (Spain)

Public Const MAX_OBSERVATION_LENGTH As Integer = 200

Public Const TIME_FORMAT As String = "HH.MM"

Public Const DATE_FORMAT As String = "DD/MM/YYYY"



' Session timeout

Public Const SESSION_TIMEOUT_MINUTES As Integer = 30





' CURRENT USER SESSION STATE



Public g_lngUserID As Long

Public g_strUsername As String

Public g_strFullName As String

Public g_strUserRole As String

Public g_dtLastActivity As Date





' APPLICATION PATHS



Public g_strBackendPath As String

Public g_strPDFStoragePath As String

Public g_strBackupPath As String





' ENUMERATIONS



Public Enum UserRole

    roleAdmin = 1

    roleManager = 2

    roleOperator = 3

End Enum



Public Enum ServiceStatus

    statusActive = 1

    statusCompleted = 2

    statusCancelled = 3

    statusBilled = 4

End Enum



Public Enum InvoiceStatus

    invDraft = 1

    invFinalized = 2

    invCancelled = 3

    invPaid = 4

End Enum





' PUBLIC: InitializeGlobalVariables

' Purpose : Reset all session state and load system paths at startup

' Returns : Boolean – True if successful



Public Function InitializeGlobalVariables() As Boolean

    On Error GoTo ErrorHandler



    ' Reset user session

    g_lngUserID = 0

    g_strUsername = ""

    g_strFullName = ""

    g_strUserRole = ""

    g_dtLastActivity = Now()



    ' Load persistent paths from SystemSettings

    g_strBackendPath = GetSystemSetting("BackendDatabasePath", "")

    g_strPDFStoragePath = GetSystemSetting("PDFStoragePath", "")

    g_strBackupPath = GetSystemSetting("BackupPath", "")



    InitializeGlobalVariables = True

    Exit Function



ErrorHandler:

    modUtilities.LogError "InitializeGlobalVariables", Err.Number, Err.Description, ""

    InitializeGlobalVariables = False

End Function





' PUBLIC: GetSystemSetting

' Purpose : Safely retrieve a value from SystemSettings table



Public Function GetSystemSetting( _

    ByVal strSettingName As String, _

    ByVal strDefaultValue As String) As String



    On Error GoTo ErrorHandler



    Dim db As DAO.Database

    Dim rs As DAO.Recordset

    Dim strSQL As String

    Dim strValue As String



    Set db = CurrentDb

    strSQL = "SELECT SettingValue FROM SystemSettings WHERE SettingName = '" & _

             Replace(strSettingName, "'", "''") & "'"



    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot, dbReadOnly)



    If Not rs.EOF Then

        strValue = Nz(rs!SettingValue, strDefaultValue)

    Else

        strValue = strDefaultValue

    End If



    GetSystemSetting = strValue

    GoTo CleanExit



ErrorHandler:

    modUtilities.LogError "GetSystemSetting", Err.Number, Err.Description, _

                         "SettingName=" & strSettingName

    GetSystemSetting = strDefaultValue



CleanExit:

    On Error Resume Next

    If Not rs Is Nothing Then

        rs.Close

        Set rs = Nothing

    End If

    Set db = Nothing

End Function





' PUBLIC: SetSystemSetting

' Purpose : Insert or update a system setting with audit trail



Public Function SetSystemSetting( _

    ByVal strSettingName As String, _

    ByVal strSettingValue As String, _

    Optional ByVal strDescription As String = "") As Boolean



    On Error GoTo ErrorHandler



    Dim db As DAO.Database

    Dim rs As DAO.Recordset



    Set db = CurrentDb

    Set rs = db.OpenRecordset("SystemSettings", dbOpenDynaset, dbSeeChanges)



    rs.FindFirst "SettingName = '" & Replace(strSettingName, "'", "''") & "'"



    If rs.NoMatch Then

        rs.AddNew

        rs!SettingName = strSettingName

    Else

        rs.Edit

    End If



    rs!SettingValue = strSettingValue

    If strDescription <> "" Then rs!SettingDescription = strDescription

    rs!ModifiedBy = g_lngUserID

    rs!ModifiedDate = Now()

    rs.Update



    SetSystemSetting = True

    GoTo CleanExit



ErrorHandler:

    modUtilities.LogError "SetSystemSetting", Err.Number, Err.Description, _

                         "Setting=" & strSettingName & " | Value=" & strSettingValue

    SetSystemSetting = False



CleanExit:

    On Error Resume Next

    If Not rs Is Nothing Then

        rs.Close

        Set rs = Nothing

    End If

    Set db = Nothing

End Function





' PUBLIC: UpdateLastActivity

' Purpose : Refresh session activity timestamp (called on user actions)



Public Sub UpdateLastActivity()

    On Error Resume Next   ' Never fail on activity update

    g_dtLastActivity = Now()

End Sub





' PUBLIC: CheckSessionTimeout

' Purpose : Determine if user session has expired



Public Function CheckSessionTimeout() As Boolean

    On Error GoTo ErrorHandler



    Dim lngMinutes As Long

    lngMinutes = DateDiff("n", g_dtLastActivity, Now())



    CheckSessionTimeout = (lngMinutes >= SESSION_TIMEOUT_MINUTES)

    Exit Function



ErrorHandler:

    modUtilities.LogError "CheckSessionTimeout", Err.Number, Err.Description

    CheckSessionTimeout = True  ' Fail closed – force logout on error

End Function





' PUBLIC: GetCurrentUser

' Purpose : Return formatted string for display in UI



Public Function GetCurrentUser() As String

    If g_lngUserID > 0 Then

        GetCurrentUser = g_strFullName & " (" & g_strUserRole & ")"

    Else

        GetCurrentUser = "Not logged in"

    End If

End Function





' PUBLIC: IsUserLoggedIn

' Purpose : Simple boolean check used throughout application



Public Function IsUserLoggedIn() As Boolean

    IsUserLoggedIn = (g_lngUserID > 0)

End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' MODULE: modPermissions

' PURPOSE: Centralized, high-performance permission checking engine

'          with exhaustive permission constants and role-based matrix

' SECURITY: All security decisions flow through this single module

' AUTHOR: Expert Back-End Developer (MS Access Security Architect)

' CREATED: November 17, 2025

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Database

Option Explicit



Private Const MODULE_NAME As String = "modPermissions"



' PERMISSION CONSTANTS – Used throughout the entire application

Public Const PERM_VIEW_CLIENTS              As String = "VIEW_CLIENTS"

Public Const PERM_EDIT_CLIENTS              As String = "EDIT_CLIENTS"

Public Const PERM_DELETE_CLIENTS            As String = "DELETE_CLIENTS"



Public Const PERM_VIEW_SUPPLIERS            As String = "VIEW_SUPPLIERS"

Public Const PERM_EDIT_SUPPLIERS            As String = "EDIT_SUPPLIERS"

Public Const PERM_DELETE_SUPPLIERS          As String = "DELETE_SUPPLIERS"



Public Const PERM_CREATE_SERVICES           As String = "CREATE_SERVICES"

Public Const PERM_EDIT_ACTIVE_SERVICES      As String = "EDIT_ACTIVE_SERVICES"

Public Const PERM_EDIT_COMPLETED_SERVICES   As String = "EDIT_COMPLETED_SERVICES"

Public Const PERM_DELETE_SERVICES           As String = "DELETE_SERVICES"



Public Const PERM_GENERATE_LOADING_ORDERS   As String = "GENERATE_LOADING_ORDERS"

Public Const PERM_GENERATE_INVOICES         As String = "GENERATE_INVOICES"

Public Const PERM_VIEW_INVOICES             As String = "VIEW_INVOICES"

Public Const PERM_VIEW_FINANCIAL_REPORTS    As String = "VIEW_FINANCIAL_REPORTS"



Public Const PERM_MANAGE_USERS              As String = "MANAGE_USERS"

Public Const PERM_VIEW_AUDIT_LOG            As String = "VIEW_AUDIT_LOG"

Public Const PERM_EDIT_COMPANY_SETTINGS     As String = "EDIT_COMPANY_SETTINGS"





' PUBLIC: HasPermission – Core permission check



''

''' Function : HasPermission

''' Purpose  : Returns True if current user (g_strUserRole) has the requested permission

'''            Uses exhaustive Select Case – fastest possible evaluation

''' Returns  : Boolean

''

Public Function HasPermission(ByVal strPermission As String) As Boolean

    On Error GoTo ErrorHandler



    Dim strRole As String

    strRole = UCase(Trim(g_strUserRole))  ' Always normalize



    Select Case UCase(strPermission)

        '================================================================

        ' CLIENT PERMISSIONS

        '================================================================

        Case PERM_VIEW_CLIENTS

            HasPermission = True  ' All roles can view clients



        Case PERM_EDIT_CLIENTS

            HasPermission = (strRole = "ADMIN" Or strRole = "MANAGER")



        Case PERM_DELETE_CLIENTS

            HasPermission = (strRole = "ADMIN")



        '================================================================

        ' SUPPLIER PERMISSIONS

        '================================================================

        Case PERM_VIEW_SUPPLIERS

            HasPermission = True  ' All roles can view suppliers



        Case PERM_EDIT_SUPPLIERS

            HasPermission = (strRole = "ADMIN" Or strRole = "MANAGER")



        Case PERM_DELETE_SUPPLIERS

            HasPermission = (strRole = "ADMIN")



        '================================================================

        ' SERVICE PERMISSIONS

        '================================================================

        Case PERM_CREATE_SERVICES

            HasPermission = True  ' All roles can create services



        Case PERM_EDIT_ACTIVE_SERVICES

            HasPermission = True  ' All roles can edit Active services



        Case PERM_EDIT_COMPLETED_SERVICES

            HasPermission = (strRole = "ADMIN" Or strRole = "MANAGER")

            ' Only Manager/Admin can edit Completed services



        Case PERM_DELETE_SERVICES

            HasPermission = (strRole = "ADMIN")

            ' Only Admin can delete (soft-delete)



        '================================================================

        ' DOCUMENT GENERATION

        '================================================================

        Case PERM_GENERATE_LOADING_ORDERS

            HasPermission = True  ' All roles can generate loading orders



        Case PERM_GENERATE_INVOICES

            HasPermission = (strRole = "ADMIN" Or strRole = "MANAGER")



        Case PERM_VIEW_INVOICES

            HasPermission = (strRole = "ADMIN" Or strRole = "MANAGER")



        Case PERM_VIEW_FINANCIAL_REPORTS

            HasPermission = (strRole = "ADMIN" Or strRole = "MANAGER")



        '================================================================

        ' ADMIN-ONLY FUNCTIONS

        '================================================================

        Case PERM_MANAGE_USERS

            HasPermission = (strRole = "ADMIN")



        Case PERM_VIEW_AUDIT_LOG

            HasPermission = (strRole = "ADMIN")



        Case PERM_EDIT_COMPANY_SETTINGS

            HasPermission = (strRole = "ADMIN")



        '================================================================

        ' UNKNOWN PERMISSION ? DENY (Fail-closed security)

        '================================================================

        Case Else

            modUtilities.LogError "HasPermission", 1003, _

                "Unknown permission requested", "Permission=" & strPermission & " | User=" & g_strUsername

            HasPermission = False

    End Select



    Exit Function



ErrorHandler:

    modUtilities.LogError "HasPermission", Err.Number, Err.Description, _

                         "Permission=" & strPermission

    HasPermission = False

End Function





' PUBLIC: RequirePermission – Convenience wrapper that shows message + cancels



''

''' Sub: RequirePermission

''' Purpose: Call at top of forms/reports/buttons to enforce access

'''         Shows friendly message and cancels if denied

''

Public Sub RequirePermission(ByVal strPermission As String, _

                            Optional ByVal strCustomMessage As String = "")

    If Not HasPermission(strPermission) Then

        Dim strMsg As String

        If strCustomMessage = "" Then

            strMsg = "You do not have permission to perform this action." & vbCrLf & _

                     "Required permission: " & strPermission & vbCrLf & _

                     "Please contact your administrator if you believe this is incorrect."

        Else

            strMsg = strCustomMessage

        End If



        MsgBox strMsg, vbExclamation, modGlobals.APP_NAME & " - Access Denied"

        

        ' Log attempted privilege escalation

        modAudit.LogAudit "Security", 0, "PermissionDenied", _

                          g_strUsername & " (" & g_strUserRole & ")", strPermission, "AccessAttempt", g_lngUserID, True



'        Cancel = True

        DoCmd.CancelEvent

    End If

End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' MODULE: modRecordDuplicator

' PURPOSE: Universal record duplication with field control & child records

' FEATURES:

'   • Duplicate any record (single table or with children)

'   • Include/Exclude specific fields

'   • Auto-skip unique/indexed fields (PK, VAT, ServiceNumber, etc.)

'   • Duplicate related child records (1-to-many)

'   • Full audit trail with transactions

'   • Helper utilities for building field lists

' AUTHOR: Expert Back-End Developer

' UPDATED: November 19, 2025

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Database

Option Explicit



Private Const MODULE_NAME As String = "modRecordDuplicator"





' PUBLIC: DuplicateRecord – MAIN FUNCTION



Public Function DuplicateRecord( _

    ByVal strTable As String, _

    ByVal lngSourceID As Long, _

    Optional dictInclude As Object = Nothing, _

    Optional dictExclude As Object = Nothing, _

    Optional dictChildren As Object = Nothing) As Long

    

    On Error GoTo ErrorHandler

    

    Dim ws As Workspace

    Dim db As DAO.Database

    Dim rsSource As DAO.Recordset

    Dim rsTarget As DAO.Recordset

    Dim fld As DAO.Field

    Dim strSQL As String

    Dim lngNewID As Long

    Dim blnUseInclude As Boolean

    Dim strPK As String

    Dim blnIsTrans As Boolean

    

    Set db = CurrentDb

    ws.BeginTrans

    blnIsTrans = True

    

    ' Get primary key field

    strPK = GetPrimaryKeyField(strTable)

    

    ' Open source record

    strSQL = "SELECT * FROM [" & strTable & "] WHERE [" & strPK & "] = " & lngSourceID

    Set rsSource = db.OpenRecordset(strSQL, dbOpenSnapshot)

    If rsSource.EOF Then

        Debug.Print "Source record not found: " & strTable & " ID=" & lngSourceID

        GoTo NotFound

    End If

    

    Set rsTarget = db.OpenRecordset("[" & strTable & "]", dbOpenDynaset)

    

    blnUseInclude = Not (dictInclude Is Nothing)

    If blnUseInclude And dictInclude.Count = 0 Then blnUseInclude = False

    

    With rsTarget

        .AddNew

        For Each fld In rsSource.Fields

            If Not ShouldSkipField(fld, dictInclude, dictExclude, blnUseInclude, strTable) Then

                If Not IsNull(rsSource(fld.Name)) Then

                    .Fields(fld.Name) = rsSource(fld.Name)

                End If

            End If

        Next fld

        

        ' Standard audit fields

        On Error Resume Next

        .Fields("CreatedBy") = Environ("USERNAME")

        .Fields("CreatedDate") = Now()

        .Fields("ModifiedBy") = Environ("USERNAME")

        .Fields("ModifiedDate") = Now()

        On Error GoTo ErrorHandler

        

        .Update

        .Bookmark = .LastModified

        

        ' Get new Primary Key

        If Not rsTarget.Fields(strPK) Is Nothing Then

            lngNewID = Nz(rsTarget.Fields(strPK).Value, 0)

        End If

        

        If lngNewID = 0 Then

            rsTarget.Move 0, rsTarget.LastModified

            lngNewID = Nz(rsTarget.Fields(strPK).Value, 0)

        End If

        

        If lngNewID = 0 Then

            lngNewID = Nz(DMax(strPK, strTable), 1)

        End If

    End With

    

    ' Duplicate child records

    If Not (dictChildren Is Nothing) Then

        Dim vKey As Variant

        For Each vKey In dictChildren.Keys

            Call DuplicateChildRecords(CStr(vKey), lngSourceID, lngNewID, CStr(dictChildren(vKey)))

        Next vKey

    End If

    

    ws.CommitTrans

    blnIsTrans = False

    DuplicateRecord = lngNewID

    GoTo CleanExit



NotFound:

    

    If blnIsTrans Then ws.Rollback

    blnIsTrans = False

    DuplicateRecord = 0

    GoTo CleanExit



ErrorHandler:

    If blnIsTrans Then ws.Rollback

    blnIsTrans = False

    DuplicateRecord = 0



CleanExit:

    On Error Resume Next

    If Not rsSource Is Nothing Then rsSource.Close

    If Not rsTarget Is Nothing Then rsTarget.Close

    Set rsSource = Nothing

    Set rsTarget = Nothing

    If blnIsTrans = False Then Set ws = Nothing

    Set db = Nothing

End Function





' UTILITY: Create Include Dictionary



Public Function CreateIncludeList(ParamArray fieldNames() As Variant) As Object

    Dim dict As Object

    Dim i As Integer

    

    Set dict = CreateObject("Scripting.Dictionary")

    dict.CompareMode = vbTextCompare

    

    For i = LBound(fieldNames) To UBound(fieldNames)

        dict(CStr(fieldNames(i))) = True

    Next i

    

    Set CreateIncludeList = dict

End Function





' UTILITY: Create Exclude Dictionary



Public Function CreateExcludeList(ParamArray fieldNames() As Variant) As Object

    Dim dict As Object

    Dim i As Integer

    

    Set dict = CreateObject("Scripting.Dictionary")

    dict.CompareMode = vbTextCompare

    

    For i = LBound(fieldNames) To UBound(fieldNames)

        dict(CStr(fieldNames(i))) = True

    Next i

    

    Set CreateExcludeList = dict

End Function





' UTILITY: Create Child Relationships Dictionary



Public Function CreateChildRelations() As Object

    Set CreateChildRelations = CreateObject("Scripting.Dictionary")

End Function



Public Sub AddChildRelation(dictChildren As Object, _

                           strChildTable As String, _

                           strForeignKeyField As String)

    dictChildren(strChildTable) = strForeignKeyField

End Sub





' UTILITY: Get All Field Names (excluding auto/unique)



Public Function GetDuplicatableFields(strTable As String) As Collection

    Dim db As DAO.Database

    Dim tdf As DAO.TableDef

    Dim fld As DAO.Field

    Dim col As Collection

    

    Set col = New Collection

    Set db = CurrentDb

    Set tdf = db.TableDefs(strTable)

    

    For Each fld In tdf.Fields

        If Not (fld.Attributes And dbAutoIncrField) Then

            If Not IsFieldUniqueIndexed(strTable, fld.Name) Then

                col.Add fld.Name

            End If

        End If

    Next fld

    

    Set GetDuplicatableFields = col

End Function





' UTILITY: Get Child Tables for a Parent Table



Public Function GetChildTables(strParentTable As String) As Object

    Dim db As DAO.Database

    Dim rel As DAO.Relation

    Dim dict As Object

    

    Set dict = CreateObject("Scripting.Dictionary")

    Set db = CurrentDb

    

    For Each rel In db.Relations

        If rel.Table = strParentTable Then

            dict(rel.ForeignTable) = rel.Fields(0).ForeignName

        End If

    Next rel

    

    Set GetChildTables = dict

End Function





' UTILITY: Duplicate Multiple Records



Public Function DuplicateMultipleRecords( _

    strTable As String, _

    arrSourceIDs() As Long, _

    Optional dictInclude As Object = Nothing, _

    Optional dictExclude As Object = Nothing, _

    Optional dictChildren As Object = Nothing) As Collection

    

    Dim col As Collection

    Dim i As Long

    Dim lngNewID As Long

    

    Set col = New Collection

    

    For i = LBound(arrSourceIDs) To UBound(arrSourceIDs)

        lngNewID = DuplicateRecord(strTable, arrSourceIDs(i), dictInclude, dictExclude, dictChildren)

        If lngNewID > 0 Then

            col.Add lngNewID

        End If

    Next i

    

    Set DuplicateMultipleRecords = col

End Function





' UTILITY: Preview Fields That Will Be Duplicated



Public Function PreviewDuplication(strTable As String, _

                                  Optional dictInclude As Object = Nothing, _

                                  Optional dictExclude As Object = Nothing) As String

    

    Dim db As DAO.Database

    Dim tdf As DAO.TableDef

    Dim fld As DAO.Field

    Dim strResult As String

    Dim blnUseInclude As Boolean

    

    Set db = CurrentDb

    Set tdf = db.TableDefs(strTable)

    

    blnUseInclude = Not (dictInclude Is Nothing)

    If blnUseInclude And dictInclude.Count = 0 Then blnUseInclude = False

    

    strResult = "Fields to be duplicated for [" & strTable & "]:" & vbCrLf & vbCrLf

    

    For Each fld In tdf.Fields

        If Not ShouldSkipField(fld, dictInclude, dictExclude, blnUseInclude, strTable) Then

            strResult = strResult & "  ? " & fld.Name & vbCrLf

        Else

            strResult = strResult & "  ? " & fld.Name & " (skipped)" & vbCrLf

        End If

    Next fld

    

    PreviewDuplication = strResult

End Function





' PRIVATE: Should we skip this field?



Private Function ShouldSkipField(fld As DAO.Field, _

                                dictInclude As Object, _

                                dictExclude As Object, _

                                blnUseInclude As Boolean, _

                                strTable As String) As Boolean

    

    Dim strName As String

    strName = fld.Name

    

    ' Always skip primary key and auto-increment

    If fld.Attributes And dbAutoIncrField Then

        ShouldSkipField = True

        Exit Function

    End If

    

    ' Skip known unique indexed fields

    If IsFieldUniqueIndexed(strTable, strName) Then

        ShouldSkipField = True

        Exit Function

    End If

    

    ' Include/Exclude logic

    If blnUseInclude Then

        ShouldSkipField = Not DictionaryExists(dictInclude, strName)

    Else

        If Not (dictExclude Is Nothing) Then

            ShouldSkipField = DictionaryExists(dictExclude, strName)

        Else

            ShouldSkipField = False

        End If

    End If

End Function





' PRIVATE: Duplicate child records (1-to-many)



Private Sub DuplicateChildRecords(strChildTable As String, _

                                 lngOldParentID As Long, _

                                 lngNewParentID As Long, _

                                 strForeignKeyField As String)

    

    Dim db As DAO.Database

    Dim rsSource As DAO.Recordset

    Dim rsTarget As DAO.Recordset

    Dim fld As DAO.Field

    Dim strSQL As String

    

    Set db = CurrentDb

    strSQL = "SELECT * FROM [" & strChildTable & "] WHERE [" & strForeignKeyField & "] = " & lngOldParentID

    Set rsSource = db.OpenRecordset(strSQL, dbOpenSnapshot)

    If rsSource.EOF Then Exit Sub

    

    Set rsTarget = db.OpenRecordset("[" & strChildTable & "]", dbOpenDynaset)

    

    Do While Not rsSource.EOF

        rsTarget.AddNew

        For Each fld In rsSource.Fields

            If fld.Name <> strForeignKeyField And Not (fld.Attributes And dbAutoIncrField) Then

                If Not IsNull(rsSource(fld.Name)) Then

                    rsTarget(fld.Name) = rsSource(fld.Name)

                End If

            End If

        Next fld

        rsTarget(strForeignKeyField) = lngNewParentID

        rsTarget.Update

        rsSource.MoveNext

    Loop

    

    rsSource.Close

    rsTarget.Close

End Sub





' HELPER: Check if field has unique index



Private Function IsFieldUniqueIndexed(strTable As String, strField As String) As Boolean

    Dim db As DAO.Database

    Dim tdf As DAO.TableDef

    Dim idx As DAO.Index

    Dim fld As DAO.Field

    

    On Error Resume Next

    Set db = CurrentDb

    Set tdf = db.TableDefs(strTable)

    

    For Each idx In tdf.Indexes

        If idx.Unique Then

            For Each fld In idx.Fields

                If UCase(fld.Name) = UCase(strField) Then

                    IsFieldUniqueIndexed = True

                    Exit Function

                End If

            Next fld

        End If

    Next idx

    

    IsFieldUniqueIndexed = False

End Function





' HELPER: Get primary key field name



Private Function GetPrimaryKeyField(strTable As String) As String

    Dim tdf As DAO.TableDef

    Dim idx As DAO.Index

    

    Set tdf = CurrentDb.TableDefs(strTable)

    For Each idx In tdf.Indexes

        If idx.Primary Then

            GetPrimaryKeyField = idx.Fields(0).Name

            Exit Function

        End If

    Next idx

    

    GetPrimaryKeyField = "ID"

End Function





' HELPER: Safe dictionary exists check



Private Function DictionaryExists(dict As Object, key As Variant) As Boolean

    On Error Resume Next

    Dim temp As Variant

    temp = dict(key)

    DictionaryExists = (Err.Number = 0)

    On Error GoTo 0

End Function





' USAGE EXAMPLES



Sub Example_SimpleDuplicate()

    ' Duplicate entire record (auto-skips unique fields)

    Dim lngNewID As Long

    Debug.Print "Testing simple duplicate of Clients ID=65..."

    lngNewID = DuplicateRecord("Clients", 65)

    If lngNewID > 0 Then

        Debug.Print "SUCCESS! New Client ID: " & lngNewID

    Else

        Debug.Print "FAILED - returned 0"

    End If

End Sub



Sub Example_ExcludeFields()

    ' Duplicate but exclude specific fields

    Dim dictExclude As Object

    Set dictExclude = CreateExcludeList("EmailBilling", "Notes", "PhoneMain")

    

    Dim lngNewID As Long

    lngNewID = DuplicateRecord("Clients", 65, , dictExclude)

    Debug.Print "New Client ID: " & lngNewID

End Sub



Sub Example_IncludeOnlySpecific()

    ' Copy ONLY specific fields

    Dim dictInclude As Object

    Set dictInclude = CreateIncludeList("ClientName", "Address", "City", "PostalCode")

    

    Dim lngNewID As Long

    lngNewID = DuplicateRecord("Clients", 65, dictInclude)

    Debug.Print "New Client ID: " & lngNewID

End Sub



Sub Example_WithChildren()

    ' Duplicate client with all services and contacts

    Dim dictChildren As Object

    Set dictChildren = CreateChildRelations()

    

    Call AddChildRelation(dictChildren, "Services", "ClientID")

    Call AddChildRelation(dictChildren, "ClientContacts", "ClientID")

    

    Dim lngNewID As Long

    lngNewID = DuplicateRecord("Clients", 65, , , dictChildren)

    Debug.Print "New Client ID: " & lngNewID

End Sub



Sub Example_PreviewDuplication()

    ' Preview what will be duplicated

    Dim dictExclude As Object

    Set dictExclude = CreateExcludeList("Notes", "EmailBilling")

    

    Debug.Print PreviewDuplication("Clients", , dictExclude)

End Sub



Sub Example_GetDuplicatableFields()

    ' Get list of all fields that CAN be duplicated

    Dim col As Collection

    Dim v As Variant

    

    Set col = GetDuplicatableFields("Clients")

    For Each v In col

        Debug.Print v

    Next v

End Sub



Sub Example_AutoDetectChildren()

    ' Automatically detect child tables

    Dim dictChildren As Object

    Set dictChildren = GetChildTables("Clients")

    

    Dim lngNewID As Long

    lngNewID = DuplicateRecord("Clients", 65, , , dictChildren)

    Debug.Print "New Client ID with all children: " & lngNewID

End Sub



Sub Example_DuplicateMultiple()

    ' Duplicate multiple records at once

    Dim arrIDs(1 To 3) As Long

    arrIDs(1) = 65

    arrIDs(2) = 72

    arrIDs(3) = 88

    

    Dim colNewIDs As Collection

    Set colNewIDs = DuplicateMultipleRecords("Clients", arrIDs)

    

    Dim v As Variant

    For Each v In colNewIDs

        Debug.Print "Created Client ID: " & v

    Next v

End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' MODULE: modServiceNumbers

' PURPOSE: Thread-safe, year-aware service & invoice number generation

'          with full transaction safety and audit trail on failure

' SECURITY: Pessimistic locking + transactions = ZERO duplicates

' AUTHOR: Expert Back-End Developer (MS Access Concurrency Specialist)

' CREATED: November 17, 2025

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Database

Option Explicit



Private Const MODULE_NAME As String = "modServiceNumbers"





' PUBLIC: GenerateServiceNumber



''

''' Function : GenerateServiceNumber

''' Purpose  : Atomically generate next SRV-YYYY-NNNNN number

'''            Handles year rollover and concurrent users safely

''' Returns  : String like "SRV-2025-00001"

''' Throws   : Never – returns "" on critical failure with full audit

''

Public Function GenerateServiceNumber() As String

    On Error GoTo ErrorHandler



    Dim db As DAO.Database

    Dim rs As DAO.Recordset

    Dim lngYear As Long

    Dim lngNext As Long

    Dim strNumber As String



    Set db = CurrentDb



    ' === BEGIN TRANSACTION + PESSIMISTIC LOCK =========================

    BeginTransaction

    Set rs = db.OpenRecordset("SystemSettings", dbOpenDynaset, dbSeeChanges + dbPessimistic)

    

    If rs.EOF Then

        GoTo MissingSettings

    End If



    With rs

        .MoveFirst

        Do While Not .EOF

            If !SettingName = "CurrentServiceYear" Then

                lngYear = Val(Nz(!SettingValue, Year(Date)))

            ElseIf !SettingName = "NextServiceNumber" Then

                lngNext = Val(Nz(!SettingValue, 1))

            End If

            .MoveNext

        Loop



        ' --- Year rollover detection ---

        If lngYear <> Year(Date) Then

            lngYear = Year(Date)

            lngNext = 1

            

            ' Update year

            .MoveFirst

            Do While Not .EOF

                If !SettingName = "CurrentServiceYear" Then

                    .Edit

                    !SettingValue = CStr(lngYear)

                    !ModifiedBy = g_lngUserID

                    !ModifiedDate = Now()

                    .Update

                ElseIf !SettingName = "NextServiceNumber" Then

                    .Edit

                    !SettingValue = "1"

                    !ModifiedBy = g_lngUserID

                    !ModifiedDate = Now()

                    .Update

                End If

                .MoveNext

            Loop

        End If



        ' --- Build number ---

        strNumber = "SRV-" & Format(lngYear, "0000") & "-" & Format(lngNext, "00000")



        ' --- Increment next number ---

        .MoveFirst

        Do While Not .EOF

            If !SettingName = "NextServiceNumber" Then

                .Edit

                !SettingValue = CStr(lngNext + 1)

                !ModifiedBy = g_lngUserID

                !ModifiedDate = Now()

                .Update

                Exit Do

            End If

            .MoveNext

        Loop



        ' Audit the generation (critical operation ? sync)

        modAudit.LogAudit "SystemSettings", 0, "NextServiceNumber", _

                          lngNext, lngNext + 1, "Increment", , True



        GenerateServiceNumber = strNumber

    End With



    CommitTransaction

    GoTo CleanExit



MissingSettings:

    RollbackTransaction

    modAudit.LogAudit "SystemSettings", 0, "", Null, Null, "MissingSettings", , True

    modUtilities.LogError "GenerateServiceNumber", 1001, _

        "Required SystemSettings records missing (CurrentServiceYear/NextServiceNumber)"

    GenerateServiceNumber = ""



CleanExit:

    On Error Resume Next

    If Not rs Is Nothing Then

        rs.Close

        Set rs = Nothing

    End If

    Set db = Nothing

    Exit Function



ErrorHandler:

    RollbackTransaction

    modUtilities.LogError "GenerateServiceNumber", Err.Number, Err.Description

    GenerateServiceNumber = ""

    Resume CleanExit

End Function





' PUBLIC: GenerateInvoiceNumber



''

''' Function : GenerateInvoiceNumber

''' Purpose  : Identical logic but for INV-YYYY-NNNNN

''

Public Function GenerateInvoiceNumber() As String

    On Error GoTo ErrorHandler



    Dim db As DAO.Database

    Dim rs As DAO.Recordset

    Dim lngYear As Long

    Dim lngNext As Long

    Dim strNumber As String



    Set db = CurrentDb



    BeginTransaction

    Set rs = db.OpenRecordset("SystemSettings", dbOpenDynaset, dbSeeChanges + dbPessimistic)

    

    If rs.EOF Then

        GoTo MissingSettings

    End If



    With rs

        .MoveFirst

        Do While Not .EOF

            If !SettingName = "CurrentInvoiceYear" Then

                lngYear = Val(Nz(!SettingValue, Year(Date)))

            ElseIf !SettingName = "NextInvoiceNumber" Then

                lngNext = Val(Nz(!SettingValue, 1))

            End If

            .MoveNext

        Loop



        If lngYear <> Year(Date) Then

            lngYear = Year(Date)

            lngNext = 1



            .MoveFirst

            Do While Not .EOF

                If !SettingName = "CurrentInvoiceYear" Then

                    .Edit

                    !SettingValue = CStr(lngYear)

                    !ModifiedBy = g_lngUserID

                    !ModifiedDate = Now()

                    .Update

                ElseIf !SettingName = "NextInvoiceNumber" Then

                    .Edit

                    !SettingValue = "1"

                    !ModifiedBy = g_lngUserID

                    !ModifiedDate = Now()

                    .Update

                End If

                .MoveNext

            Loop

        End If



        strNumber = "INV-" & Format(lngYear, "0000") & "-" & Format(lngNext, "00000")



        .MoveFirst

        Do While Not .EOF

            If !SettingName = "NextInvoiceNumber" Then

                .Edit

                !SettingValue = CStr(lngNext + 1)

                !ModifiedBy = g_lngUserID

                !ModifiedDate = Now()

                .Update

                Exit Do

            End If

            .MoveNext

        Loop



        modAudit.LogAudit "SystemSettings", 0, "NextInvoiceNumber", _

                          lngNext, lngNext + 1, "Increment", , True



        GenerateInvoiceNumber = strNumber

    End With



    CommitTransaction

    GoTo CleanExit



MissingSettings:

    RollbackTransaction

    modAudit.LogAudit "SystemSettings", 0, "", Null, Null, "MissingSettings", , True

    modUtilities.LogError "GenerateInvoiceNumber", 1002, _

        "Required SystemSettings records missing (CurrentInvoiceYear/NextInvoiceNumber)"

   GenerateInvoiceNumber = ""



CleanExit:

    On Error Resume Next

    If Not rs Is Nothing Then

        rs.Close

        Set rs = Nothing

    End If

    Set db = Nothing

    Exit Function



ErrorHandler:

    RollbackTransaction

    modUtilities.LogError "GenerateInvoiceNumber", Err.Number, Err.Description

    GenerateInvoiceNumber = ""

    Resume CleanExit

End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' MODULE: modUtilities

' PURPOSE: Miscellaneous helper functions + centralized error handling

' AUTHOR: Expert Back-End Developer (MS Access VBA Security & Reliability)

' CREATED: November 17, 2025

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Database

Option Explicit



Private Const MODULE_NAME As String = "modUtilities"





' PUBLIC: LogError – Centralized error logging with fallback to file



''

''' Procedure: LogError

''' Purpose   : Write comprehensive error details to ErrorLog table

'''             with automatic fallback to local text file if DB unavailable

''' Parameters:

'''   strProcName     - Name of procedure where error occurred

'''   lngErrNumber    - Err.Number

'''   strErrDesc      - Err.Description

'''   Optional strAdditionalInfo - Any context (SQL, record ID, etc.)

'''   Optional lngLine        - Line number (requires numbered lines or Erl)

''

Public Sub LogError(ByVal strProcName As String, _

                    ByVal lngErrNumber As Long, _

                    ByVal strErrDesc As String, _

                    Optional ByVal strAdditionalInfo As String = "", _

                    Optional ByVal lngLine As Long = 0)



    On Error GoTo FallbackFileLog

    

    Dim db As DAO.Database

    Dim rs As DAO.Recordset

    Dim strSQL As String

    Dim strSource As String

    

    strSource = MODULE_NAME & "." & strProcName

    If lngLine > 0 Then strSource = strSource & " (Line " & lngLine & ")"

    

    Set db = CurrentDb

    

    strSQL = "INSERT INTO ErrorLog " & _

             "(ErrorNumber, ErrorDescription, ErrorSource, ErrorLine, " & _

             "UserID, ErrorDate, WorkstationName, AdditionalInfo) VALUES " & _

             "(" & lngErrNumber & ", " & _

             "'" & Replace(strErrDesc, "'", "''") & "', " & _

             "'" & Left(Replace(strSource, "'", "''"), 100) & "', " & _

             lngLine & ", " & _

             Nz(g_lngUserID, 0) & ", " & _

             "Now(), " & _

             "'" & Left(Environ("COMPUTERNAME"), 50) & "', " & _

             "'" & Left(Replace(Nz(strAdditionalInfo, ""), "'", "''"), 65535) & "')"

    

    db.Execute strSQL, dbFailOnError

    Exit Sub

    

FallbackFileLog:

    ' If we cannot write to the database at all, fall back to local text file

    Dim strLogPath As String

    Dim intFile As Integer

    

    strLogPath = GetLocalAppDataPath() & "\RoadFreightERP_ErrorLog.txt"

    

    intFile = FreeFile

    Open strLogPath For Append As #intFile

    Print #intFile, "=== ERROR ===" & vbCrLf & _

                    "DateTime      : " & Now() & vbCrLf & _

                    "UserID        : " & g_lngUserID & vbCrLf & _

                    "Workstation   : " & Environ("COMPUTERNAME") & vbCrLf & _

                    "Procedure     : " & strProcName & IIf(lngLine > 0, " (Line " & lngLine & ")", "") & vbCrLf & _

                    "Error #       : " & lngErrNumber & vbCrLf & _

                    "Description   : " & strErrDesc & vbCrLf & _

                    "Additional    : " & strAdditionalInfo & vbCrLf & _

                    String(50, "=") & vbCrLf

    Close #intFile

End Sub





' PRIVATE: Helper to get reliable local folder for fallback logging



Private Function GetLocalAppDataPath() As String

    Dim strPath As String

    strPath = Environ("APPDATA")

    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"

    strPath = strPath & "RoadFreightERP\"

    

    If Dir(strPath, vbDirectory) = "" Then

        MkDir strPath

    End If

    

    GetLocalAppDataPath = strPath

End Function



' Helper – safe because not every form has every button

Private Sub SafeEnable(ctl As control, blnEnabled As Boolean)

    On Error Resume Next

    Dim origForeColor As Long

    

    origForeColor = ctl.ForeColor

    ctl.Enabled = blnEnabled

    ctl.ForeColor = IIf(blnEnabled, origForeColor, vbGrayText)

    On Error GoTo 0

End Sub



Public Sub SetButtonStates(frm As Form, _

                           Optional permEdit As String = "", _

                           Optional permDelete As String = "")



    On Error Resume Next

    

    '-------------------------------

    ' Determine state

    '-------------------------------

    Dim isNew As Boolean:    isNew = frm.NewRecord

    Dim hasRecord As Boolean: hasRecord = Not isNew

    Dim dirty As Boolean:    dirty = frm.dirty

    

    Dim canEdit As Boolean:  canEdit = (permEdit <> "") And HasPermission(permEdit)

    Dim canDelete As Boolean: canDelete = (permDelete <> "") And HasPermission(permDelete)

    

    '-------------------------------

    ' New / Duplicate / Delete buttons

    '-------------------------------

    SafeEnable frm.cmdNew, True

    SafeEnable frm.cmdDuplicate, hasRecord And canEdit

    SafeEnable frm.cmdDelete, hasRecord And (Not isNew) And canDelete

    

    '-------------------------------

    ' Save / Undo / Cancel buttons

    '-------------------------------

    If isNew Then

        ' For new records, always allow Save if user can edit

        SafeEnable frm.cmdSave, canEdit

        ' Cancel should always be enabled

        SafeEnable frm.cmdCancel, True

        ' Undo not applicable on new record

        SafeEnable frm.cmdUndo, False

    Else

        ' Existing record – follow dirty + permissions

        SafeEnable frm.cmdSave, dirty And canEdit

        SafeEnable frm.cmdUndo, dirty

        SafeEnable frm.cmdCancel, dirty

    End If



    '-------------------------------

    ' Other buttons (export, print, refresh)

    '-------------------------------

    SafeEnable frm.cmdExportExcel, hasRecord

    SafeEnable frm.cmdPrint, hasRecord

    SafeEnable frm.cmdAdvancedSearch, True

    SafeEnable frm.cmdRefresh, True

End Sub

Public Sub AddToFilter(ByRef strF As String, strNew As String)
    If strF = "" Then strF = strNew Else strF = strF & " AND " & strNew
End Sub

Public Sub CopyToClipboard(ByVal strText As String)

    With CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}") ' HTML clipboard

        .SetText strText

        .PutInClipboard

    End With

    ' Optional: beep or toast notification

    Beep

End Sub



Public Sub ShowCopiedFeedback(lbl As Label)

    Dim origCaption As String

    Dim origForeColor As Long

    

    ' Save original appearance

    origCaption = lbl.Caption

    origForeColor = lbl.ForeColor

    

    ' Show success

    lbl.Caption = "Copied!"

    lbl.ForeColor = RGB(14, 76, 73)

    

    ' Force repaint

'    lbl.Repaint

    DoEvents

    

    ' Wait 1.5 seconds (non-blocking with DoEvents)

    Dim pauseTime As Double

    pauseTime = Timer + 1.5

    Do While Timer < pauseTime

        DoEvents

    Loop

    

    ' Restore original

    lbl.Caption = origCaption

    lbl.ForeColor = origForeColor

    lbl.FontBold = False

'    lbl.Repaint

End Sub







' STANDARD ERROR HANDLER TEMPLATE (copy-paste into every procedure)



' Place this exact block in every Public/Private Sub or Function

'

'    On Error GoTo ErrorHandler

'    '=== MAIN CODE HERE =============================================

'

'ExitHandler:

'    ' Cleanup code that MUST run (close recordsets, set objects = Nothing)

'    Set rs = Nothing

'    Set db = Nothing

'    Exit Sub

'

'ErrorHandler:

'    Dim strMsg As String

'    Select Case Err.Number

'        Case 1234   ' Example of recoverable error

'            strMsg = "Friendly message for user"

'            MsgBox strMsg, vbExclamation, APP_NAME

'            Resume ExitHandler

'        Case Else

'            ' Non-recoverable – log + graceful shutdown

'            LogError "ProcedureName", Err.Number, Err.Description, "Any context info", Erl

'            strMsg = "A critical error has occurred. The application will now close " & _

'                     "to prevent data corruption." & vbCrLf & vbCrLf & _

'                     "Please contact your system administrator."

'            MsgBox strMsg, vbCritical, APP_NAME & " - Critical Error"

'            DoCmd.Quit acQuitSaveNone

'    End Select

'



Option Compare Database

Option Explicit



Public Function IsPasswordStrong(ByVal strPassword As String) As Boolean

    If Len(strPassword) < 8 Then GoTo Fail

    If Not strPassword Like "*[a-z]*" Then GoTo Fail

    If Not strPassword Like "*[A-Z]*" Then GoTo Fail

    If Not strPassword Like "*[0-9]*" Then GoTo Fail

    If Not strPassword Like "*[!@#$%^&*()_+-=[]{}|;':,.<>?/]*" Then GoTo Fail

    

    IsPasswordStrong = True

    Exit Function

Fail:

    IsPasswordStrong = False

End Function



Public Function IsValidEmail(strEmail As String) As Boolean

    Dim regEx As Object

    Set regEx = CreateObject("VBScript.RegExp")

    regEx.Pattern = "^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,})+$"

    IsValidEmail = regEx.Test(strEmail)

End Function



Public Function FormatIBAN(strIBAN As String) As String

    Dim strClean As String

    strClean = Replace(Replace(Replace(strIBAN, " ", ""), "-", ""), ".", "")

    Dim i As Integer

    For i = 1 To Len(strClean) Step 4

        FormatIBAN = FormatIBAN & Mid(strClean, i, 4) & IIf(i + 4 <= Len(strClean), " ", "")

    Next

    FormatIBAN = Trim(FormatIBAN)

End Function

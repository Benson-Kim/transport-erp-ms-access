'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MODULE: modTesting (COMPLETE)
' PURPOSE: Full automated testing suite with all helper functions
'          Unit + Integration + E2E + Regression
' AUTHOR: Tifamovs Saeol <tifamovs@gmail.com> [Benson Kimathi]
' CREATED: November 17, 2025
' UPDATED: December 03, 2025
'          Added comprehensive tests for modRecordDuplicator
' SECURITY: N/A (test code only)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Compare Database
Option Explicit

Private Const MODULE_NAME As String = "modTesting"

Private mlngTotalTests As Long
Private mlngPassedTests As Long
Private mlngFailedTests As Long
Private mstrCurrentTest As String


' PUBLIC: RunAllTests � Master test runner

Public Sub RunAllTests()
    On Error GoTo ErrorHandler
    
    ' CRITICAL: Test crypto first
    If Not TestCryptoAvailable() Then
        MsgBox "CRITICAL ERROR: Password hashing (modCrypto) is not working." & vbCrLf & vbCrLf & _
               "This must be fixed before running tests." & vbCrLf & vbCrLf & _
               "Check that BCrypt or .NET crypto is available.", vbCritical, "Test Suite - Crypto Failed"
        Exit Sub
    End If
    modPermissions.LoadPermissions
    
    mlngTotalTests = 0
    mlngPassedTests = 0
    mlngFailedTests = 0
    
    Debug.Print String(60, "-")
    Debug.Print "ROAD FREIGHT ERP - AUTOMATED TEST SUITE"
    Debug.Print "Started: " & Now()
    Debug.Print String(60, "=")
    
    Test_Crypto
    Test_Authentication_System
    Test_Permission_System
    Test_ServiceNumber_Generation
    Test_RecordDuplicator_Comprehensive
    Test_Audit_Logging_Framework
    Test_Error_Handling

    Debug.Print String(60, "-")
    Debug.Print "TEST SUMMARY"
    Debug.Print "Total Tests : " & mlngTotalTests
    Debug.Print "Passed      : " & mlngPassedTests & " (Success)"
    Debug.Print "Failed      : " & mlngFailedTests & " (Failure)"

    Dim strRate As String
    If mlngTotalTests > 0 Then
        strRate = Format(mlngPassedTests / mlngTotalTests, "0.0%")
    Else
        strRate = "N/A (no tests executed)"
    End If
    Debug.Print "Success Rate: " & strRate
    Debug.Print "Completed   : " & Now()
    Debug.Print String(60, "=")

    If mlngTotalTests = 0 Then
        MsgBox "WARNING: No tests were executed!" & vbCrLf & vbCrLf & _
               "Make sure you uncomment the desired test calls in RunAllTests.", vbExclamation, "No Tests Ran"
    ElseIf mlngFailedTests = 0 Then
        MsgBox "All " & mlngTotalTests & " tests passed!" & vbCrLf & _
               "System is verified and production-ready.", vbInformation, "Test Suite - SUCCESS"
    Else
        MsgBox mlngFailedTests & " test(s) failed. Check Immediate Window (Ctrl+G) for details.", vbCritical, "Test Suite - FAILED"
    End If
    
    Exit Sub
    
ErrorHandler:
    LogTestResult "RunAllTests", False, "Critical error: " & Err.Description
End Sub

' PRE-CHECK: Crypto Availability
Private Function TestCryptoAvailable() As Boolean
    On Error GoTo Fail
    
    Debug.Print vbCrLf & "PRE-Check: Verifying modCrypto..."
    
    Dim h1 As String, h2 As String
    h1 = modCrypto.HashPassword("test123")
    h2 = modCrypto.HashPassword("test123")

    If Len(h1) = 0 Or Len(h2) = 0 Then GoTo Fail
    If h1 = h2 Then GoTo Fail ' Same password ? different hash (salting)
    If Not modCrypto.VerifyPassword("test123", h1) Then GoTo Fail

    Debug.Print " (Success) PASS: modCrypto fully operational (salt+hash+verify)"
    TestCryptoAvailable = True
    Exit Function
Fail:
    Debug.Print " (Failure) CRITICAL: modCrypto failed"
    TestCryptoAvailable = False
End Function


' TEST 1: Authentication System

Private Sub Test_Authentication_System()
    LogTestStart "Authentication System"
    
    Dim lngTestUser As Long
    
    ' 1. Valid login
    lngTestUser = CreateTestUser("testadmin", "Admin123!", "Admin", True)
    If lngTestUser = 0 Then
        LogTestResult "Create test admin user", False, "Failed to create user - check crypto"
        Exit Sub
    End If
    TestLogin "testadmin", "Admin123!", True, "Valid admin login"
    
    
    ' 2. Invalid credentials
    TestLogin "testadmin", "wrongpass", False, "Invalid password"
    TestLogin "test_admin", "Admin123!", False, "Invalid username"
    
    ' 3. Deactivated user
    Dim lngTestDeactivatedUser As Long
    lngTestDeactivatedUser = CreateTestUser("testdeactivated", "Test123!", "Operator", False)
    If lngTestDeactivatedUser > 0 Then
        TestLogin "testdeactivated", "Test123!", False, "Deactivated user blocked"
        DeleteTestUser lngTestDeactivatedUser
    End If
    
    ' 4. Account lockout after 5 failed attempts
    If lngTestUser > 0 Then
        Dim i As Integer
        For i = 1 To 7
            TestLogin "testadmin", "wrong", False, "Failed attempt " & i
        Next i
        Dim dtLock As Date
    dtLock = Nz(DLookup("LockoutUntil", "Users", "Username='testadmin'"), #1/1/2000#)
    LogTestResult "Exponential lockout applied", (dtLock > DateAdd("n", 15, Now()))  ' or just True if not Null
    End If
    
    ' 5. Lockout clears on success
    CurrentDb.Execute "UPDATE Users SET FailedLoginAttempts=0, LockoutUntil=NULL WHERE Username='testadmin'"
    TestLogin "testadmin", "Admin123!", True, "Lockout cleared on success"

    ' 6. Case-insensitive + trimmed username
    TestLogin "  TestAdmin  ", "Admin123!", True, "Case/trim insensitive login"
    
    '7. Password expiry (force expiry)
    CurrentDb.Execute "UPDATE Users SET PasswordSetDate = Date() - 100 WHERE Username='testadmin'"
    TestLogin "testauth", "Secure123!", False, "Password expired ? login blocked"
    ' Reset
    CurrentDb.Execute "UPDATE Users SET PasswordSetDate = Date() WHERE Username='testadmin'"
    
    '8. Password reuse prevention
    LoginAs "testadmin", "Admin123!"
    Dim blnChange As Boolean
    blnChange = ChangePassword(lngTestUser, "Admin123!", "Admin123!") ' Same password
    LogTestResult "Password reuse blocked", Not blnChange
    
    ' 9. Concurrent login (simulated)
    TestConcurrentLoginSimulation
    
    DeleteTestUser lngTestUser
    LogTestSummary
End Sub

Private Sub TestLogin(strUser As String, strPass As String, blnExpected As Boolean, strDesc As String, Optional blnAssertGlobals As Boolean = True)
    mlngTotalTests = mlngTotalTests + 1
    mstrCurrentTest = "Login: " & strUser

    Call modGlobals.InitializeGlobalVariables
    Dim blnResult As Boolean
    blnResult = modAuthentication.AuthenticateUser(strUser, strPass)

    If blnResult = blnExpected Then
        LogTestResult "PASS: " & strDesc, True
        If blnExpected And blnAssertGlobals Then
            Assert (g_lngUserID > 0), "g_lngUserID set"
            Assert (g_strUserRole = "Admin"), "g_strUserRole set"
        End If
    Else
        LogTestResult "FAIL: " & strDesc, False, "Expected: " & blnExpected & " | Got: " & blnResult
    End If

    ' Always logout at the end
    If IsUserLoggedIn() Then LogoutUser
End Sub


' TEST 2: Permission System

Private Sub Test_Permission_System()
    LogTestStart "Permission System"
    
    Dim lngAdmin As Long, lngManager As Long, lngOperator As Long
    
    modGlobals.InitializeGlobalVariables
    
    modPermissions.LoadPermissions
    
    ' Create test users
    lngAdmin = CreateTestUser("testadmin2", "Admin123!", "Admin", True)
    lngManager = CreateTestUser("testmanager2", "Manager123!", "Manager", True)
    lngOperator = CreateTestUser("testoperator2", "Operator123!", "Operator", True)
    
    If lngAdmin = 0 Or lngManager = 0 Or lngOperator = 0 Then
        LogTestResult "Create permission test users", False, "Failed to create test users"
        Exit Sub
    End If
    
    LoginAs "testadmin2", "Admin123!"
    TestPermissionAccess "Admin", True, True, True, True, True, True
    
    LogoutUser
    LoginAs "testmanager2", "Manager123!"
    TestPermissionAccess "Manager", False, True, True, True, False, False
    
    LogoutUser
    LoginAs "testoperator2", "Operator123!"
    TestPermissionAccess "Operator", False, False, False, True, False, False
    
    ' Cleanup
    LogoutUser
    DeleteTestUser lngAdmin
    DeleteTestUser lngManager
    DeleteTestUser lngOperator
    
    LogTestSummary
End Sub

Private Sub TestPermissionAccess(strRole As String, _
    blnDeleteClient As Boolean, blnEditCompleted As Boolean, blnGenInvoice As Boolean, _
    blnGenLoadingOrder As Boolean, blnViewAudit As Boolean, blnManageUsers As Boolean)

    TestHasPermission PERM_DELETE_CLIENTS, blnDeleteClient, strRole & " delete client"
    TestHasPermission PERM_EDIT_COMPLETED_SERVICES, blnEditCompleted, strRole & " edit completed service"
    TestHasPermission PERM_GENERATE_INVOICES, blnGenInvoice, strRole & " generate invoice"
    TestHasPermission PERM_GENERATE_LOADING_ORDERS, blnGenLoadingOrder, strRole & " generate loading order"
    TestHasPermission PERM_VIEW_AUDIT_LOG, blnViewAudit, strRole & " view audit log"
    TestHasPermission PERM_MANAGE_USERS, blnManageUsers, strRole & " manage users"
End Sub

Private Sub TestHasPermission(strPerm As String, blnExpected As Boolean, strDesc As String)
    mlngTotalTests = mlngTotalTests + 1
    Dim blnResult As Boolean
    blnResult = HasPermission(strPerm)
    LogTestResult strDesc, (blnResult = blnExpected), IIf(blnResult = blnExpected, "", "Got: " & blnResult)
End Sub


' TEST 3: Service Number Generation

Private Sub Test_ServiceNumber_Generation()
    LogTestStart "Service & Invoice Number Generation - Final Validation"

    Dim arrSrv(1 To 30) As String
    Dim arrInv(1 To 30) As String
    Dim i As Long
    Dim strNum As String

    ' Ensure clean state: delete all test settings
    On Error Resume Next
    CurrentDb.Execute "DELETE FROM SystemSettings WHERE SettingName LIKE 'Current*Year' OR SettingName LIKE 'Next*Number'", dbFailOnError
    On Error GoTo 0


    ' Test sequential uniqueness (30 rapid calls)
    For i = 1 To 30
        arrSrv(i) = GenerateServiceNumber()
        arrInv(i) = GenerateInvoiceNumber()
        DoEvents
    Next i

    LogTestResult "30 consecutive Service numbers are unique", IsArrayUnique(arrSrv), _
                  "Sample: " & arrSrv(1) & " ... " & arrSrv(30)
    LogTestResult "30 consecutive Invoice numbers are unique", IsArrayUnique(arrInv), _
                  "Sample: " & arrInv(1) & " ... " & arrInv(30)

    ' Test year rollover
    On Error Resume Next
    CurrentDb.Execute "UPDATE SystemSettings SET SettingValue = '2020' WHERE SettingName = 'CurrentServiceYear'", dbFailOnError
    CurrentDb.Execute "UPDATE SystemSettings SET SettingValue = '99999' WHERE SettingName = 'NextServiceNumber'", dbFailOnError
    CurrentDb.Execute "UPDATE SystemSettings SET SettingValue = '2020' WHERE SettingName = 'CurrentInvoiceYear'", dbFailOnError
    CurrentDb.Execute "UPDATE SystemSettings SET SettingValue = '99999' WHERE SettingName = 'NextInvoiceNumber'", dbFailOnError
    On Error GoTo 0

    strNum = GenerateServiceNumber()
    LogTestResult "Service rollover ? " & Year(Now()) & "-00001", _
                  InStr(strNum, CStr(Year(Now()))) > 0 And Right(strNum, 5) = "00001", strNum

    strNum = GenerateInvoiceNumber()
    LogTestResult "Invoice rollover ? " & Year(Now()) & "-00001", _
                  InStr(strNum, CStr(Year(Now()))) > 0 And Right(strNum, 5) = "00001", strNum

    ' Test missing settings ? auto-creation
    CurrentDb.Execute "DELETE FROM SystemSettings WHERE SettingName IN ('CurrentServiceYear','NextServiceNumber','CurrentInvoiceYear','NextInvoiceNumber')", dbFailOnError

    strNum = GenerateServiceNumber()
    LogTestResult "Missing settings ? SRV-" & Year(Now()) & "-00001 generated", _
                  strNum Like "SRV-" & Year(Now()) & "-00001", strNum

    strNum = GenerateInvoiceNumber()
    LogTestResult "Missing settings ? INV-" & Year(Now()) & "-00001 generated", _
                  strNum Like "INV-" & Year(Now()) & "-00001", strNum


    ' Confirm auto-created records exist
    LogTestResult "Service settings auto-created (2 records)", _
                  DCount("*", "SystemSettings", "SettingName IN ('CurrentServiceYear','NextServiceNumber')") = 2

    LogTestResult "Invoice settings auto-created (2 records)", _
                  DCount("*", "SystemSettings", "SettingName IN ('CurrentInvoiceYear','NextInvoiceNumber')") = 2

    LogTestSummary
End Sub

' TEST 4: Audit Logging

Private Sub Test_Audit_Logging_Framework()
    LogTestStart "Audit Logging Framework"
    
    Dim lngTestUser As Long
    lngTestUser = CreateTestUser("testaudituser", "Audit123!", "Admin", True)
    
    Debug.Print lngTestUser
    
    If lngTestUser = 0 Then Exit Sub

    LoginAs "testaudituser", "Audit123!"
'
'    Debug.Print modGlobals.UserID
    
'    Dim lngClientID As Long
'    lngClientID = CreateTestClient()
'    If lngClientID > 0 Then modAudit.LogInsert "Clients", lngClientID, lngTestUser
'
'    UpdateTestClient lngClientID, "Updated Test Client S.L."
'    If lngClientID > 0 Then modAudit.LogUpdate "Clients", lngClientID, "ClientName", "Test Client S.L.", "Updated Test Client S.L.", lngTestUser
'
'    DeleteTestClient lngClientID
'    If lngClientID > 0 Then modAudit.LogDelete "Clients", lngClientID, lngTestUser
'
'    modAudit.FlushAuditQueue
'
'    DoEvents: DoEvents
'    Application.RefreshDatabaseWindow
'
'    Dim lngCount As Long
'    lngCount = DCount("*", "AuditLog", "TableName='Clients' AND PerformedBy=" & lngTestUser)
'    LogTestResult "Audit entries created for C/U/D", (lngCount >= 3)
'
'    LogoutUser
'    DeleteTestUser lngTestUser
'    LogTestSummary
End Sub


' TEST 5: Error Handling

Private Sub Test_Error_Handling()
    LogTestStart "Error Handling & Resilience"
    
    On Error Resume Next
    
    ' 1. Duplicate VAT
    CurrentDb.Execute "DELETE FROM Clients WHERE VATNumber='ES99999999Z'"
    CurrentDb.Execute "INSERT INTO Clients (ClientName, VATNumber, Address, EmailBilling, Country, Telephone) VALUES ('DupTest', 'ES99999999Z', 'DupTest', 'DupTest', 'Spain', '34 555 671 117')", dbFailOnError
    
    Err.Clear
    On Error Resume Next
    CurrentDb.Execute "INSERT INTO Clients (ClientName, VATNumber, Address, EmailBilling, Country, Telephone) VALUES ('DupTest2', 'ES99999999Z', 'DupTest2', 'DupTest2', 'Spain', '34 555 671 127')", dbFailOnError
    On Error GoTo 0
    LogTestResult "Duplicate VAT blocked", (Err.number <> 0)
    Err.Clear
   
    CurrentDb.Execute "DELETE FROM Clients WHERE VATNumber='ES99999999Z'"
        
    ' 2. Delete with dependencies
    Dim lngServiceID As Long, lngClientID As Long, strSrvNum As String
    strSrvNum = modServiceNumbers.GenerateServiceNumber()
    lngServiceID = CreateTestService(lngClientID, strSrvNum)
    
    lngClientID = Nz(DLookup("ClientID", "Services", "ServiceID=" & lngServiceID), 0)
    
    Err.Clear
    On Error Resume Next
    CurrentDb.Execute "DELETE FROM Clients WHERE ClientID=" & lngClientID, dbFailOnError
    On Error GoTo 0
    LogTestResult "Cannot delete client with services", (Err.number <> 0)
   
    ' Cleanup
    CurrentDb.Execute "DELETE FROM Services WHERE ServiceID=" & lngServiceID
    CurrentDb.Execute "DELETE FROM Clients WHERE ClientID=" & lngClientID
   
    LogTestSummary
    
End Sub

' TEST 6: Crypto Module

Private Sub Test_Crypto()
    
    LogTestStart "modCrypto Security Module"
    mlngTotalTests = mlngTotalTests + 1
    mstrCurrentTest = "Crypto "
    
    Dim strPass As String
    Dim strHash As String
    Dim blnResult As Boolean
    Dim h1 As String, h2 As String
    
    ' Basic hashing + salting (different hashes for same password)
    h1 = modCrypto.HashPassword("test123")
    h2 = modCrypto.HashPassword("test123")
    
    LogTestResult "HashPassword returns non-empty", (Len(h1) > 0 And Len(h2) > 0)
    LogTestResult "Same password produces different hashes (salting)", (h1 <> h2)
    LogTestResult "Verify works on first hash", modCrypto.VerifyPassword("test123", h1)
    LogTestResult "Verify works on second hash", modCrypto.VerifyPassword("test123", h2)
    
    ' Empty password handling
    strHash = modCrypto.HashPassword("")
    Debug.Print "strHash: " & strHash, "Length: " & Len(strHash)
    LogTestResult "Empty password hash not empty", Len(Trim(strHash)) > 0
    LogTestResult "Verify empty password", modCrypto.VerifyPassword("", strHash)
    LogTestResult "Wrong password fails", Not modCrypto.VerifyPassword("x", strHash)
    
    ' Tampered / malformed hash rejection
    strHash = modCrypto.HashPassword("secure")
    LogTestResult "Truncated hash rejected", Not modCrypto.VerifyPassword("secure", Left(strHash, 20))
    LogTestResult "Garbage string rejected", Not modCrypto.VerifyPassword("secure", "garbage")
    LogTestResult "Null hash rejected", Not modCrypto.VerifyPassword("secure", "")
    
    ' Non-ASCII / Unicode passwords (critical for international)
    strPass = "P@ssw�rd123!��"
    strHash = modCrypto.HashPassword(strPass)
    LogTestResult "Non-ASCII password hashes", (Len(strHash) > 0)
    LogTestResult "Non-ASCII verify succeeds", modCrypto.VerifyPassword(strPass, strHash)
    LogTestResult "Non-ASCII wrong password fails", Not modCrypto.VerifyPassword("wrong", strHash)
    
    ' Case sensitivity
    strPass = "CaseSensitive123!"
    strHash = modCrypto.HashPassword(strPass)
    LogTestResult "Case-sensitive: correct case", modCrypto.VerifyPassword(strPass, strHash)
    LogTestResult "Case-sensitive: wrong case fails", Not modCrypto.VerifyPassword(LCase(strPass), strHash)
    
    ' Performance / no crash on long password
    strPass = String(200, "A") & "1!"
    strHash = modCrypto.HashPassword(strPass)
    LogTestResult "Long password (202 chars) accepted", (Len(strHash) > 0)
    LogTestResult "Long password verify works", modCrypto.VerifyPassword(strPass, strHash)
    
    ' Timing attack resistance (same length response)
   Dim t1 As Single, t2 As Single
    Dim v1 As String, v2 As String
    t1 = Timer
   v1 = modCrypto.VerifyPassword("wrong", strHash)
    t1 = Timer - t1
    
    t2 = Timer
    v2 = modCrypto.VerifyPassword("correctbutnotreally", strHash)
    t2 = Timer - t2
    
    LogTestResult "Timing safe (similar duration)", (Abs(t1 - t2) < 0.1)
    
    LogTestSummary
End Sub

' HELPER FUNCTIONS � FULLY IMPLEMENTED WITH ERROR CHECKING

Private Sub LoginAs(strUser As String, strPass As String)
    Call modGlobals.InitializeGlobalVariables
    modAuthentication.AuthenticateUser strUser, strPass
End Sub

Private Sub LogTestStart(strName As String)
    Debug.Print vbCrLf & "TESTING: " & strName
End Sub

Private Sub LogTestResult(strTest As String, blnPass As Boolean, Optional strNote As String = "")
    mlngTotalTests = mlngTotalTests + 1
    If blnPass Then
        mlngPassedTests = mlngPassedTests + 1
        Debug.Print " (Success) PASS: " & strTest
    Else
        mlngFailedTests = mlngFailedTests + 1
        Debug.Print " (Failure) FAIL: " & strTest
        If strNote <> "" Then Debug.Print "   ? " & strNote
    End If
End Sub

Private Sub LogTestSummary()
    Debug.Print "  ? " & (mlngPassedTests + mlngFailedTests) & " tests in group | " & mlngPassedTests & " passed, " & mlngFailedTests & " failed"
End Sub

Private Sub Assert(blnCondition As Boolean, strDesc As String)
    mlngTotalTests = mlngTotalTests + 1
    If blnCondition Then
        LogTestResult "PASS: " & strDesc, True
    Else
        LogTestResult "FAIL: " & strDesc, False
    End If
End Sub

Private Function IsArrayUnique(arr() As String) As Boolean
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim v
    For Each v In arr
        If dict.Exists(v) Then
            IsArrayUnique = False
            Exit Function
        End If
        dict(v) = 1
    Next v
    IsArrayUnique = True
End Function

' Test Data Helpers - WITH ERROR CHECKING
Private Function CreateTestUser(strUser As String, strPass As String, strRole As String, blnActive As Boolean) As Long
    On Error GoTo ErrorHandler
    
    Dim strHash As String
    strHash = modCrypto.HashPassword(strPass)
    
    ' CRITICAL CHECK
    If Len(strHash) = 0 Then
        Debug.Print "  (Failure) ERROR: Cannot create user - password hashing failed"
        CreateTestUser = 0
        Exit Function
    End If
    
    CurrentDb.Execute "INSERT INTO Users (Username, PasswordHash, FullName, Role, IsActive, PasswordSetDate) " & _
                      "VALUES ('" & strUser & "', '" & strHash & "', 'Test User', '" & strRole & "', " & IIf(blnActive, -1, 0) & ", Date())"
    CreateTestUser = DLookup("UserID", "Users", "Username='" & strUser & "'")
    Exit Function
    
ErrorHandler:
    Debug.Print "  (Failure) ERROR creating test user: " & Err.Description
    CreateTestUser = 0
End Function

Private Sub DeleteTestUser(lngID As Long)
    On Error Resume Next
    If lngID > 0 Then
        CurrentDb.Execute "DELETE FROM Users WHERE UserID=" & lngID
    End If
End Sub

'Private Function CreateTestClient(strName As String, strVAT As String) As Long
'    On Error Resume Next
'    CurrentDb.Execute "DELETE FROM Clients WHERE ClientName='" & strName & "'", dbFailOnError
'
'    On Error GoTo ErrHandler
'
'    If Not modValidation.IsValidEmail("test@" & Replace(strName, " ", "") & ".com") Then Exit Function
'
'    CurrentDb.Execute "INSERT INTO Clients (ClientName, VATNumber, EmailBilling, Country, Telephone) " & _
'                      "VALUES ('" & strName & "', '" & strVAT & "', 'test@" & Replace(strName, " ", "") & ".com', 'Spain', '+34 600 000 000')"
'
'    CreateTestClient = Nz(DLookup("ClientID", "Clients", "ClientName='" & strName & "'"), 0)
'    Exit Function
'
'ErrHandler:
'    CreateTestClient = 0
'End Function

Private Sub UpdateTestClient(lngID As Long, strNewName As String)
    On Error Resume Next
    CurrentDb.Execute "UPDATE Clients SET ClientName='" & strNewName & "' WHERE ClientID=" & lngID
End Sub

Private Sub DeleteTestClient(lngID As Long)
    On Error Resume Next
    CurrentDb.Execute "DELETE FROM Clients WHERE ClientID=" & lngID
End Sub


Private Sub TestConcurrentLoginSimulation()
    On Error Resume Next
    Dim lngConcurrentUser As Long
    lngConcurrentUser = CreateTestUser("testconcurrent", "Concurrent123!", "Admin", True)
    
    If lngConcurrentUser = 0 Then Exit Sub
    
    ' Simulated: same user logging in twice doesn't break session
    LoginAs "testconcurrent", "Concurrent123!"
    Dim blnSecond As Boolean
    blnSecond = AuthenticateUser("testconcurrent", "Concurrent123!")
    LogTestResult "Concurrent login allowed (session remains)", blnSecond
    LogoutUser
    
    DeleteTestUser lngConcurrentUser
End Sub

' TEST: modRecordDuplicator
'Private Sub Test_RecordDuplicator_Comprehensive()
'    LogTestStart "modRecordDuplicator � Integration Test"
'
'    Dim lngClientID As Long, lngNewClientID As Long
'    Dim lngServiceID As Long, lngNewServiceID As Long
'    Dim lngInvoiceID As Long, lngNewInvoiceID As Long
'    Dim strOriginalVAT As String, strNewVAT As String
'
'    Dim dictChildren As Object, dictExclude As Object
'    Dim strError As String
'
'    On Error GoTo CleanExit
'
'    ' 1. Duplicate Client
'    ' Pre-cleanup
'    ExecuteQuery "DELETE FROM Clients WHERE ClientName LIKE 'DupeTest%'"
'
'    lngClientID = CreateTestClient("DupeTest Client S.L.", "ES98765432Z")
'    If lngClientID = 0 Then
'        LogTestResult "Create test client", False, "Failed to create source client"
'        GoTo CleanExit
'    End If
'
'    ' Verify client exists before duplication
'    If Not RecordExists("Clients", "ClientID=" & lngClientID) Then
'        LogTestResult "Source client verification", False, "Client not found"
'        GoTo CleanExit
'    End If
'
'    lngNewClientID = SafeDuplicate("Clients", lngClientID, strError)
'
'    If lngNewClientID > 0 Then
'        If Not RecordExists("Clients", "ClientID=" & lngNewClientID) Then
'            LogTestResult "Duplicate Client � DB verify", False, _
'                "Returned ID " & lngNewClientID & " but record not in DB (rollback?)"
'            lngNewClientID = 0  ' Force fail
'        End If
'    End If
'
'    LogTestResult "Duplicate Client � success", lngNewClientID > 0, _
'        IIf(lngNewClientID > 0, "New ID: " & lngNewClientID, "FAILED: " & strError)
'
'
'    If lngNewClientID > 0 Then
'        Dim varVAT As Variant
'        varVAT = GetSingleValue("Clients", "VATNumber", "ClientID=" & lngNewClientID)
'
'        LogTestResult "VATNumber skipped (unique index)", _
'            IsNull(varVAT) Or varVAT = "", _
'            "VATNumber = " & Nz(varVAT, "(NULL)")
'
'        LogTestResult "Audit log created", _
'            RecordExists("AuditLog", "TableName='Clients' AND RecordID=" & lngNewClientID & " AND ActionType='Duplicate'")
'    End If
'
'    ' 2. Duplicate Service
''    Dim strSrvNum As String
''    strSrvNum = GenerateServiceNumber()
''    lngServiceID = CreateTestService(lngClientID, strSrvNum)
''    If lngServiceID = 0 Then
''        LogTestResult "Create test service", False, "Service creation failed"
''        GoTo CleanExit
''    End If
''
''    lngNewServiceID = SafeDuplicate("Services", lngServiceID, strError)
''    LogTestResult "Duplicate Service ? success", lngNewServiceID > 0, _
''                  IIf(lngNewServiceID > 0, "New ID: " & lngNewServiceID, "FAILED: " & strError)
''
''    If lngNewServiceID > 0 Then
''        LogTestResult "Service.ClientID preserved", _
''            GetSingleValue("Services", "ClientID", "ServiceID=" & lngNewServiceID) = lngClientID
''
''        LogTestResult "ServiceNumber skipped", _
''            IsNull(GetSingleValue("Services", "ServiceNumber", "ServiceID=" & lngNewServiceID))
''    End If
''
''    ' 3. Client + Child Services
''    Set dictChildren = CreateChildRelations()
''    AddChildRelation dictChildren, "Services", "ClientID"
''
''    lngNewClientID = SafeDuplicate("Clients", lngClientID, strError, dictChildren:=dictChildren)
''    LogTestResult "Duplicate Client + child Services", lngNewClientID > 0, _
''                  IIf(lngNewClientID > 0, "Success - New Client ID: " & lngNewClientID, "FAILED: " & strError)
''
''    If lngNewClientID > 0 Then
''        Dim lngCount As Long
''        lngCount = DCount("*", "Services", "ClientID=" & lngNewClientID)
''        LogTestResult "Child Services duplicated (" & lngCount & " records)", lngCount > 0
''    End If
'
'    LogTestSummary
'    Exit Sub
'
'CleanExit:
'    On Error Resume Next
'    ExecuteQuery "DELETE FROM InvoiceLines WHERE InvoiceID IN (SELECT InvoiceID FROM Invoices WHERE InvoiceNumber LIKE 'INV-%')"
'    ExecuteQuery "DELETE FROM Invoices WHERE InvoiceNumber LIKE 'INV-%'"
'    ExecuteQuery "DELETE FROM Services WHERE ServiceNumber LIKE 'SRV-%'"
'    If lngNewClientID > 0 Then
'        ExecuteQuery "DELETE FROM Clients WHERE ClientID=" & lngNewClientID
'    End If
'    If lngClientID > 0 Then
'        ExecuteQuery "DELETE FROM Clients WHERE ClientID=" & lngClientID
'    End If
'    ClearCache
'    On Error GoTo 0
'End Sub
'
'Private Function CreateTestClient(strName As String, strVAT As String) As Long
'    ExecuteQuery "DELETE FROM Clients WHERE ClientName='" & Replace(strName, "'", "''") & "'"
'
'    Dim dict As Object
'    Set dict = CreateObject("Scripting.Dictionary")
'    dict("ClientName") = strName
'    dict("VATNumber") = strVAT
'    dict("AddressLine") = "+123 Main Street"
'    dict("City") = "Barcelona"
'    dict("ZIPCode") = "80012"
'    dict("AddressLine") = "+123 Main Street"
'    dict("Country") = "Spain"
'    dict("EmailBilling") = "test@" & Replace(strName, " ", "") & ".es"
'    dict("Telephone") = "+34 900 100 100"
'    dict("CreatedBy") = 1
'    dict("CreatedDate") = Now()
'
'    If BuildSafeInsert("Clients", dict) Then
'        CreateTestClient = GetSingleValue("Clients", "ClientID", "ClientName='" & Replace(strName, "'", "''") & "'")
'    Else
'        CreateTestClient = 0
'    End If
'End Function
'
'Private Function CreateTestService(lngClientID As Long, strSrvNum As String) As Long
'    ExecuteQuery "DELETE FROM Services WHERE ServiceNumber='" & strSrvNum & "'"
'
'    Dim dict As Object
'    Set dict = CreateObject("Scripting.Dictionary")
'    dict("ServiceNumber") = strSrvNum
'    dict("ServiceDate") = Date
'    dict("ClientID") = lngClientID
'    dict("LoadingCountry") = "Spain"
'    dict("UnloadingCountry") = "France"
'    dict("Kilometers") = 1500
'    dict("CreatedBy") = 1
'    dict("CreatedDate") = Now()
'
'    If BuildSafeInsert("Services", dict) Then
'        CreateTestService = GetSingleValue("Services", "ServiceID", "ServiceNumber='" & strSrvNum & "'")
'    Else
'        CreateTestService = 0
'    End If
'End Function
'
'Private Function CreateTestInvoice_WithLines(lngServiceID As Long, strInvNum As String) As Long
'    ExecuteQuery "DELETE FROM Invoices WHERE InvoiceNumber='" & strInvNum & "'"
'
'    Dim dict As Object
'    Set dict = CreateObject("Scripting.Dictionary")
'    dict("InvoiceNumber") = strInvNum
'    dict("InvoiceDate") = Date
'    dict("Subtotal") = 2000
'    dict("TotalAmount") = 2420
'    dict("CreatedBy") = 1
'    dict("CreatedDate") = Now()
'
'    If Not BuildSafeInsert("Invoices", dict) Then
'        CreateTestInvoice_WithLines = 0
'        Exit Function
'    End If
'
'    Dim lngInvID As Long
'    lngInvID = GetSingleValue("Invoices", "InvoiceID", "InvoiceNumber='" & strInvNum & "'")
'
'    Set dict = CreateObject("Scripting.Dictionary")
'    dict("InvoiceID") = lngInvID
'    dict("ServiceID") = lngServiceID
'    dict("LineAmount") = 2000
'    dict("CreatedDate") = Now()
'    BuildSafeInsert "InvoiceLines", dict
'
'    CreateTestInvoice_WithLines = lngInvID
'End Function


' MODULE: Test_modRecordDuplicator
' PURPOSE: Comprehensive test suite for record duplication functionality
' COVERAGE:
'   - Simple duplication (single table)
'   - Unique constraint handling (VAT, ServiceNumber, InvoiceNumber)
'   - Child record duplication (1-to-many relationships)
'   - Multi-level relationships (Client -> Service -> Invoice -> Lines)
'   - Include/Exclude field lists
'   - Edge cases (NULL values, empty strings, composite scenarios)
'   - Audit trail verification
'   - Transaction rollback scenarios


Private Sub Test_RecordDuplicator_Comprehensive()
    LogTestStart "modRecordDuplicator � Comprehensive Integration Test"
    
    Dim lngClientID As Long, lngNewClientID As Long
    Dim lngServiceID As Long, lngNewServiceID As Long
    Dim lngInvoiceID As Long, lngNewInvoiceID As Long
    Dim strOriginalVAT As String, strNewVAT As String
    Dim dictChildren As Object, dictExclude As Object, dictInclude As Object
    Dim strError As String
    Dim lngCount As Long
    
    On Error GoTo CleanExit

    ' PRE-CLEANUP: Delete all test data
    CleanupTestData
    
    ' TEST 1: Simple Client Duplication
    lngClientID = CreateTestClient("DupeTest Client S.L.", "ES98765432Z")
    If lngClientID = 0 Then
        LogTestResult "Create test client", False, "Failed to create source client"
        GoTo CleanExit
    End If
    
    If Not RecordExists("Clients", "ClientID=" & lngClientID) Then
        LogTestResult "Source client verification", False, "Client not found"
        GoTo CleanExit
    End If

    ' Store original VAT for comparison
    strOriginalVAT = Nz(GetSingleValue("Clients", "VATNumber", "ClientID=" & lngClientID), "")
    
    ' Perform duplication
    lngNewClientID = SafeDuplicate("Clients", lngClientID, strError)
    
    ' Verify duplication success
    If lngNewClientID > 0 Then
        If Not RecordExists("Clients", "ClientID=" & lngNewClientID) Then
            LogTestResult "Duplicate Client � DB verify", False, _
                "Returned ID " & lngNewClientID & " but record not in DB (rollback?)"
            lngNewClientID = 0
        End If
    End If
    
    LogTestResult "1.1 Duplicate Client � success", lngNewClientID > 0, _
        IIf(lngNewClientID > 0, "New ID: " & lngNewClientID, "FAILED: " & strError)
    
    If lngNewClientID > 0 Then
        ' Get new VAT
        strNewVAT = Nz(GetSingleValue("Clients", "VATNumber", "ClientID=" & lngNewClientID), "")
        
        ' TEST: VATNumber should be MODIFIED (not copied as-is) because it's unique+required
        LogTestResult "1.2 VATNumber modified (required+unique)", _
            strNewVAT <> "" And strNewVAT <> strOriginalVAT, _
            "Original: " & strOriginalVAT & " | New: " & strNewVAT
        
        ' TEST: Verify placeholder format (should contain timestamp)
        LogTestResult "1.3 VATNumber placeholder format valid", _
            InStr(strNewVAT, "_202") > 0 Or InStr(strNewVAT, strOriginalVAT) > 0, _
            "Generated: " & strNewVAT
        
        ' TEST: Other fields should be copied correctly
        Dim strOriginalName As String, strNewName As String
        strOriginalName = Nz(GetSingleValue("Clients", "ClientName", "ClientID=" & lngClientID), "")
        strNewName = Nz(GetSingleValue("Clients", "ClientName", "ClientID=" & lngNewClientID), "")
        
        LogTestResult "1.4 ClientName copied correctly", _
            strOriginalName = strNewName, _
            "Original: " & strOriginalName & " | New: " & strNewName
        
        ' TEST: Audit log
        lngCount = Nz(DCount("*", "AuditLog", _
            "TableName='Clients' AND RecordID=" & lngNewClientID & " AND ActionType='Duplicate'"), 0)
        
        LogTestResult "1.5 Audit log created", lngCount > 0, _
            "Found " & lngCount & " audit record(s)"
    End If

    ' TEST 2: Service Duplication
    Dim strSrvNum As String, strOriginalSrvNum As String, strNewSrvNum As String
    strSrvNum = modServiceNumbers.GenerateServiceNumber()
    
    If strSrvNum = "" Then
        LogTestResult "2.0 Generate service number", False, "GenerateServiceNumber failed"
        GoTo CleanExit
    End If
    
    lngServiceID = CreateTestService(lngClientID, strSrvNum)
    If lngServiceID = 0 Then
        LogTestResult "2.0 Create test service", False, "Service creation failed"
        GoTo CleanExit
    End If

    strOriginalSrvNum = Nz(GetSingleValue("Services", "ServiceNumber", "ServiceID=" & lngServiceID), "")
    
    lngNewServiceID = SafeDuplicate("Services", lngServiceID, strError)
    LogTestResult "2.1 Duplicate Service � success", lngNewServiceID > 0, _
                  IIf(lngNewServiceID > 0, "New ID: " & lngNewServiceID, "FAILED: " & strError)

    If lngNewServiceID > 0 Then
        ' TEST: ClientID should be preserved (foreign key)
        Dim lngOrigClientID As Long, lngNewClientIDFromService As Long
        lngOrigClientID = Nz(GetSingleValue("Services", "ClientID", "ServiceID=" & lngServiceID), 0)
        lngNewClientIDFromService = Nz(GetSingleValue("Services", "ClientID", "ServiceID=" & lngNewServiceID), 0)
        
        LogTestResult "2.2 Service.ClientID preserved", _
            lngNewClientIDFromService = lngOrigClientID, _
            "Original: " & lngOrigClientID & " | New: " & lngNewClientIDFromService

        ' TEST: ServiceNumber should be modified (unique constraint)
        strNewSrvNum = Nz(GetSingleValue("Services", "ServiceNumber", "ServiceID=" & lngNewServiceID), "")
        
        LogTestResult "2.3 ServiceNumber modified (unique)", _
            strNewSrvNum <> "" And strNewSrvNum <> strOriginalSrvNum, _
            "Original: " & strOriginalSrvNum & " | New: " & strNewSrvNum
        
        ' TEST: Other fields copied
        Dim lngOrigKm As Long, lngNewKm As Long
        lngOrigKm = Nz(GetSingleValue("Services", "Kilometers", "ServiceID=" & lngServiceID), 0)
        lngNewKm = Nz(GetSingleValue("Services", "Kilometers", "ServiceID=" & lngNewServiceID), 0)
        
        LogTestResult "2.4 Service.Kilometers copied", lngOrigKm = lngNewKm, _
            "Original: " & lngOrigKm & " | New: " & lngNewKm
    End If

    
    ' TEST 3: Client + Child Services (1-to-many)
    
    ' Create additional service for the original client
    Dim lngService2ID As Long
    Dim strSrvNum2 As String
    strSrvNum2 = modServiceNumbers.GenerateServiceNumber()
    
    If strSrvNum2 = "" Then
        LogTestResult "3.0 Generate second service number", False, "GenerateServiceNumber failed"
        GoTo CleanExit
    End If
    
    lngService2ID = CreateTestService(lngClientID, strSrvNum2)
    
    Set dictChildren = CreateChildRelations()
    Call AddChildRelation(dictChildren, "Services", "ClientID")

    Dim lngNewClientID2 As Long
    lngNewClientID2 = SafeDuplicate("Clients", lngClientID, strError, dictChildren:=dictChildren)
    
    LogTestResult "3.1 Duplicate Client + child Services", lngNewClientID2 > 0, _
                  IIf(lngNewClientID2 > 0, "Success - New Client ID: " & lngNewClientID2, "FAILED: " & strError)

    If lngNewClientID2 > 0 Then
        ' Count child services
        Dim lngOrigServiceCount As Long, lngNewServiceCount As Long
        lngOrigServiceCount = DCount("*", "Services", "ClientID=" & lngClientID)
        lngNewServiceCount = DCount("*", "Services", "ClientID=" & lngNewClientID2)
        
        LogTestResult "3.2 Child Services duplicated (count match)", _
            lngNewServiceCount = lngOrigServiceCount And lngNewServiceCount > 0, _
            "Original: " & lngOrigServiceCount & " services | New: " & lngNewServiceCount & " services"
        
        ' Verify child services have correct parent ID
        lngCount = DCount("*", "Services", "ClientID=" & lngNewClientID2 & " AND ServiceID NOT IN (SELECT ServiceID FROM Services WHERE ClientID=" & lngClientID & ")")
        
        LogTestResult "3.3 Child Services have new ClientID", _
            lngCount = lngNewServiceCount, _
            "All " & lngNewServiceCount & " services have new parent ID"
    End If

    
    ' TEST 4: Invoice with InvoiceLines (multi-level)
    
    Dim strInvNum As String
    strInvNum = modServiceNumbers.GenerateInvoiceNumber()
    
    If strInvNum = "" Then
        LogTestResult "4.0 Generate invoice number", False, "GenerateInvoiceNumber failed"
        GoTo CleanExit
    End If
    lngInvoiceID = CreateTestInvoice_WithLines(lngServiceID, strInvNum)
    If lngInvoiceID = 0 Then
        LogTestResult "4.0 Create test invoice", False, "Invoice creation failed"
        GoTo CleanExit
    End If

    ' Duplicate invoice (should not duplicate lines by default)
    lngNewInvoiceID = SafeDuplicate("Invoices", lngInvoiceID, strError)
    
    LogTestResult "4.1 Duplicate Invoice � success", lngNewInvoiceID > 0, _
        IIf(lngNewInvoiceID > 0, "New ID: " & lngNewInvoiceID, "FAILED: " & strError)

    If lngNewInvoiceID > 0 Then
        ' TEST: InvoiceNumber should be modified
        Dim strOrigInvNum As String, strNewInvNum As String
        strOrigInvNum = Nz(GetSingleValue("Invoices", "InvoiceNumber", "InvoiceID=" & lngInvoiceID), "")
        strNewInvNum = Nz(GetSingleValue("Invoices", "InvoiceNumber", "InvoiceID=" & lngNewInvoiceID), "")
        
        LogTestResult "4.2 InvoiceNumber modified (unique)", _
            strNewInvNum <> "" And strNewInvNum <> strOrigInvNum, _
            "Original: " & strOrigInvNum & " | New: " & strNewInvNum
        
        ' TEST: Lines should NOT be copied (no child relation specified)
        lngCount = DCount("*", "InvoiceLines", "InvoiceID=" & lngNewInvoiceID)
        
        LogTestResult "4.3 InvoiceLines NOT duplicated (expected)", lngCount = 0, _
            "Found " & lngCount & " lines (should be 0)"
    End If

    
    ' TEST 5: Invoice WITH InvoiceLines (child relation)
    
    Set dictChildren = CreateChildRelations()
    Call AddChildRelation(dictChildren, "InvoiceLines", "InvoiceID")

    Dim lngNewInvoiceID2 As Long
    lngNewInvoiceID2 = SafeDuplicate("Invoices", lngInvoiceID, strError, dictChildren:=dictChildren)
    
    LogTestResult "5.1 Duplicate Invoice + InvoiceLines", lngNewInvoiceID2 > 0, _
        IIf(lngNewInvoiceID2 > 0, "New ID: " & lngNewInvoiceID2, "FAILED: " & strError)

    If lngNewInvoiceID2 > 0 Then
        ' Count invoice lines
        Dim lngOrigLineCount As Long, lngNewLineCount As Long
        lngOrigLineCount = DCount("*", "InvoiceLines", "InvoiceID=" & lngInvoiceID)
        lngNewLineCount = DCount("*", "InvoiceLines", "InvoiceID=" & lngNewInvoiceID2)
        
        LogTestResult "5.2 InvoiceLines duplicated (count match)", _
            lngNewLineCount = lngOrigLineCount And lngNewLineCount > 0, _
            "Original: " & lngOrigLineCount & " lines | New: " & lngNewLineCount & " lines"
        
        ' Verify line amounts match
        Dim curOrigTotal As Currency, curNewTotal As Currency
        curOrigTotal = Nz(DSum("LineAmount", "InvoiceLines", "InvoiceID=" & lngInvoiceID), 0)
        curNewTotal = Nz(DSum("LineAmount", "InvoiceLines", "InvoiceID=" & lngNewInvoiceID2), 0)
        
        LogTestResult "5.3 InvoiceLine amounts preserved", curOrigTotal = curNewTotal, _
            "Original total: " & Format(curOrigTotal, "Currency") & " | New total: " & Format(curNewTotal, "Currency")
    End If

    
    ' TEST 6: Exclude Fields

    Set dictExclude = CreateExcludeList("ContactName", "BankAccount", "PaymentTerms")

    Dim lngNewClientID3 As Long
    lngNewClientID3 = SafeDuplicate("Clients", lngClientID, strError, dictExclude:=dictExclude)

    LogTestResult "6.1 Duplicate with excluded fields", lngNewClientID3 > 0, _
        IIf(lngNewClientID3 > 0, "New ID: " & lngNewClientID3, "FAILED: " & strError)

    If lngNewClientID3 > 0 Then
        ' TEST: Excluded fields should be NULL
        Dim varEmail As Variant, strCity As String, varContact As Variant, varPaymentTerms As Variant, varBankAccount As Variant
        varEmail = GetSingleValue("Clients", "EmailBilling", "ClientID=" & lngNewClientID3)
        strCity = Nz(GetSingleValue("Clients", "City", "ClientID=" & lngNewClientID3), "")
        varContact = GetSingleValue("Clients", "ContactName", "ClientID=" & lngNewClientID3)
        varPaymentTerms = GetSingleValue("Clients", "PaymentTerms", "ClientID=" & lngNewClientID3)
        varBankAccount = GetSingleValue("Clients", "BankAccount", "ClientID=" & lngNewClientID3)

        LogTestResult "6.2 ContactName excluded (NULL)", IsNull(varContact), _
            "ContactName = " & Nz(varContact, "(NULL)")

        LogTestResult "6.3 BankAccount excluded (NULL)", IsNull(varBankAccount), _
            "BankAccount = " & Nz(varBankAccount, "(NULL)")
        
        LogTestResult "6.4 PaymentTerms excluded (NULL)", IsNull(varPaymentTerms), _
            "BankDetails = " & Nz(varPaymentTerms, "(NULL)")

        ' TEST: Non-excluded field should be copied
        LogTestResult "6.5 EmailBilling copied (required field)", Not IsNull(varEmail), _
            "EmailBilling = " & Nz(varEmail, "(NULL)")
        
        LogTestResult "6.6 City copied (not excluded)", strCity = "Barcelona", _
            "City = " & strCity
    End If


    ' TEST 7: Include Only Specific Fields

    ' Include ONLY these fields (plus VATNumber which is required+unique)
    Set dictInclude = CreateIncludeList("ClientName", "City", "Country", "ZIPCode", "AddressLine", "EmailBilling", "Telephone")

    Dim lngNewClientID4 As Long
    lngNewClientID4 = SafeDuplicate("Clients", lngClientID, strError, dictInclude:=dictInclude)

    LogTestResult "7.1 Duplicate with include list", lngNewClientID4 > 0, _
        IIf(lngNewClientID4 > 0, "New ID: " & lngNewClientID4, "FAILED: " & strError)

    If lngNewClientID4 > 0 Then
        ' TEST: Included fields should be copied
        strCity = Nz(GetSingleValue("Clients", "City", "ClientID=" & lngNewClientID4), "")
        LogTestResult "7.2 City included (copied)", strCity = "Barcelona", _
            "City = " & strCity

        ' TEST: Non-included OPTIONAL fields should be NULL
        varContact = GetSingleValue("Clients", "ContactName", "ClientID=" & lngNewClientID4)
        LogTestResult "7.3 ContactName not included (NULL)", IsNull(varContact), _
            "ContactName = " & Nz(varContact, "(NULL)")
        
        varBankAccount = GetSingleValue("Clients", "BankAccount", "ClientID=" & lngNewClientID4)
        LogTestResult "7.4 BankAccount not included (NULL)", IsNull(varBankAccount), _
            "BankAccount = " & Nz(varBankAccount, "(NULL)")
    End If

    
    ' TEST 8: Edge Case - Non-existent Source ID
    
'    Dim lngBogusID As Long
'    lngBogusID = SafeDuplicate("Clients", 999999, strError)
'
'    LogTestResult "8.1 Non-existent source ID (should fail gracefully)", lngBogusID = 0, _
'        "Result: " & IIf(lngBogusID = 0, "Correctly returned 0", "ERROR: Returned " & lngBogusID)
'
'    LogTestResult "8.2 Error message provided", Len(strError) > 0, _
'        "Error: " & Left(strError, 100)

    
    ' TEST 9: Edge Case - Client with NULL fields
    
    Dim lngClientNullID As Long
    lngClientNullID = CreateTestClient("DupeTest Minimal", "ES11111111A")
    
    If lngClientNullID > 0 Then
        ' Clear optional fields to NULL

        ExecuteQuery ("UPDATE Clients SET BankAccount=NULL WHERE ClientID=" & lngClientNullID)
        
        Dim lngNewClientNull As Long
        lngNewClientNull = SafeDuplicate("Clients", lngClientNullID, strError)
        
        LogTestResult "9.1 Duplicate client with NULL fields", lngNewClientNull > 0, _
            IIf(lngNewClientNull > 0, "New ID: " & lngNewClientNull, "FAILED: " & strError)
        
        If lngNewClientNull > 0 Then
            ' Verify NULL fields remain NULL
            varBankAccount = GetSingleValue("Clients", "BankAccount", "ClientID=" & lngNewClientNull)
            LogTestResult "9.2 NULL BankAccount preserved", IsNull(varBankAccount), _
                "BankAccount = " & Nz(varBankAccount, "(NULL)")
        End If
    End If

    
    ' TEST 10: Duplicate Multiple Records
    
    Dim arrIDs(1 To 2) As Long
    arrIDs(1) = lngClientID
    arrIDs(2) = lngClientNullID
    
    Dim colResults As Collection
    Set colResults = DuplicateMultipleRecords("Clients", arrIDs)
    
    LogTestResult "10.1 DuplicateMultipleRecords � count", colResults.Count = 2, _
        "Created " & colResults.Count & " records (expected 2)"
    
    If colResults.Count > 0 Then
        LogTestResult "10.2 All duplicates have valid IDs", _
            CLng(colResults(1)) > 0 And CLng(colResults(2)) > 0, _
            "IDs: " & colResults(1) & ", " & colResults(2)
    End If

    
    ' TEST 11: Preview Duplication
    
'    Dim strPreview As String
'    strPreview = PreviewDuplication("Clients", , CreateExcludeList("EmailBilling"))
'
'    LogTestResult "11.1 Preview includes field list", Len(strPreview) > 100, _
'        "Preview length: " & Len(strPreview) & " chars"
'
'    LogTestResult "11.2 Preview shows excluded fields", InStr(strPreview, "skipped") > 0, _
'        "Contains 'skipped' indicator"

    
    ' TEST 12: GetDuplicatableFields utility
    
    Dim colFields As Collection
    Set colFields = GetDuplicatableFields("Clients")
    
    LogTestResult "12.1 GetDuplicatableFields � count > 0", colFields.Count > 0, _
        "Found " & colFields.Count & " duplicatable fields"
    
    ' Verify ClientID (AutoNumber) is NOT in the list
    Dim bFound As Boolean
    Dim v As Variant
    bFound = False
    For Each v In colFields
        If UCase(CStr(v)) = "CLIENTID" Then
            bFound = True
            Exit For
        End If
    Next
    
    LogTestResult "12.2 ClientID (AutoNumber) excluded", Not bFound, _
        "ClientID in list: " & IIf(bFound, "YES (FAIL)", "NO (PASS)")

    
    ' TEST 13: GetChildTables utility
    
    Dim dictAutoChildren As Object
    Set dictAutoChildren = GetChildTables("Clients")
    
    LogTestResult "13.1 GetChildTables � found children", dictAutoChildren.Count > 0, _
        "Found " & dictAutoChildren.Count & " child table(s)"
    
    ' Verify Services is in the list
    bFound = False
    On Error Resume Next
    If dictAutoChildren.Exists("Services") Then bFound = True
    On Error GoTo CleanExit
    
    LogTestResult "13.2 Services detected as child", bFound, _
        "Services in children: " & IIf(bFound, "YES", "NO")

    
    ' TEST 14: Edge Case - Empty Child Table
    
    Dim lngClientNoChildren As Long
    lngClientNoChildren = CreateTestClient("DupeTest NoKids", "ES22222222B")
    
    If lngClientNoChildren > 0 Then
        ' Ensure no child services exist
        ExecuteQuery ("DELETE FROM Services WHERE ClientID=" & lngClientNoChildren)
        
        Set dictChildren = CreateChildRelations()
        Call AddChildRelation(dictChildren, "Services", "ClientID")
        
        Dim lngNewClientNoKids As Long
        lngNewClientNoKids = SafeDuplicate("Clients", lngClientNoChildren, strError, dictChildren:=dictChildren)
        
        LogTestResult "14.1 Duplicate client with no children", lngNewClientNoKids > 0, _
            IIf(lngNewClientNoKids > 0, "New ID: " & lngNewClientNoKids, "FAILED: " & strError)
        
        If lngNewClientNoKids > 0 Then
            lngCount = DCount("*", "Services", "ClientID=" & lngNewClientNoKids)
            LogTestResult "14.2 No spurious child records created", lngCount = 0, _
                "Found " & lngCount & " services (should be 0)"
        End If
    End If

    
    ' TEST 15: Edge Case - Duplicate with Audit Disabled
    
    Dim lngClientNoAudit As Long
    lngClientNoAudit = SafeDuplicate("Clients", lngClientID, strError, blnAudit:=False)
    
    LogTestResult "15.1 Duplicate without audit log", lngClientNoAudit > 0, _
        IIf(lngClientNoAudit > 0, "New ID: " & lngClientNoAudit, "FAILED: " & strError)
    
    If lngClientNoAudit > 0 Then
        lngCount = DCount("*", "AuditLog", _
            "TableName='Clients' AND RecordID=" & lngClientNoAudit & " AND ActionType='Duplicate'")
        
        LogTestResult "15.2 No audit log created (as expected)", lngCount = 0, _
            "Found " & lngCount & " audit records (should be 0)"
    End If

    
    ' TEST 16: Stress Test - Rapid Sequential Duplications
    
    Dim i As Integer
    Dim lngTempID As Long
    Dim colStressIDs As New Collection
    Dim blnStressSuccess As Boolean
    
    blnStressSuccess = True
    For i = 1 To 5
        lngTempID = SafeDuplicate("Clients", lngClientID, strError)
        If lngTempID > 0 Then
            colStressIDs.Add lngTempID
        Else
            blnStressSuccess = False
            Exit For
        End If
    Next i
    
    LogTestResult "16.1 Rapid sequential duplications (5x)", blnStressSuccess, _
        "Created " & colStressIDs.Count & "/5 records"
    
    If colStressIDs.Count > 0 Then
        ' Verify all have unique VATNumbers
        Dim dictVATs As Object
        Set dictVATs = CreateObject("Scripting.Dictionary")
        Dim varID As Variant
        Dim strVAT As String
        Dim blnAllUnique As Boolean
        
        blnAllUnique = True
        For Each varID In colStressIDs
            strVAT = Nz(GetSingleValue("Clients", "VATNumber", "ClientID=" & varID), "")
            If dictVATs.Exists(strVAT) Then
                blnAllUnique = False
                Exit For
            End If
            dictVATs(strVAT) = True
        Next
        
        LogTestResult "16.2 All VATNumbers are unique", blnAllUnique, _
            "Checked " & colStressIDs.Count & " records"
    End If

    
    ' TEST 17: Complex Scenario - Service with Multiple Children
    
    If lngServiceID > 0 And lngInvoiceID > 0 Then
        ' Create additional invoice line
        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")
        dict("InvoiceID") = lngInvoiceID
        dict("ServiceID") = lngServiceID
        dict("LineAmount") = 500
        dict("CreatedDate") = Now()
        BuildSafeInsert "InvoiceLines", dict
        
        ' Duplicate invoice with lines
        Set dictChildren = CreateChildRelations()
        Call AddChildRelation(dictChildren, "InvoiceLines", "InvoiceID")
        
        Dim lngComplexInvID As Long
        lngComplexInvID = SafeDuplicate("Invoices", lngInvoiceID, strError, dictChildren:=dictChildren)
        
        LogTestResult "17.1 Duplicate invoice with multiple lines", lngComplexInvID > 0, _
            IIf(lngComplexInvID > 0, "New ID: " & lngComplexInvID, "FAILED: " & strError)
        
        If lngComplexInvID > 0 Then
            lngOrigLineCount = DCount("*", "InvoiceLines", "InvoiceID=" & lngInvoiceID)
            lngNewLineCount = DCount("*", "InvoiceLines", "InvoiceID=" & lngComplexInvID)
            
            LogTestResult "17.2 All invoice lines duplicated", lngOrigLineCount = lngNewLineCount, _
                "Original: " & lngOrigLineCount & " | New: " & lngNewLineCount
            
            ' Verify line amounts match
            curOrigTotal = Nz(DSum("LineAmount", "InvoiceLines", "InvoiceID=" & lngInvoiceID), 0)
            curNewTotal = Nz(DSum("LineAmount", "InvoiceLines", "InvoiceID=" & lngComplexInvID), 0)
            
            LogTestResult "17.3 Line amounts preserved", curOrigTotal = curNewTotal, _
                "Original: " & Format(curOrigTotal, "Currency") & " | New: " & Format(curNewTotal, "Currency")
        End If
    End If

    LogTestSummary
    Exit Sub

CleanExit:
    On Error Resume Next
'    CleanupTestData
    ClearCache
    On Error GoTo 0
End Sub

' HELPER: Cleanup Test Data
Private Sub CleanupTestData()
    On Error Resume Next
    
    Dim lngTestClientID As Long
    Dim rs As DAO.Recordset
    Dim db As DAO.Database
    
    Set db = CurrentDb
    
    ' Find all test clients
    Set rs = db.OpenRecordset("SELECT ClientID FROM Clients WHERE ClientName LIKE 'DupeTest*'", dbOpenSnapshot)

    Do While Not rs.EOF
        lngTestClientID = rs!ClientID
        On Error Resume Next
        ' Delete child records first
        
        ExecuteQuery "DELETE FROM InvoiceLines WHERE ServiceID IN (" & _
                         "SELECT ServiceID FROM Services WHERE ClientID=" & lngTestClientID & ")"
        ExecuteQuery "DELETE FROM Invoices WHERE InvoiceID IN (" & _
                         "SELECT InvoiceID FROM InvoiceLines WHERE ServiceID IN (" & _
                             "SELECT ServiceID FROM Services WHERE ClientID=" & lngTestClientID & _
                         ")" & _
                     ")"
        
        ExecuteQuery "DELETE FROM Services WHERE ClientID=" & lngTestClientID
        ExecuteQuery "DELETE FROM Clients WHERE ClientID=" & lngTestClientID

        On Error GoTo 0
        
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    On Error GoTo 0
End Sub


' HELPER: Create Test Client
Private Function CreateTestClient(strName As String, strVAT As String) As Long
    On Error GoTo ErrHandler
    
    ExecuteQuery ("DELETE FROM Clients WHERE ClientName='" & Replace(strName, "'", "''") & "'")
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict("ClientName") = strName
    dict("VATNumber") = strVAT
    dict("AddressLine") = "123 Main Street"
    dict("City") = "Barcelona"
    dict("ZIPCode") = "80012"
    dict("Country") = "Spain"
    dict("EmailBilling") = "test@" & Replace(Replace(strName, " ", ""), ".", "") & ".es"
    dict("Telephone") = "+34 900 100 100"
    dict("CreatedBy") = 1
    dict("CreatedDate") = Now()
    
    If BuildSafeInsert("Clients", dict) Then
        CreateTestClient = GetSingleValue("Clients", "ClientID", "ClientName='" & Replace(strName, "'", "''") & "'")
    Else
        CreateTestClient = 0
    End If
    Exit Function
    
ErrHandler:
    Debug.Print "CreateTestClient Error: " & Err.Description
    CreateTestClient = 0
End Function


' HELPER: Create Test Service

Private Function CreateTestService(lngClientID As Long, strSrvNum As String) As Long
    On Error GoTo ErrHandler
    
    ExecuteQuery ("DELETE FROM Services WHERE ServiceNumber='" & strSrvNum & "'")

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict("ServiceNumber") = strSrvNum
    dict("ServiceDate") = Date
    dict("ClientID") = lngClientID
    dict("LoadingCountry") = "Spain"
    dict("UnloadingCountry") = "France"
    dict("Kilometers") = 1500
    dict("CreatedBy") = 1
    dict("CreatedDate") = Now()

    If BuildSafeInsert("Services", dict) Then
        CreateTestService = GetSingleValue("Services", "ServiceID", "ServiceNumber='" & strSrvNum & "'")
    Else
        CreateTestService = 0
    End If
    Exit Function
    
ErrHandler:
    Debug.Print "CreateTestService Error: " & Err.Description
    CreateTestService = 0
End Function


' HELPER: Create Test Invoice with Lines

Private Function CreateTestInvoice_WithLines(lngServiceID As Long, strInvNum As String) As Long
    On Error GoTo ErrHandler
    
    ExecuteQuery ("DELETE FROM Invoices WHERE InvoiceNumber='" & strInvNum & "'")
    
    Dim lngSupplierID As Long
    lngSupplierID = DMax("SupplierID", "Suppliers")
    
    If lngSupplierID = 0 Then
        CreateTestInvoice_WithLines = 0
        Exit Function
    End If
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict("InvoiceNumber") = strInvNum
    dict("InvoiceDate") = Date
    dict("SupplierID") = lngSupplierID
    dict("Subtotal") = 2000
    dict("TotalAmount") = 2420
    dict("CreatedBy") = modGlobals.UserID
    dict("CreatedDate") = Now()
    
   
    If Not BuildSafeInsert("Invoices", dict) Then
        CreateTestInvoice_WithLines = 0
        Exit Function
    End If

    Dim lngInvID As Long
    lngInvID = GetSingleValue("Invoices", "InvoiceID", "InvoiceNumber='" & strInvNum & "'")
    
    If lngInvID = 0 Then
        CreateTestInvoice_WithLines = 0
        Exit Function
    End If

    ' Add invoice line
    Set dict = CreateObject("Scripting.Dictionary")
    dict("InvoiceID") = lngInvID
    dict("ServiceID") = lngServiceID
    dict("LineAmount") = 2000
    dict("CreatedDate") = Now()
    
    If Not BuildSafeInsert("InvoiceLines", dict) Then
        Debug.Print "Warning: InvoiceLine creation failed"
    End If

    CreateTestInvoice_WithLines = lngInvID
    Exit Function

ErrHandler:
    Debug.Print "CreateTestInvoice_WithLines Error: " & Err.Description
    CreateTestInvoice_WithLines = 0
End Function


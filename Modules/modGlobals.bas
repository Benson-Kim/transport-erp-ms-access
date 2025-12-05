'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MODULE    : modGlobals
' PURPOSE   : Centralized global constants, variables, and application-wide
'               settings with enterprise-grade error handling and logging
' AUTHOR    : Tifamovs Saeol <tifamovs@gmail.com> [Benson Kimathi]
' CREATED   : November 18, 2025
' UPDATED   : November 23, 2025 
'    - Encapsulated globals with properties, fixed activity reset,
'    - Added session timeout check on updates, added AuditAction enum.
' NOTES     : 
'    This module serves as the backbone for managing application-wide
'    settings and user session state in a secure and maintainable manner.
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
Public Const SEARCH_DELAY_MS As Long = 600

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

Public Enum AuditActionType
    actInsert = 1
    actUpdate = 2
    actDelete = 3
    actLogin = 4
    actLogout = 5
    actSecurityAlert = 6
    actUserManagement = 7
    actPurgeOld = 8
End Enum


' PROPERTIES FOR SESSION STATE
Public Property Get UserID() As Long
    UserID = g_lngUserID
End Property

Public Property Let UserID(ByVal lngValue As Long)
    g_lngUserID = lngValue
End Property

Public Property Get Username() As String
    Username = g_strUsername
End Property

Public Property Let Username(ByVal strValue As String)
    g_strUsername = strValue
End Property

Public Property Get FullName() As String
    FullName = g_strFullName
End Property

Public Property Let FullName(ByVal strValue As String)
    g_strFullName = strValue
End Property

Public Property Get UserRole() As String
    UserRole = g_strUserRole
End Property

Public Property Let UserRole(ByVal strValue As String)
    g_strUserRole = strValue
End Property

Public Property Get LastActivity() As Date
    LastActivity = g_dtLastActivity
End Property

Public Property Let LastActivity(ByVal dtValue As Date)
    g_dtLastActivity = dtValue
End Property



' PROPERTIES FOR PATHS

Public Property Get backEndPath() As String
    backEndPath = g_strBackendPath
End Property

Public Property Let backEndPath(ByVal strValue As String)
    g_strBackendPath = strValue
End Property

Public Property Get PDFStoragePath() As String
    PDFStoragePath = g_strPDFStoragePath
End Property

Public Property Let PDFStoragePath(ByVal strValue As String)
    g_strPDFStoragePath = strValue
End Property

Public Property Get BackupPath() As String
    BackupPath = g_strBackupPath
End Property

Public Property Let BackupPath(ByVal strValue As String)
    g_strBackupPath = strValue
End Property

' PUBLIC: InitializeGlobalVariables
' Purpose : Reset all session state and load system paths at startup
' Returns : Boolean � True if successful

Public Function InitializeGlobalVariables() As Boolean
    On Error GoTo ErrorHandler

    ' Reset user session
    UserID = 0
    Username = ""
    FullName = ""
    UserRole = ""
    LastActivity = #1/1/1900#  ' Default to old date - reset on login

    EnsurePermissionsLoaded
    
    modAudit.InitializeAuditQueue
    
    ' Load persistent paths from SystemSettings
    backEndPath = GetSystemSetting("BackendDatabasePath", "")
    PDFStoragePath = GetSystemSetting("PDFStoragePath", "")
    BackupPath = GetSystemSetting("BackupPath", "")

    InitializeGlobalVariables = True
    Exit Function

ErrorHandler:
    modUtilities.LogError MODULE_NAME & ".InitializeGlobalVariables", Err.number, Err.Description, ""
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
    modUtilities.LogError MODULE_NAME & ".GetSystemSetting", Err.number, Err.Description, _
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
    rs!ModifiedBy = UserID
    rs!ModifiedDate = Now()
    rs.Update

    SetSystemSetting = True
    GoTo CleanExit

ErrorHandler:
    modUtilities.LogError MODULE_NAME & ".SetSystemSetting", Err.number, Err.Description, _
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
    
    If CheckSessionTimeout() Then
        MsgBox "Your session has expired due to inactivity." & vbCrLf & _
               "You will be logged out for security reasons.", vbInformation, APP_NAME
        
        LogoutUser
    Else
        LastActivity = Now()
    End If
End Sub


' PUBLIC: CheckSessionTimeout
' Purpose : Determine if user session has expired

Public Function CheckSessionTimeout() As Boolean
    On Error GoTo ErrorHandler

    Dim lngMinutes As Long
    lngMinutes = DateDiff("n", LastActivity, Now())

    CheckSessionTimeout = (lngMinutes >= SESSION_TIMEOUT_MINUTES)
    Exit Function

ErrorHandler:
    modUtilities.LogError MODULE_NAME & ".CheckSessionTimeout", Err.number, Err.Description
    CheckSessionTimeout = True  ' Fail closed � force logout on error
End Function


' PUBLIC: GetCurrentUser
' Purpose : Return formatted string for display in UI

Public Function GetCurrentUser() As String
    If UserID > 0 Then
        GetCurrentUser = FullName & " (" & UserRole & ")"
    Else
        GetCurrentUser = "Not logged in"
    End If
End Function


' PUBLIC: IsUserLoggedIn
' Purpose : Simple boolean check used throughout application

Public Function IsUserLoggedIn() As Boolean
    IsUserLoggedIn = (UserID > 0)
End Function

' PUBLIC: EnsurePermissionsLoaded
' Purpose : Load permissions if not yet loaded

Public Sub EnsurePermissionsLoaded()
    If modPermissions.mdictPermissions Is Nothing Then
        modPermissions.LoadPermissions
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MODULE    : modPermissions
' PURPOSE   : Centralized, high-performance permission checking engine
'          with exhaustive permission constants and role-based matrix
' SECURITY  : All security decisions flow through this single module
' AUTHOR    : Tifamovs Saeol <tifamovs@gmail.com> [Benson Kimathi]
' CREATED   : November 17, 2025
' UPDATED   : November 17, 2025
' NOTES     : 
'   - Permissions are defined as string constants for clarity
'   - Permission matrix is loaded into a dictionary on app startup
'   - HasPermission function uses exhaustive Select Case for speed
'   - RequirePermission sub shows friendly message and cancels action
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Compare Database
Option Explicit

Private Const MODULE_NAME As String = "modPermissions"

' PERMISSION CONSTANTS � Used throughout the entire application
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

Public mdictPermissions As Object



' PUBLIC: LoadPermissions � Load from DB on app startup (call from modStartup)
Public Sub LoadPermissions()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim strKey As String

    Set mdictPermissions = Nothing
    Set mdictPermissions = CreateObject("Scripting.Dictionary")

    mdictPermissions.CompareMode = vbTextCompare

    Set db = CurrentDb
    strSQL = "SELECT Role, Permission, Allowed FROM RolePermissions"

    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)

    Do While Not rs.EOF
        strKey = UCase(rs!Role & "|" & rs!Permission)
        mdictPermissions(strKey) = Nz(rs!Allowed, False)
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    
    Exit Sub

ErrorHandler:
    modUtilities.LogError MODULE_NAME & ".LoadPermissions", Err.number, Err.Description
    ' Fail closed: empty dict means all denied
End Sub

' Function : HasPermission � Core permission check
' Purpose  : Returns True if current user (g_strUserRole) has the requested permission
'            Uses exhaustive Select Case � fastest possible evaluation
' Returns  : Boolean

Public Function HasPermission(ByVal strPermission As String) As Boolean
    On Error GoTo ErrorHandler

    Dim strRole As String, strKey As String
    
    EnsurePermissionsLoaded
    
    strRole = UCase(Trim(modGlobals.UserRole))
    
    If strRole = "" Then
        HasPermission = False
        Exit Function
    End If
    
    ' Check DB matrix (loaded on startup)
    strKey = strRole & "|" & UCase(strPermission)

    If mdictPermissions.Exists(strKey) Then
        HasPermission = mdictPermissions(strKey)
        Exit Function
    End If
    
    Select Case UCase(strPermission)
        ' CLIENT PERMISSIONS
        Case PERM_VIEW_CLIENTS
            HasPermission = True  ' All roles can view clients

        Case PERM_EDIT_CLIENTS
            HasPermission = (strRole = "ADMIN" Or strRole = "MANAGER")

        Case PERM_DELETE_CLIENTS
            HasPermission = (strRole = "ADMIN")

        ' SUPPLIER PERMISSIONS
        Case PERM_VIEW_SUPPLIERS
            HasPermission = True  ' All roles can view suppliers

        Case PERM_EDIT_SUPPLIERS
            HasPermission = (strRole = "ADMIN" Or strRole = "MANAGER")

        Case PERM_DELETE_SUPPLIERS
            HasPermission = (strRole = "ADMIN")

        ' SERVICE PERMISSIONS
        Case PERM_CREATE_SERVICES
            HasPermission = True
            
        Case PERM_EDIT_ACTIVE_SERVICES
            HasPermission = True
            
        Case PERM_EDIT_COMPLETED_SERVICES
            HasPermission = (strRole = "ADMIN" Or strRole = "MANAGER")

        Case PERM_DELETE_SERVICES
            HasPermission = (strRole = "ADMIN")

        ' DOCUMENT GENERATION
        Case PERM_GENERATE_LOADING_ORDERS
            HasPermission = True
            
        Case PERM_GENERATE_INVOICES
            HasPermission = (strRole = "ADMIN" Or strRole = "MANAGER")

        Case PERM_VIEW_INVOICES
            HasPermission = (strRole = "ADMIN" Or strRole = "MANAGER")

        Case PERM_VIEW_FINANCIAL_REPORTS
            HasPermission = (strRole = "ADMIN" Or strRole = "MANAGER")

        ' ADMIN-ONLY FUNCTIONS
        Case PERM_MANAGE_USERS
            HasPermission = (strRole = "ADMIN")

        Case PERM_VIEW_AUDIT_LOG
            HasPermission = (strRole = "ADMIN")

        Case PERM_EDIT_COMPANY_SETTINGS
            HasPermission = (strRole = "ADMIN")


        ' UNKNOWN PERMISSION ? DENY (Fail-closed security)
        Case Else
            modUtilities.LogError MODULE_NAME & ".HasPermission", 1003, _
                "Unknown permission requested", "Permission=" & strPermission & " | User=" & modGlobals.Username
            HasPermission = False
    End Select

    Exit Function

ErrorHandler:
    modUtilities.LogError MODULE_NAME & ".HasPermission", Err.number, Err.Description, _
                         "Permission=" & strPermission
    HasPermission = False
End Function

' Sub: RequirePermission � Convenience wrapper that shows message + cancels
' Purpose: Call at top of forms/reports/buttons to enforce access
'          Shows friendly message and cancels if denied

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

        modAudit.LogAudit "Security", 0, "PermissionDenied", _
                          modGlobals.Username & " (" & modGlobals.UserRole & ")", strPermission, "AccessAttempt", modGlobals.UserID, True

        DoCmd.CancelEvent
    End If
End Sub

' GetRoleString: Map UserRole Enum to String (for DB comparison)
Private Function GetRoleString(ByVal enumRole As UserRole) As String

    Select Case enumRole
        Case roleAdmin: GetRoleString = "ADMIN"
        Case roleManager: GetRoleString = "MANAGER"
        Case roleOperator: GetRoleString = "OPERATOR"
        Case Else: GetRoleString = ""
    End Select

End Function

Attribute VB_Name = "modPermissions"
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


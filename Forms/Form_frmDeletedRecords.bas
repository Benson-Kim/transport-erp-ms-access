'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form      : frmDeletedRecords
' Purpose   : Admin-only form to view, restore, or permanently delete archived records
' Author    : Tifamovs Saeol <tifamovs@gmail.com> [Benson Kimathi]
' Date      : 2025-11-30
' Security  : Only accessible to Admin role
' Features  : - View deleted Clients, Suppliers, and Services
'             - Restore archived records with audit logging
'             - Permanent deletion (Admin only, requires typed confirmation)
'             - Dependency checking before restoration
'             - Search and filter capabilities
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Database
Option Explicit
' MODULE-LEVEL VARIABLES

Private m_strEntityType As String
Private m_blnLoading As Boolean



' FORM EVENTS - INITIALIZATION & SECURITY
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo ErrorHandler
    
    ' CRITICAL: Admin-only access
    If Not modPermissions.HasPermission(modPermissions.PERM_VIEW_AUDIT_LOG) Then
        MsgBox "This feature requires Administrator privileges.", _
            vbCritical, modGlobals.APP_NAME & " - Access Denied"
        Cancel = True
        Exit Sub
    End If
    
    Exit Sub
    
ErrorHandler:
    modUtilities.LogError "frmDeletedRecords.Form_Open", Err.number, Err.Description
    Cancel = True
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    m_blnLoading = True
    
    Me.Caption = "Archived Records Manager - Admin Only"
    Me.AutoHeader.Caption = "Archived Records - Recovery & Permanent Deletion"
    
    ' Populate entity type dropdown
    Me.cboEntity.value = ""
    
    Me.txtSearch.value = Null
    Me.txtDeletedFrom.value = Null
    Me.txtDeletedTo.value = Null
    
    Call UpdateButtonStates
    Call UpdateStatistics
    
    m_blnLoading = False
    
    m_strEntityType = "Clients"
    Call RefreshSubform
    
    Exit Sub
    
ErrorHandler:
    modUtilities.LogError "frmDeletedRecords.Form_Load", Err.number, Err.Description
    m_blnLoading = False
End Sub


' ENTITY TYPE SELECTION
Private Sub cboEntity_AfterUpdate()
    On Error GoTo ErrorHandler
    
    If m_blnLoading Then Exit Sub
    
    m_strEntityType = Nz(Me.cboEntity.value, "Clients")
    
    Me.txtSearch.value = Null
    
    ' Refresh display
    Call RefreshSubform
    Call UpdateButtonStates
    Call UpdateStatistics
    
    Exit Sub
    
ErrorHandler:
    modUtilities.LogError "frmDeletedRecords.cboEntity_AfterUpdate", Err.number, Err.Description
End Sub


' SUBFORM SQL GENERATION
Private Function BuildSubformSQL() As String
    On Error GoTo ErrorHandler
    
    Dim strSQL As String
    Dim strWhere As String
    Dim pkField As String
    Dim nameField As String
    Dim emailField As String
    
    If Len(Trim(m_strEntityType)) = 0 Then Exit Function
    
    strWhere = "IsDeleted = True"
    
    pkField = modRecordDuplicator.GetPrimaryKeyField(m_strEntityType)
    
    If Not IsNull(Me.txtSearch) And Len(Trim(Me.txtSearch)) > 0 Then
        Dim strSearch As String
        strSearch = Replace(Trim(Me.txtSearch), "'", "''")
        
        Select Case m_strEntityType
            Case "Clients"
                strWhere = strWhere & " AND (ClientName LIKE '*" & strSearch & "*' OR " & _
                           "VATNumber LIKE '*" & strSearch & "*' OR " & _
                           "Telephone LIKE '*" & strSearch & "*')"
            Case "Suppliers"
                strWhere = strWhere & " AND (SupplierName LIKE '*" & strSearch & "*' OR " & _
                           "VATNumber LIKE '*" & strSearch & "*' OR " & _
                           "Telephone LIKE '*" & strSearch & "*')"
            Case "Services"
                strWhere = strWhere & " AND (ServiceNumber LIKE '*" & strSearch & "*' OR " & _
                           "ClientName LIKE '*" & strSearch & "*')"
        End Select
    End If
    
    If Not IsNull(Me.txtDeletedFrom) Then
        strWhere = strWhere & " AND ModifiedDate >= #" & Format(Me.txtDeletedFrom, "mm/dd/yyyy") & "#"
    End If
    
    If Not IsNull(Me.txtDeletedTo) Then
        strWhere = strWhere & " AND ModifiedDate <= #" & Format(Me.txtDeletedTo, "mm/dd/yyyy") & " 23:59:59#"
    End If
    
    Select Case m_strEntityType
        Case "Clients"
            nameField = "ClientName"
            emailField = "EmailBilling"
            
            strSQL = "SELECT " & pkField & ", " & nameField & ", " & _
                     "VATNumber, Country, Telephone, " & emailField & ", " & _
                     "Format(ModifiedDate, 'DD/MM/YYYY HH:NN') AS [Deleted Date], " & _
                     "DLookup('FullName','Users','UserID=' & [ModifiedBy]) AS [Deleted By], " & _
                     "DCount('*','Services','ClientID=' & " & pkField & " AND IsDeleted=False) AS [Active Services] " & _
                     "FROM " & m_strEntityType & " " & _
                     "WHERE " & strWhere & " " & _
                     "ORDER BY ModifiedDate DESC"
        
        Case "Suppliers"
            nameField = "SupplierName"
            emailField = "Email"
            
            strSQL = "SELECT " & pkField & ", " & nameField & ", " & _
                     "VATNumber, Country, Telephone, " & emailField & ", " & _
                     "Format(IRPFPercentage, '0.00') & '%' AS [IRPF %], " & _
                     "Format(ModifiedDate, 'DD/MM/YYYY HH:NN') AS [Deleted Date], " & _
                     "DLookup('FullName','Users','UserID=' & [ModifiedBy]) AS [Deleted By], " & _
                     "DCount('*','Services','SupplierID=' & " & pkField & " AND IsDeleted=False) AS [Active Services] " & _
                     "FROM " & m_strEntityType & " " & _
                     "WHERE " & strWhere & " " & _
                     "ORDER BY ModifiedDate DESC"
        
        Case "Services"
            strSQL = "SELECT ServiceID, ServiceNumber, " & _
                     "DLookup('ClientName','Clients','ClientID=' & [ClientID]) AS [Client], " & _
                     "DLookup('SupplierName','Suppliers','SupplierID=' & [SupplierID]) AS [Supplier], " & _
                     "Format(ServiceDate, 'DD/MM/YYYY') AS [Service Date], " & _
                     "Format(TotalCost, 'Currency') AS [Cost], " & _
                     "ServiceStatus AS [Status], " & _
                     "Format(ModifiedDate, 'DD/MM/YYYY HH:NN') AS [Deleted Date], " & _
                     "DLookup('FullName','Users','UserID=' & [ModifiedBy]) AS [Deleted By] " & _
                     "FROM Services " & _
                     "WHERE " & strWhere & " " & _
                     "ORDER BY ModifiedDate DESC"
        
        Case Else
            modUtilities.LogError "frmDeletedRecords.BuildSubformSQL", 1003, _
                "Invalid entity type: " & m_strEntityType
            strSQL = ""
    End Select
    
    BuildSubformSQL = strSQL
    Exit Function
    
ErrorHandler:
    modUtilities.LogError "frmDeletedRecords.BuildSubformSQL", Err.number, Err.Description
    BuildSubformSQL = ""
End Function


' SUBFORM REFRESH
Private Sub RefreshSubform()
    On Error GoTo ErrorHandler
    
    Dim strSQL As String
    strSQL = BuildSubformSQL()
    
    If Len(strSQL) > 0 Then
        modDatabase.UpdateQuerySQL "qryDeletedRecords", strSQL
        
        Me.subDeletedList.SourceObject = "Query.qryDeletedRecords"
        
        With Me.subDeletedList.Form
            .AllowAdditions = False
            .AllowDeletions = False
            .AllowEdits = False
            .RecordSelectors = True
        End With
        
    Else
        modDatabase.UpdateQuerySQL "qryDeletedRecords", "SELECT 'No records found' AS [Message]"
        
        Me.subDeletedList.SourceObject = "Query.qryDeletedRecords"
    End If
    
    Exit Sub
    
ErrorHandler:
    modUtilities.LogError "frmDeletedRecords.RefreshSubform", Err.number, Err.Description
    MsgBox "Error refreshing display: " & Err.Description, vbExclamation, modGlobals.APP_NAME
End Sub


' SEARCH & FILTER
Private Sub txtSearch_Change()
    RestartSearchTimer Me  ' 600ms debounce
End Sub

Private Sub Form_Timer()
    Me.TimerInterval = 0
    Call RefreshSubform
End Sub

Private Sub txtDeletedFrom_AfterUpdate()
    RestartSearchTimer Me
End Sub

Private Sub txtDeletedTo_AfterUpdate()
    RestartSearchTimer Me
End Sub

Private Sub cmdClearFilters_Click()
    Me.txtSearch.value = Null
    Me.txtDeletedFrom.value = Null
    Me.txtDeletedTo.value = Null
    Call RefreshSubform
End Sub


' RESTORE FUNCTIONALITY
Private Sub cmdRestore_Click()
    On Error GoTo ErrorHandler
    
    Dim lngRecordID As Long
    lngRecordID = GetSelectedRecordID()
    
    If lngRecordID = 0 Then
        MsgBox "Please select a record to restore.", vbInformation, modGlobals.APP_NAME
        Exit Sub
    End If
    
    ' Get record name for confirmation
    Dim strRecordName As String
    strRecordName = GetRecordName(lngRecordID)
    
    ' Confirm restoration
    If MsgBox("Restore this " & LCase(m_strEntityType) & "?" & vbCrLf & vbCrLf & _
              strRecordName & vbCrLf & vbCrLf & _
              "The record will be reactivated and visible in the main system.", _
              vbQuestion + vbYesNo, "Confirm Restore") = vbNo Then
        Exit Sub
    End If
    
    ' Check for conflicts (VAT number uniqueness)
    If Not CheckRestoreConflicts(lngRecordID) Then Exit Sub
    
    ' Perform restoration
    If RestoreRecord(lngRecordID) Then
        MsgBox "Record restored successfully!", vbInformation, modGlobals.APP_NAME
        
        ' Refresh display
        Call RefreshSubform
        Call UpdateStatistics
    End If
    
    Exit Sub
    
ErrorHandler:
    modUtilities.LogError "frmDeletedRecords.cmdRestore_Click", Err.number, Err.Description
    MsgBox "Error restoring record: " & Err.Description, vbCritical, modGlobals.APP_NAME
End Sub

Private Function RestoreRecord(lngRecordID As Long) As Boolean
    On Error GoTo ErrorHandler
    
    RestoreRecord = False
    
    Dim pkField As String
    pkField = modRecordDuplicator.GetPrimaryKeyField(m_strEntityType)
    
    modDatabase.BeginTransaction
    
    modDatabase.ExecuteQuery _
        "UPDATE " & m_strEntityType & " " & _
        "SET IsDeleted = False, " & _
        "ModifiedBy = " & modGlobals.UserID & ", " & _
        "ModifiedDate = Now() " & _
        "WHERE " & pkField & " = " & lngRecordID
        
    modAudit.LogAudit m_strEntityType, lngRecordID, "IsDeleted", True, False, _
        "Restore", modGlobals.UserID, True
    
    modDatabase.CommitTransaction
    
    RestoreRecord = True
    Exit Function
    
ErrorHandler:
    modDatabase.RollbackTransaction
    modUtilities.LogError "frmDeletedRecords.RestoreRecord", Err.number, Err.Description, _
        "EntityType=" & m_strEntityType & " | ID=" & lngRecordID
    RestoreRecord = False
End Function

Private Function CheckRestoreConflicts(lngRecordID As Long) As Boolean
    On Error GoTo ErrorHandler
    
    CheckRestoreConflicts = True
    
    ' Only check VAT uniqueness for Clients and Suppliers
    If m_strEntityType <> "Clients" And m_strEntityType <> "Suppliers" Then Exit Function
    
    Dim pkField As String
    pkField = modRecordDuplicator.GetPrimaryKeyField(m_strEntityType)
    
    ' Get VAT number of record to restore
    Dim strVAT As String
    strVAT = Nz(DLookup("VATNumber", m_strEntityType, pkField & " = " & lngRecordID), "")
    
    If Len(strVAT) = 0 Then Exit Function
    
    ' Check if another active record has same VAT
    Dim lngConflictID As Variant
    lngConflictID = DLookup(pkField, m_strEntityType, _
        "VATNumber = '" & Replace(strVAT, "'", "''") & "' AND " & _
        "IsDeleted = False AND " & _
        pkField & " <> " & lngRecordID)
    
    If Not IsNull(lngConflictID) Then
        Dim strConflictName As String
        Dim nameField As String
        
        nameField = IIf(m_strEntityType = "Clients", "ClientName", "SupplierName")
        strConflictName = Nz(DLookup(nameField, m_strEntityType, pkField & " = " & lngConflictID), "")
        
        MsgBox "Cannot restore: VAT Number " & strVAT & " is already in use by:" & vbCrLf & vbCrLf & _
               strConflictName & " (ID: " & lngConflictID & ")" & vbCrLf & vbCrLf & _
               "Please resolve the conflict first.", _
               vbExclamation, modGlobals.APP_NAME
        
        CheckRestoreConflicts = False
    End If
    
    Exit Function
    
ErrorHandler:
    modUtilities.LogError "frmDeletedRecords.CheckRestoreConflicts", Err.number, Err.Description
    CheckRestoreConflicts = False
End Function


' PERMANENT DELETE FUNCTIONALITY (Admin Only)
Private Sub cmdDelete_Click()
    On Error GoTo ErrorHandler
    
    ' Double-check admin permission
    If Not modPermissions.HasPermission(modPermissions.PERM_MANAGE_USERS) Then
        MsgBox "Permanent deletion requires Administrator privileges.", _
            vbCritical, modGlobals.APP_NAME & " - Access Denied"
        Exit Sub
    End If
    
    ' Get selected record
    Dim lngRecordID As Long
    lngRecordID = GetSelectedRecordID()
    
    If lngRecordID = 0 Then
        MsgBox "Please select a record to permanently delete.", vbInformation, modGlobals.APP_NAME
        Exit Sub
    End If
    
    Dim strRecordName As String
    strRecordName = GetRecordName(lngRecordID)
    
    ' Check dependencies
    Dim strDependencies As String
    strDependencies = CheckDependencies(lngRecordID)
    
    If Len(strDependencies) > 0 Then
        MsgBox "Cannot permanently delete this record due to dependencies:" & vbCrLf & vbCrLf & _
               strDependencies & vbCrLf & vbCrLf & _
               "Please delete or reassign dependent records first.", _
               vbExclamation, modGlobals.APP_NAME
        Exit Sub
    End If
    
    ' WARNING MESSAGE
    If MsgBox("? PERMANENT DELETION WARNING ?" & vbCrLf & vbCrLf & _
              "You are about to PERMANENTLY DELETE:" & vbCrLf & _
              strRecordName & vbCrLf & vbCrLf & _
              "This action:" & vbCrLf & _
              "� CANNOT be undone" & vbCrLf & _
              "� Will REMOVE ALL AUDIT HISTORY" & vbCrLf & _
              "� Is IRREVERSIBLE" & vbCrLf & vbCrLf & _
              "Are you ABSOLUTELY SURE?", _
              vbCritical + vbYesNo + vbDefaultButton2, _
              "PERMANENT DELETION") = vbNo Then
        Exit Sub
    End If
    
    ' Require typed confirmation
    Dim strConfirm As String
    strConfirm = InputBox( _
        "Type PERMANENTLY DELETE in capital letters to confirm:" & vbCrLf & vbCrLf & _
        "Record: " & strRecordName, _
        "Final Confirmation Required")
    
    If StrComp(strConfirm, "PERMANENTLY DELETE", vbBinaryCompare) <> 0 Then
        MsgBox "Permanent deletion cancelled.", vbInformation, modGlobals.APP_NAME
        Exit Sub
    End If
    
    ' Perform permanent deletion
    If PermanentlyDeleteRecord(lngRecordID, strRecordName) Then
        MsgBox "Record permanently deleted.", vbInformation, modGlobals.APP_NAME
        
        ' Refresh display
        Call RefreshSubform
        Call UpdateStatistics
    End If
    
    Exit Sub
    
ErrorHandler:
    modUtilities.LogError "frmDeletedRecords.cmdDelete_Click", Err.number, Err.Description
    MsgBox "Error during permanent deletion: " & Err.Description, vbCritical, modGlobals.APP_NAME
End Sub

Private Function PermanentlyDeleteRecord(lngRecordID As Long, strRecordName As String) As Boolean
    On Error GoTo ErrorHandler
    
    PermanentlyDeleteRecord = False
    
    Dim pkField As String
    pkField = modRecordDuplicator.GetPrimaryKeyField(m_strEntityType)
    
    ' Begin transaction
    modDatabase.BeginTransaction
    
    ' Log before deletion (audit will be lost after delete)
    modAudit.LogAudit m_strEntityType, lngRecordID, "", Null, _
        "PERMANENTLY DELETED: " & strRecordName, "PermanentDelete", modGlobals.UserID, True
    
    ' Hard delete from database
    CurrentDb.Execute _
        "DELETE FROM " & m_strEntityType & " " & _
        "WHERE " & pkField & " = " & lngRecordID, _
        dbFailOnError
    
    modDatabase.CommitTransaction
    
    PermanentlyDeleteRecord = True
    Exit Function
    
ErrorHandler:
    modDatabase.RollbackTransaction
    modUtilities.LogError "frmDeletedRecords.PermanentlyDeleteRecord", Err.number, Err.Description
    PermanentlyDeleteRecord = False
End Function


' HELPER FUNCTIONS
Private Function GetSelectedRecordID() As Long
    On Error Resume Next
    
    Dim pkField As String
    pkField = modRecordDuplicator.GetPrimaryKeyField(m_strEntityType)
    
    GetSelectedRecordID = 0
    If Me.subDeletedList.Form.RecordsetClone.RecordCount = 0 Then Exit Function
       
    GetSelectedRecordID = Nz(Me.subDeletedList.Form(pkField).value, 0)
End Function

Public Sub SubformRowChanged()
    On Error Resume Next
    
    Dim pkField As String
    pkField = modRecordDuplicator.GetPrimaryKeyField(m_strEntityType)

    Dim selectedID As Long
    selectedID = Nz(Me.subDeletedList.Form(pkField).value, 0)

End Sub

Private Function GetRecordName(lngRecordID As Long) As String
    On Error Resume Next
    
    Dim nameField As String
    Dim pkField As String
    
    pkField = modRecordDuplicator.GetPrimaryKeyField(m_strEntityType)
    
    Select Case m_strEntityType
        Case "Clients"
            nameField = "ClientName"
        Case "Suppliers"
            nameField = "SupplierName"
        Case "Services"
            nameField = "ServiceNumber"
        Case Else
            GetRecordName = "[Unknown]"
            Exit Function
    End Select
    
    GetRecordName = Nz(DLookup(nameField, m_strEntityType, pkField & " = " & lngRecordID), "[Not Found]")
End Function

Private Function CheckDependencies(lngRecordID As Long) As String
    On Error Resume Next
    
    Dim strResult As String
    Dim lngCount As Long
    
    Select Case m_strEntityType
        Case "Clients"
            ' Check for services
            lngCount = DCount("*", "Services", "ClientID = " & lngRecordID)
            If lngCount > 0 Then
                strResult = "� " & lngCount & " service(s) linked to this client"
            End If
            
        Case "Suppliers"
            ' Check for services
            lngCount = DCount("*", "Services", "SupplierID = " & lngRecordID)
            If lngCount > 0 Then
                strResult = "� " & lngCount & " service(s) linked to this supplier"
            End If
            
        Case "Services"
            ' Services typically have no dependencies
            strResult = ""
    End Select
    
    CheckDependencies = strResult
End Function

Private Sub UpdateButtonStates()
    On Error Resume Next
    
    Dim hasSelection As Boolean
    hasSelection = (GetSelectedRecordID() > 0)
    
    Me.cmdRestore.Enabled = hasSelection
    Me.cmdDelete.Enabled = hasSelection And _
        modPermissions.HasPermission(modPermissions.PERM_MANAGE_USERS)
    
    ' Update button colors
    If Me.cmdDelete.Enabled Then
        Me.cmdDelete.ForeColor = vbWhite
    Else
        Me.cmdDelete.ForeColor = vbGrayText
    End If
End Sub

Private Sub UpdateStatistics()
    On Error Resume Next
    
    Dim strCriteria As String
    strCriteria = "IsDeleted = True"
    
     With Me
        .txtCountClients.value = CachedDCount("*", "Clients", strCriteria)
        .txtCountSuppliers.value = CachedDCount("*", "Suppliers", strCriteria)
        .txtCountServices.value = CachedDCount("*", "Services", strCriteria)
    End With
    
End Sub


' SUBFORM EVENTS
Private Sub subDeletedList_Click()
    Call UpdateButtonStates
End Sub

Private Sub subDeletedList_DblClick(Cancel As Integer)
    ' Double-click to restore
    Call cmdRestore_Click
End Sub


' EXPORT FUNCTIONALITY
Private Sub cmdExport_Click()
    On Error GoTo ErrorHandler
    
    Dim strPath As String
    strPath = CurrentProject.Path & "\ArchivedRecords_" & _
              m_strEntityType & "_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsx"
    
    If modDatabase.ExportToExcel("qryDeletedRecords", strPath, , False) Then
        MsgBox "Archived records exported to:" & vbCrLf & vbCrLf & strPath, _
            vbInformation, modGlobals.APP_NAME
    End If
    
    Exit Sub
    
ErrorHandler:
    modUtilities.LogError "frmDeletedRecords.cmdExport_Click", Err.number, Err.Description
End Sub


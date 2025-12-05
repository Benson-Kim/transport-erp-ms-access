'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MODULE    : modClientSupplierForms
' PURPOSE   : Centralized business logic for Client and Supplier management forms
'               Implements DRY principles by providing shared validation, duplication,
'                deletion, and utility functions for both entity types
' AUTHOR    : Tifamovs Saeol <tifamovs@gmail.com> [Benson Kimathi]
' DATE      : 2025-11-30
' SECURITY  : All functions include permission checks and comprehensive error handling
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Database
Option Explicit

Private Const MODULE_NAME As String = "modClientSupplierForms"


' '''''''''''''''''' MAIN VALIDATION FUNCTION ' ''''''''''''''''''

Public Function ValidateEntityData( _
    EntityType As String, _
    EntityID As Variant, _
    EntityName As String, _
    VATNumber As String, _
    Email As String, _
    Telephone As String, _
    Address As String, _
    Country As String, _
    Optional IRPFPercentage As Variant, _
    Optional ContactName As String, _
    Optional City As String, _
    Optional ZIPCode As String) As ValidationResult
    
    On Error GoTo ErrorHandler
    
    Dim result As ValidationResult
    Set result.ErrorMessages = New Collection
    Set result.FieldsWithErrors = New Collection
    result.isValid = True
    
    Dim config As EntityConfig
    config = GetEntityConfig(EntityType)
    
    ' '''''''''''''''''' REQUIRED FIELD VALIDATIONS ''''''''''''''''''
    ' Name validation
    If Len(Trim(EntityName)) < 2 Then
        result.ErrorMessages.Add "Name must be at least 2 characters"
        result.FieldsWithErrors.Add config.nameField
        result.isValid = False
    ElseIf Len(Trim(EntityName)) > 100 Then
        result.ErrorMessages.Add "Name cannot exceed 100 characters"
        result.FieldsWithErrors.Add config.nameField
        result.isValid = False
    End If
    
    ' VAT Number validation
    If Len(Trim(VATNumber)) = 0 Then
        result.ErrorMessages.Add "VAT Number is required"
        result.FieldsWithErrors.Add "VATNumber"
        result.isValid = False
    Else
        ' Check uniqueness
        If Not ValidateVATUniqueness(config.tableName, config.pkField, EntityID, VATNumber) Then
            result.ErrorMessages.Add "VAT Number already exists for another " & LCase(EntityType)
            result.FieldsWithErrors.Add "VATNumber"
            result.isValid = False
        End If
        
        ' Format validation for Spanish VAT
        If Not modValidation.IsValidSpanishVAT(VATNumber) Then
            result.ErrorMessages.Add "Invalid Spanish VAT format. Expected: ES + letter/digit + 8 digits + optional letter" & vbCrLf & _
                                   "Example: ESA12345678 or ES12345678A"
            result.FieldsWithErrors.Add "VATNumber"
            result.isValid = False
        End If
    End If
    
    ' Email validation
    If Len(Trim(Email)) = 0 Then
        result.ErrorMessages.Add "Email is required"
        result.FieldsWithErrors.Add config.emailField
        result.isValid = False
    ElseIf Not modValidation.IsValidEmail(Email) Then
        result.ErrorMessages.Add "Invalid email address format"
        result.FieldsWithErrors.Add config.emailField
        result.isValid = False
    End If
    
    ' Telephone validation
    If Len(Trim(Telephone)) = 0 Then
        result.ErrorMessages.Add "Telephone is required"
        result.FieldsWithErrors.Add "Telephone"
        result.isValid = False
    ElseIf Not ValidateTelephone(Telephone) Then
        result.ErrorMessages.Add "Telephone must contain at least one digit and only valid characters (+, -, (, ), space, digits)"
        result.FieldsWithErrors.Add "Telephone"
        result.isValid = False
    End If
    
    ' Address validation
    If Len(Trim(Address)) < 5 Then
        result.ErrorMessages.Add "Address must be at least 5 characters"
        result.FieldsWithErrors.Add Address
        result.isValid = False
    ElseIf Len(Trim(Address)) > 255 Then
        result.ErrorMessages.Add "Address cannot exceed 255 characters"
        result.FieldsWithErrors.Add Address
        result.isValid = False
    End If
    
    ' Address validation
    If Len(Trim(City)) = 0 Then
        result.ErrorMessages.Add "City is required"
        result.FieldsWithErrors.Add "City"
        result.isValid = False
    End If
    
    ' Zipcode validation
    If Len(Trim(ZIPCode)) = 0 Then
        result.ErrorMessages.Add "ZIPCode is required"
        result.FieldsWithErrors.Add "ZIPCode"
        result.isValid = False
    End If
    
    ' Country validation
    If Len(Trim(Country)) = 0 Then
        result.ErrorMessages.Add "Country is required"
        result.FieldsWithErrors.Add "Country"
        result.isValid = False
    End If
    
    ' '''''''''''''''''' OPTIONAL FIELD VALIDATIONS ''''''''''''''''''
    
    ' City validation (if provided)
    If Not IsMissing(City) And Len(Trim(City)) > 0 Then
        If Len(Trim(City)) > 100 Then
            result.ErrorMessages.Add "City cannot exceed 100 characters"
            result.FieldsWithErrors.Add "City"
            result.isValid = False
        End If
    End If
    
    ' Zip Code validation (if provided)
    If Not IsMissing(ZIPCode) And Len(Trim(ZIPCode)) > 0 Then
        If Not ValidateZipCode(ZIPCode, Country) Then
            result.ErrorMessages.Add "Invalid postal code format for " & Country
            result.FieldsWithErrors.Add "ZipCode"
            result.isValid = False
        End If
    End If
    
    ' '''''''''''''''''' SUPPLIER-SPECIFIC VALIDATIONS ''''''''''''''''''
    
    If EntityType = "Supplier" Then
        ' IRPF percentage validation
        If Not IsNull(IRPFPercentage) Then
            Dim dblIRPF As Double
            dblIRPF = CDbl(Nz(IRPFPercentage, 0))
            
            If dblIRPF < 0 Or dblIRPF > 15 Then
                result.ErrorMessages.Add "IRPF percentage must be between 0% and 15%"
                result.FieldsWithErrors.Add "IRPFPercentage"
                result.isValid = False
            End If
        End If
    End If
    
    ValidateEntityData = result
    Exit Function
    
ErrorHandler:
    modUtilities.LogError MODULE_NAME & ".ValidateEntityData", _
        Err.number, Err.Description, "EntityType=" & EntityType
    
    result.isValid = False
    result.ErrorMessages.Add "Validation error: " & Err.Description
    ValidateEntityData = result
End Function

' ''''''''''''''''''  ENTITY DUPLICATION FUNCTION ' ''''''''''''''''''
Public Function DuplicateEntity( _
    EntityType As String, _
    EntityID As Long, _
    ByRef NewEntityID As Long) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim config As EntityConfig
    Dim dictExclude As Object
    Dim strError As String
    
    config = GetEntityConfig(EntityType)
    
    
    ' Validate permissions
    Dim strPermission As String
    strPermission = IIf(EntityType = "Client", modPermissions.PERM_EDIT_CLIENTS, modPermissions.PERM_EDIT_SUPPLIERS)
    
    If Not modPermissions.HasPermission(strPermission) Then
        MsgBox "You do not have permission to duplicate " & LCase(EntityType) & "s.", _
            vbExclamation, modGlobals.APP_NAME
        DuplicateEntity = False
        NewEntityID = 0
        Exit Function
    End If
    
    NewEntityID = modRecordDuplicator.SafeDuplicate(config.tableName, EntityID, strError)
    
    If NewEntityID = 0 Or IsNull(NewEntityID) Then
        MsgBox "Duplication failed: " & strError, vbCritical, modGlobals.APP_NAME
        DuplicateEntity = False
        NewEntityID = 0
        Exit Function
    End If
    
    DuplicateEntity = True
    Exit Function
    
ErrorHandler:
    modDatabase.RollbackTransaction
    modUtilities.LogError MODULE_NAME & ".DuplicateEntity", _
        Err.number, Err.Description, "EntityType=" & EntityType & " | ID=" & EntityID
    MsgBox "Error duplicating " & LCase(EntityType) & ": " & Err.Description, vbCritical, modGlobals.APP_NAME
    DuplicateEntity = False
    NewEntityID = 0
End Function

' '''''''''''''''''' SAFE DELETION FUNCTION WITH DEPENDENCY CHECKING ' ''''''''''''''''''

Public Function SafeDeleteEntity( _
    EntityType As String, _
    EntityID As Long, _
    EntityName As String) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim config As EntityConfig
    Dim dependencyCount As Long
    Dim strConfirm As String
    Dim strDependencies As String
    
    config = GetEntityConfig(EntityType)
    
    ' Validate permissions
    Dim strPermission As String
    strPermission = IIf(EntityType = "Client", _
        modPermissions.PERM_DELETE_CLIENTS, _
        modPermissions.PERM_DELETE_SUPPLIERS)
    
    If Not modPermissions.HasPermission(strPermission) Then
        MsgBox "You do not have permission to delete " & LCase(EntityType) & "s.", _
            vbExclamation, modGlobals.APP_NAME
        SafeDeleteEntity = False
        Exit Function
    End If
    
    ' Check for active services
    dependencyCount = DCount("*", "Services", _
        config.pkField & " = " & EntityID & " AND IsDeleted = False")
    
    If dependencyCount > 0 Then
        ' Get sample service numbers for display
        strDependencies = GetServiceNumbers(config.pkField, EntityID, 5)
        
        Dim response As VbMsgBoxResult
        response = MsgBox( _
            "This " & LCase(EntityType) & " has " & dependencyCount & " active service(s):" & vbCrLf & vbCrLf & _
            strDependencies & vbCrLf & vbCrLf & _
            "SOFT DELETION will:" & vbCrLf & _
            "� Mark the " & LCase(EntityType) & " as inactive" & vbCrLf & _
            "� Preserve all historical data and audit trail" & vbCrLf & _
            "� Keep services linked for reporting" & vbCrLf & vbCrLf & _
            "The " & LCase(EntityType) & " can be restored later by an Administrator." & vbCrLf & vbCrLf & _
            "Continue with soft deletion?", _
            vbQuestion + vbYesNo + vbDefaultButton2 + vbExclamation, _
            "Confirm Deletion - Active Services Exist")
        
        If response = vbNo Then
            SafeDeleteEntity = False
            Exit Function
        End If
    End If
    
    ' Require typed confirmation
    strConfirm = InputBox( _
        "This action will ARCHIVE:" & vbCrLf & vbCrLf & _
        EntityName & vbCrLf & vbCrLf & _
        "The record will be hidden from normal views but preserved for audit purposes." & vbCrLf & vbCrLf & _
        "Type DELETE in capital letters to confirm:", _
        "Confirm Soft Deletion", "")
    
    If StrComp(strConfirm, "DELETE", vbBinaryCompare) <> 0 Then
        MsgBox "Deletion cancelled.", vbInformation, modGlobals.APP_NAME
        SafeDeleteEntity = False
        Exit Function
    End If
    
    ' Perform soft delete within transaction
    modDatabase.BeginTransaction
    
    Dim strSQL As String
    strSQL = "UPDATE " & config.tableName & " SET " & _
             "IsDeleted = True, " & _
             "ModifiedBy = " & modGlobals.UserID & ", " & _
             "ModifiedDate = Now() " & _
             "WHERE " & config.pkField & " = " & EntityID
    
    CurrentDb.Execute strSQL, dbFailOnError
    
    ' Log deletion
    modAudit.LogDelete config.tableName, EntityID, modGlobals.UserID
    
    modDatabase.CommitTransaction
    
    MsgBox EntityType & " '" & EntityName & "' has been archived successfully." & vbCrLf & vbCrLf & _
           "An Administrator can restore it from the Deleted Records form.", _
           vbInformation, modGlobals.APP_NAME & " - Deletion Complete"
    
    SafeDeleteEntity = True
    Exit Function
    
ErrorHandler:
    modDatabase.RollbackTransaction
    modUtilities.LogError MODULE_NAME & ".SafeDeleteEntity", _
        Err.number, Err.Description, "EntityType=" & EntityType & " | ID=" & EntityID
    MsgBox "Error deleting " & LCase(EntityType) & ": " & Err.Description, _
        vbCritical, modGlobals.APP_NAME
    SafeDeleteEntity = False
End Function

' '''''''''''''''''' RESTORE DELETED ENTITY (ADMIN ONLY) ' ''''''''''''''''''

Public Function RestoreDeletedEntity( _
    EntityType As String, _
    EntityID As Long, _
    EntityName As String) As Boolean
    
    On Error GoTo ErrorHandler
    
    ' Admin-only function
    If Not modPermissions.HasPermission(modPermissions.PERM_MANAGE_USERS) Then
        MsgBox "Only Administrators can restore deleted records.", _
            vbExclamation, modGlobals.APP_NAME
        RestoreDeletedEntity = False
        Exit Function
    End If
    
    Dim config As EntityConfig
    config = GetEntityConfig(EntityType)
    
    ' Confirm restoration
    If MsgBox("Restore " & LCase(EntityType) & ":" & vbCrLf & vbCrLf & _
              EntityName & vbCrLf & vbCrLf & _
              "This will make the record visible and active again.", _
              vbQuestion + vbYesNo, "Confirm Restoration") = vbNo Then
        RestoreDeletedEntity = False
        Exit Function
    End If
    
    ' Perform restoration
    modDatabase.BeginTransaction
    
    Dim strSQL As String
    strSQL = "UPDATE " & config.tableName & " SET " & _
             "IsDeleted = False, " & _
             "ModifiedBy = " & modGlobals.UserID & ", " & _
             "ModifiedDate = Now() " & _
             "WHERE " & config.pkField & " = " & EntityID
    
    CurrentDb.Execute strSQL, dbFailOnError
    
    ' Log restoration
    modAudit.LogAudit config.tableName, EntityID, "IsDeleted", "True", "False", _
        "Restore", modGlobals.UserID, True
    
    modDatabase.CommitTransaction
    
    MsgBox EntityType & " restored successfully.", vbInformation, modGlobals.APP_NAME
    
    RestoreDeletedEntity = True
    Exit Function
    
ErrorHandler:
    modDatabase.RollbackTransaction
    modUtilities.LogError MODULE_NAME & ".RestoreDeletedEntity", _
        Err.number, Err.Description, "EntityType=" & EntityType & " | ID=" & EntityID
    MsgBox "Error restoring " & LCase(EntityType) & ": " & Err.Description, _
        vbCritical, modGlobals.APP_NAME
    RestoreDeletedEntity = False
End Function

' '''''''''''''''''' PERMANENT DELETION (ADMIN ONLY - DANGEROUS) '''''''''''''''''' '
Public Function PermanentlyDeleteEntity( _
    EntityType As String, _
    EntityID As Long, _
    EntityName As String) As Boolean
    
    On Error GoTo ErrorHandler
    
    ' Admin-only function
    If Not modPermissions.HasPermission(modPermissions.PERM_MANAGE_USERS) Then
        MsgBox "Only Administrators can permanently delete records.", _
            vbExclamation, modGlobals.APP_NAME
        PermanentlyDeleteEntity = False
        Exit Function
    End If
    
    Dim config As EntityConfig
    config = GetEntityConfig(EntityType)
    
    ' Check for dependencies
    Dim dependencyCount As Long
    dependencyCount = DCount("*", "Services", config.pkField & " = " & EntityID)
    
    If dependencyCount > 0 Then
        MsgBox "Cannot permanently delete: " & dependencyCount & " service(s) still reference this " & _
               LCase(EntityType) & "." & vbCrLf & vbCrLf & _
               "You must first reassign or delete all related services.", _
               vbCritical, modGlobals.APP_NAME & " - Delete Blocked"
        PermanentlyDeleteEntity = False
        Exit Function
    End If
    
    ' Require strict confirmation
    Dim strConfirm As String
    strConfirm = InputBox( _
        "PERMANENT DELETION WARNING" & vbCrLf & vbCrLf & _
        "You are about to PERMANENTLY DELETE:" & vbCrLf & _
        EntityName & vbCrLf & vbCrLf & _
        "This action:" & vbCrLf & _
        "� CANNOT BE UNDONE" & vbCrLf & _
        "� Will remove ALL audit history" & vbCrLf & _
        "� Will erase ALL data permanently" & vbCrLf & vbCrLf & _
        "Type 'PERMANENTLY DELETE' (exact case) to confirm:", _
        "FINAL WARNING - Permanent Deletion", "")
    
    If StrComp(strConfirm, "PERMANENTLY DELETE", vbBinaryCompare) <> 0 Then
        MsgBox "Permanent deletion cancelled.", vbInformation, modGlobals.APP_NAME
        PermanentlyDeleteEntity = False
        Exit Function
    End If
    
    ' Final confirmation
    If MsgBox("LAST CHANCE TO CANCEL" & vbCrLf & vbCrLf & _
              "Delete '" & EntityName & "' forever?", _
              vbCritical + vbYesNo + vbDefaultButton2, _
              "Final Confirmation") = vbNo Then
        PermanentlyDeleteEntity = False
        Exit Function
    End If
    
    ' Perform hard delete
    modDatabase.BeginTransaction
    
    ' Log before deletion (last audit entry)
    modAudit.LogAudit config.tableName, EntityID, "", Null, _
        "PERMANENTLY DELETED by " & modGlobals.FullName, _
        "Delete", modGlobals.UserID, True
    
    ' Delete from table
    Dim strSQL As String
    strSQL = "DELETE FROM " & config.tableName & " WHERE " & config.pkField & " = " & EntityID
    CurrentDb.Execute strSQL, dbFailOnError
    
    modDatabase.CommitTransaction
    
    MsgBox "Record permanently deleted.", vbInformation, modGlobals.APP_NAME
    
    PermanentlyDeleteEntity = True
    Exit Function
    
ErrorHandler:
    modDatabase.RollbackTransaction
    modUtilities.LogError MODULE_NAME & ".PermanentlyDeleteEntity", _
        Err.number, Err.Description, "EntityType=" & EntityType & " | ID=" & EntityID
    MsgBox "Error permanently deleting " & LCase(EntityType) & ": " & Err.Description, _
        vbCritical, modGlobals.APP_NAME
    PermanentlyDeleteEntity = False
End Function

' '''''''''''''''''' HELPER FUNCTIONS - VALIDATION ''''''''''''''''''
Private Function ValidateVATUniqueness( _
    tableName As String, _
    pkField As String, _
    EntityID As Variant, _
    VATNumber As String) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim strCriteria As String
    strCriteria = "VATNumber = '" & Replace(Trim(VATNumber), "'", "''") & "' " & _
                  "AND " & pkField & " <> " & Nz(EntityID, 0) & " " & _
                  "AND IsDeleted = False"
    
    ValidateVATUniqueness = Not modDatabase.RecordExists(tableName, strCriteria)
    Exit Function
    
ErrorHandler:
    modUtilities.LogError MODULE_NAME & ".ValidateVATUniqueness", Err.number, Err.Description
    ValidateVATUniqueness = False
End Function

Private Function ValidateTelephone(Telephone As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim strClean As String
    strClean = Trim(Telephone)
    
    ' Must contain at least one digit
    If Not strClean Like "*[0-9]*" Then
        ValidateTelephone = False
        Exit Function
    End If
    
    ' Only allow: digits, spaces, +, -, (, )
    Dim i As Integer
    Dim ch As String
    
    For i = 1 To Len(strClean)
        ch = Mid(strClean, i, 1)
        If Not (ch Like "[0-9]" Or ch = " " Or ch = "+" Or ch = "-" Or ch = "(" Or ch = ")") Then
            ValidateTelephone = False
            Exit Function
        End If
    Next i
    
    ValidateTelephone = True
    Exit Function
    
ErrorHandler:
    modUtilities.LogError MODULE_NAME & ".ValidateTelephone", Err.number, Err.Description
    ValidateTelephone = False
End Function

Private Function ValidateZipCode(ZIPCode As String, Country As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim strClean As String
    strClean = Trim(ZIPCode)
    
     ' Basic validation for common formats
     ValidateZipCode = (Len(strClean) >= 3 And Len(strClean) <= 10)
   
    Exit Function
    
ErrorHandler:
    modUtilities.LogError MODULE_NAME & ".ValidateZipCode", Err.number, Err.Description
    ValidateZipCode = True  ' Default to accepting on error
End Function

' '''''''''''''''''' HELPER FUNCTIONS - CONFIGURATION ''''''''''''''''''

Private Function GetEntityConfig(EntityType As String) As EntityConfig
    Dim config As EntityConfig
    
    If EntityType = "Client" Then
        config.EntityType = "Client"
        config.tableName = "Clients"
        config.pkField = "ClientID"
        config.nameField = "ClientName"
        config.emailField = "EmailBilling"
        config.addressField = "Address"
    Else
        config.EntityType = "Supplier"
        config.tableName = "Suppliers"
        config.pkField = "SupplierID"
        config.nameField = "SupplierName"
        config.emailField = "Email"
        config.addressField = "AddressLine"
    End If
    
    GetEntityConfig = config
End Function

Private Function GetServiceNumbers( _
    pkField As String, _
    EntityID As Long, _
    MaxCount As Integer) As String
    
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim strResult As String
    Dim intCount As Integer
    
    Set db = CurrentDb
    
    strSQL = "SELECT TOP " & MaxCount & " ServiceNumber " & _
             "FROM Services " & _
             "WHERE " & pkField & " = " & EntityID & " " & _
             "AND IsDeleted = False " & _
             "ORDER BY ServiceDate DESC"
    
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    
    intCount = 0
    Do While Not rs.EOF And intCount < MaxCount
        strResult = strResult & "  � " & Nz(rs!ServiceNumber, "[No Number]") & vbCrLf
        intCount = intCount + 1
        rs.MoveNext
    Loop
    
    rs.Close
    
    ' Add "and X more..." if exceeded
    Dim lngTotal As Long
    lngTotal = DCount("*", "Services", pkField & " = " & EntityID & " AND IsDeleted = False")
    
    If lngTotal > MaxCount Then
        strResult = strResult & "  ...and " & (lngTotal - MaxCount) & " more"
    End If
    
    GetServiceNumbers = strResult
    Exit Function
    
ErrorHandler:
    modUtilities.LogError MODULE_NAME & ".GetServiceNumbers", Err.number, Err.Description
    GetServiceNumbers = "  [Unable to retrieve service numbers]"
End Function

Public Function ExportClientsToExcel( _
    Optional strFilter As String = "", _
    Optional strOutputPath As String = "") As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim strFileName As String
    Dim strSQL As String
    Dim strTempQuery As String
    Dim db As DAO.Database
    
    ' Generate filename if not provided
    If Len(strOutputPath) = 0 Then
        strFileName = "Suppliers_" & Format(Date, "yyyymmdd") & ".xlsx"
        
        strOutputPath = modUtilities.GetSaveAsPath(strFileName)
        
        If Len(strOutputPath) = 0 Then   ' User cancelled
            ExportSuppliersToExcel = False
            Exit Function
        End If
    End If
    
    ' Build export query with clean column names
    strSQL = "SELECT " & _
             "ClientID AS [Client ID], " & _
             "ClientName AS [Client Name], " & _
             "VATNumber AS [VAT Number], " & _
             "Address, " & _
             "City, " & _
             "PostalCode AS [Postal Code], " & _
             "Country, " & _
             "Telephone, " & _
             "EmailBilling AS [Billing Email], " & _
             "EmailTraffic AS [Traffic Email], " & _
             "PaymentTerms AS [Payment Terms], " & _
             "VATApplied AS [VAT Applied], " & _
             "BankAccount AS [Bank Account], " & _
             "BankDetails AS [Bank Details], " & _
             "Format(CreatedDate, 'DD/MM/YYYY HH:NN') AS [Created Date], " & _
             "Format(ModifiedDate, 'DD/MM/YYYY HH:NN') AS [Modified Date] " & _
             "FROM Clients " & _
             "WHERE IsDeleted = False"
    
    If Len(strFilter) > 0 Then
        strSQL = strSQL & " AND (" & strFilter & ")"
    End If
    
    strSQL = strSQL & " ORDER BY ClientName"
    
    ' Create temporary export query
    Set db = CurrentDb
    strTempQuery = "qryExportClients_Temp"
    
    On Error Resume Next
    db.QueryDefs.Delete strTempQuery
    On Error GoTo ErrorHandler
    
    db.CreateQueryDef strTempQuery, strSQL
    
    ' Export
    DoCmd.TransferSpreadsheet _
        acExport, _
        acSpreadsheetTypeExcel12Xml, _
        strTempQuery, _
        strOutputPath, _
        True
    
    ' Clean up
    db.QueryDefs.Delete strTempQuery
    
    ' Audit log
    modAudit.LogAudit "Clients", 0, "Export", Null, _
        "Exported to: " & strOutputPath, "Export", modGlobals.UserID, True
    
    ' Success message
    If MsgBox("Export completed successfully!" & vbCrLf & vbCrLf & _
              "Exported: " & DCount("*", "Clients", IIf(Len(strFilter) > 0, "IsDeleted = False AND (" & strFilter & ")", "IsDeleted = False")) & " clients" & vbCrLf & _
              "Path: " & strOutputPath & vbCrLf & vbCrLf & _
              "Open the file now?", vbQuestion + vbYesNo, "Export Complete") = vbYes Then
        Application.FollowHyperlink strOutputPath
    End If
    
    ExportClientsToExcel = True
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    db.QueryDefs.Delete strTempQuery
    On Error GoTo 0
    
    modUtilities.LogError MODULE_NAME & ".ExportClientsToExcel", Err.number, Err.Description
    MsgBox "Error exporting clients:" & vbCrLf & vbCrLf & Err.Description, vbCritical, APP_NAME
    ExportClientsToExcel = False
End Function

' ---------------------------------------------------------------------
' FUNCTION: ExportSuppliersToExcel
' PURPOSE: Export supplier list with optional filters
' ---------------------------------------------------------------------
Public Function ExportSuppliersToExcel( _
    Optional strFilter As String = "", _
    Optional strOutputPath As String = "") As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim strFileName As String
    Dim strSQL As String
    Dim strTempQuery As String
    Dim db As DAO.Database
    
    ' Generate filename if not provided
    If Len(strOutputPath) = 0 Then
        strFileName = "Suppliers_" & Format(Date, "yyyymmdd") & ".xlsx"
        
        strOutputPath = modUtilities.GetSaveAsPath(strFileName)
        
        If Len(strOutputPath) = 0 Then   ' User cancelled
            ExportSuppliersToExcel = False
            Exit Function
        End If
    End If
    
    ' Build export query with clean column names
    strSQL = "SELECT " & _
             "SupplierID AS [Supplier ID], " & _
             "SupplierName AS [Supplier Name], " & _
             "VATNumber AS [VAT Number], " & _
             "AddressLine AS [Address], " & _
             "City, " & _
             "ZipCode AS [Zip Code], " & _
             "Country, " & _
             "Telephone, " & _
             "Email, " & _
             "Format(IRPFPercentage, '0.00') & '%' AS [IRPF %], " & _
             "TypeOfServices AS [Service Type], " & _
             "PaymentTerms AS [Payment Terms], " & _
             "VATApplied AS [VAT Applied], " & _
             "BankAccount AS [Bank Account], " & _
             "BankDetails AS [Bank Details], " & _
             "Format(CreatedDate, 'DD/MM/YYYY HH:NN') AS [Created Date], " & _
             "Format(ModifiedDate, 'DD/MM/YYYY HH:NN') AS [Modified Date] " & _
             "FROM Suppliers " & _
             "WHERE IsDeleted = False"
    
    If Len(strFilter) > 0 Then
        strSQL = strSQL & " AND (" & strFilter & ")"
    End If
    
    strSQL = strSQL & " ORDER BY SupplierName"
    
    ' Create temporary export query
    Set db = CurrentDb
    strTempQuery = "qryExportSuppliers_Temp"
    
    On Error Resume Next
    db.QueryDefs.Delete strTempQuery
    On Error GoTo ErrorHandler
    
    db.CreateQueryDef strTempQuery, strSQL
    
    ' Export
    DoCmd.TransferSpreadsheet _
        acExport, _
        acSpreadsheetTypeExcel12Xml, _
        strTempQuery, _
        strOutputPath, _
        True
    
    ' Clean up
    db.QueryDefs.Delete strTempQuery
    
    ' Audit log
    modAudit.LogAudit "Suppliers", 0, "Export", Null, _
        "Exported to: " & strOutputPath, "Export", modGlobals.UserID, True
    
    ' Success message
    If MsgBox("Export completed successfully!" & vbCrLf & vbCrLf & _
              "Exported: " & DCount("*", "Suppliers", IIf(Len(strFilter) > 0, "IsDeleted = False AND (" & strFilter & ")", "IsDeleted = False")) & " suppliers" & vbCrLf & _
              "Path: " & strOutputPath & vbCrLf & vbCrLf & _
              "Open the file now?", vbQuestion + vbYesNo, "Export Complete") = vbYes Then
        Application.FollowHyperlink strOutputPath
    End If
    
    ExportSuppliersToExcel = True
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    db.QueryDefs.Delete strTempQuery
    On Error GoTo 0
    
    modUtilities.LogError MODULE_NAME & ".ExportSuppliersToExcel", Err.number, Err.Description
    MsgBox "Error exporting suppliers:" & vbCrLf & vbCrLf & Err.Description, vbCritical, APP_NAME
    ExportSuppliersToExcel = False
End Function



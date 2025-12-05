
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form : frmSuppliers
' Purpose : Supplier management � mirror of frmClients with supplier-specific logic
' Author    : Tifamovs Saeol <tifamovs@gmail.com> [Benson Kimathi]
' Date : 20-Nov-2025
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Compare Database
Option Explicit

Private Const FORM_MODE_NEW As String = "NEW"
Private Const FORM_MODE_EDIT As String = "EDIT"
Private Const ENTITY_TYPE As String = "Supplier"
Private Const TABLE_NAME As String = "Suppliers"

Private m_strFormMode As String
Private m_blnDirty As Boolean


Private Sub Form_Open(Cancel As Integer)
    If Not modPermissions.HasPermission(modPermissions.PERM_VIEW_SUPPLIERS) Then
        MsgBox "You do not have permission to access Suppliers.", vbCritical, modGlobals.APP_NAME & " - Access Denied"
        Cancel = True
        Exit Sub
    End If
   
    Me.Filter = "IsDeleted = False"
    Me.FilterOn = True
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    UpdateHeaderStatus "Ready"
    modUtilities.SetButtonStates Me, PERM_EDIT_SUPPLIERS, PERM_DELETE_SUPPLIERS
     
    LoadQuickStatsSidebar
     
    m_strFormMode = FORM_MODE_EDIT
    m_blnDirty = False
   
    Me.TimerInterval = 0
   
    If Me.NewRecord Then
        UpdateHeaderStatus "NEW Supplier � Enter details below"
    End If
   
ExitHandler:
    Exit Sub
ErrorHandler:
    Call modUtilities.LogError("frmSuppliers.Form_Load", Err.number, Err.Description)
    Resume ExitHandler
End Sub

Private Sub Form_Current()
    DoEvents
    
    LoadQuickStatsSidebar
    
    If Me.dirty Then Exit Sub
   
    If Me.NewRecord Then
        UpdateHeaderStatus "NEW Supplier � Enter details below"
    Else
        UpdateHeaderStatus "Supplier: " & Nz(Me.SupplierName, "")
        On Error Resume Next
        If Not Me.subList_Suppliers.Form Is Nothing Then
            Me.subList_Suppliers.Form.SyncFromParent Me!SupplierID
        End If
        On Error GoTo 0

    End If
    
    modUtilities.SetButtonStates Me, modPermissions.PERM_EDIT_SUPPLIERS, modPermissions.PERM_DELETE_SUPPLIERS
   
End Sub

Private Sub Form_Dirty(Cancel As Integer)
    m_blnDirty = True
    
    modUtilities.SetButtonStates Me, modPermissions.PERM_EDIT_SUPPLIERS, modPermissions.PERM_DELETE_SUPPLIERS, m_blnDirty
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    On Error GoTo ErrorHandler
   
    Dim valResult As modValidation.ValidationResult
    
    valResult = modClientSupplierForms.ValidateEntityData( _
        ENTITY_TYPE, _
        Nz(Me.SupplierID, 0), _
        Trim(Nz(Me.SupplierName, "")), _
        Trim(Nz(Me.VATNumber, "")), _
        Trim(Nz(Me.Email, "")), _
        Trim(Nz(Me.Telephone, "")), _
        Trim(Nz(Me.AddressLine, "")), _
        Trim(Nz(Me.Country, "")), _
        Nz(Me.IRPFPercentage, 0), _
        Trim(Nz(Me.ContactName, "")), _
        Trim(Nz(Me.City, "")), _
        Trim(Nz(Me.ZIPCode, "")) _
    )
   
    If Not valResult.isValid Then
        Cancel = True

        MsgBox modUtilities.FormatValidationErrors(valResult), vbExclamation, modGlobals.APP_NAME & " - Validation Errors"
        modUtilities.HighlightValidationErrors Me, valResult
        Exit Sub
    End If

    modAudit.AuditFormChanges Me, "Suppliers", Me.SupplierID
    
    modUtilities.ClearValidationHighlights Me
       
    m_blnDirty = False
   
ExitHandler:
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.number & ": " & Err.Description, vbCritical, , modGlobals.APP_NAME & " - Validation Error"
    Cancel = True
    Resume ExitHandler
End Sub

Private Sub Form_AfterUpdate()
    On Error Resume Next
    modUtilities.SetButtonStates Me, modPermissions.PERM_EDIT_SUPPLIERS, modPermissions.PERM_DELETE_SUPPLIERS
    Me.subList_Suppliers.Requery
End Sub

Private Sub cmdNew_Click()
    DoCmd.GoToRecord , , acNewRec
    
    Me.VATApplied = True
    Me.VATNumber = "ES"
    Me.IRPFPercentage = 0.15
    
    m_strFormMode = FORM_MODE_NEW
    m_blnDirty = False
    
    modUtilities.SetButtonStates Me, modPermissions.PERM_EDIT_SUPPLIERS, modPermissions.PERM_DELETE_SUPPLIERS
    
    If Me.SupplierName.Visible And Me.SupplierName.Enabled Then Me.SupplierName.SetFocus
    
End Sub

Private Sub cmdSave_Click()
    If Not IsNull(Me.VATNumber) Then
        Me.VATNumber = modDatabase.FormatVATNumber(Trim(Me.VATNumber))
    End If

    If Not IsNull(Me.BankAccount) Then
        Me.BankAccount = modValidation.FormatIBAN(Me.BankAccount)
    End If
    
    If Me.dirty Then
        On Error Resume Next
        Me.dirty = False
        On Error GoTo 0
    End If

End Sub

Private Sub cmdDuplicate_Click()
    On Error GoTo ErrorHandler
    
    If Me.NewRecord Or Nz(Me.SupplierID, 0) = 0 Then
        MsgBox "Please select a valid supplier to duplicate.", vbExclamation, modGlobals.APP_NAME
        Exit Sub
    End If
   
    Dim lngOriginalID As Long
    Dim lngNewID As Long
    
    lngOriginalID = Me.SupplierID
    
    If Not modClientSupplierForms.DuplicateEntity(ENTITY_TYPE, lngOriginalID, lngNewID) Then
        MsgBox "Duplication failed. Check error log for details.", vbExclamation, modGlobals.APP_NAME
        Exit Sub
    End If
    
    If lngNewID = 0 Or IsNull(lngNewID) Then
        MsgBox "Duplication failed - no new ID returned.", vbExclamation, modGlobals.APP_NAME
        Exit Sub
    End If

    Me.Requery
    Me.subList_Suppliers.Form.Requery
    
    On Error Resume Next
    Me.Recordset.FindFirst "SupplierID = " & lngNewID

    If Me.Recordset.NoMatch Then
        MsgBox "Duplicate created but could not locate the new record.", vbExclamation, modGlobals.APP_NAME
    Else
        Me.Bookmark = Me.Recordset.Bookmark

        If Me.VATNumber.Enabled And Me.VATNumber.Visible Then
            Me.VATNumber.SetFocus
        End If

        MsgBox "Supplier duplicated successfully!" & vbCrLf & vbCrLf & _
               "Please enter a unique VAT Number and review all required fields.", _
               vbInformation, modGlobals.APP_NAME & " Duplicate Complete"
    End If
    
    If Me.SupplierID = lngNewID Then
    
        m_strFormMode = FORM_MODE_EDIT
        m_blnDirty = False
        
        UpdateHeaderStatus "DUPLICATE Supplier � Modify as needed"
        LoadQuickStatsSidebar
        
        modUtilities.SetButtonStates Me, modPermissions.PERM_EDIT_SUPPLIERS, modPermissions.PERM_DELETE_SUPPLIERS
        
        DoEvents
        If Me.VATNumber.Enabled And Me.VATNumber.Visible Then
            Me.VATNumber.SetFocus
        End If
        
    Else
        MsgBox "Record created (ID: " & lngNewID & ") but navigation failed. " & _
               "Current record: " & Nz(Me.SupplierID, 0) & vbCrLf & vbCrLf & _
               "Please search for the new record manually.", _
               vbExclamation, modGlobals.APP_NAME & " Navigation Issue"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error duplicating supplier: " & Err.number & " - " & Err.Description, _
           vbCritical, modGlobals.APP_NAME & " Duplicate Failed"
End Sub

Private Sub cmdCancel_Click()
    If Me.dirty Then Me.Undo
    m_blnDirty = False
    
    m_strFormMode = FORM_MODE_EDIT
    modUtilities.SetButtonStates Me, modPermissions.PERM_EDIT_SUPPLIERS, modPermissions.PERM_DELETE_SUPPLIERS
End Sub

Private Sub cmdDelete_Click()
    If Me.NewRecord Then Exit Sub
   
    If modClientSupplierForms.SafeDeleteEntity(ENTITY_TYPE, Me.SupplierID, Nz(Me.SupplierName, "")) Then
        Me.Requery
        Me.subList_Suppliers.Form.Requery
    End If
End Sub

Private Sub cmdAdvancedSearch_Click()
    DoCmd.OpenForm "frmAdvancedSearch", , , , , acDialog, "Supplier"
    
    If IsLoaded("frmAdvancedSearch") Then
        Dim selectedID As Variant
        selectedID = Forms!frmAdvancedSearch.selectedID
        
        DoCmd.Close acForm, "frmAdvancedSearch"
        
        If Not IsNull(selectedID) Then
            Me.Recordset.FindFirst "SupplierID = " & selectedID
            
            If Not Me.Recordset.NoMatch Then
                Me.Bookmark = Me.Recordset.Bookmark
            End If
        End If
    End If
End Sub

Private Sub cmdExportExcel_Click()
    On Error GoTo ErrorHandler
    
    modPermissions.RequirePermission modPermissions.PERM_VIEW_SUPPLIERS
    
    ' Get current filter (if any)
    Dim strFilter As String
    If Me.FilterOn Then
        strFilter = Me.Filter
    End If
    
    ' Export with filter
    If modClientSupplierForms.ExportSuppliersToExcel(strFilter) Then
        MsgBox "Export completed!", vbInformation, modGlobals.APP_NAME
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Export error: " & Err.Description, vbCritical
End Sub

' Sidebar Quick Stats � refreshed on Current & AfterUpdate
Private Sub LoadQuickStatsSidebar()
    Dim lngID As Long
    Dim strCriteria As String
    
    lngID = Nz(Me.SupplierID, 0)
    If lngID = 0 Then Exit Sub
    
    strCriteria = "SupplierID=" & lngID
    
    With Me
        .txtTotalServices.value = CachedDCount("*", "Services", strCriteria)
        .txtActiveServices.value = CachedDCount("*", "Services", strCriteria & " AND ServiceStatus='Active'")
        .txtPendingPayment.value = Format(DSum("TotalCost", "Services", strCriteria & " AND ServiceStatus IN ('Active','Completed')"), "Currency")
        .txtLastService.value = Nz(DMax("ServiceDate", "Services", strCriteria), "-")
    End With
End Sub


Private Sub txtSearch_Change()
    RestartSearchTimer Me
End Sub

Private Sub cboCountry_AfterUpdate()
    RestartSearchTimer Me
End Sub

Private Sub cboServiceType_AfterUpdate()
    RestartSearchTimer Me
End Sub

Private Sub Form_Timer()
    Me.TimerInterval = 0
    ApplyFiltersAndSearch
End Sub

Private Sub ApplyFiltersAndSearch()
    On Error GoTo ExitHandler

    DoCmd.Echo False
    Me.Painting = False
    
    Dim strFilter As String
   
    If Len(Nz(Me.txtSearch, "")) > 0 Then
        strFilter = "(SupplierName LIKE '*" & Me.txtSearch & "*' " & _
                    "OR VATNumber LIKE '*" & Me.txtSearch & "*' " & _
                    "OR Email LIKE '*" & Me.txtSearch & "*' " & _
                    "OR Telephone LIKE '*" & Me.txtSearch & "*')"
    End If
   
    If Me.cboCountry & "" <> "" Then modUtilities.AddToFilter strFilter, "Country = '" & Me.cboCountry & "'"
    If Me.cboServiceType & "" <> "" Then modUtilities.AddToFilter strFilter, "TypeOfServices = '" & Me.cboServiceType & "'"
   
    If Me.Filter <> strFilter Then
        Me.Filter = strFilter
        Me.FilterOn = (strFilter <> "")
    End If

    If Not Me.subList_Suppliers.Form Is Nothing Then
        With Me.subList_Suppliers.Form
            If .Filter <> strFilter Then
                .Filter = strFilter
                .FilterOn = (strFilter <> "")
                .Requery
            End If
        End With
    End If
   
    ' Get count
    Dim strCountFilter As String
    If Len(Nz(strFilter, "")) > 0 Then strCountFilter = strFilter & " AND IsDeleted = False"
    Dim lngCount As Long: lngCount = CachedDCount("*", TABLE_NAME, Nz(strCountFilter, strFilter))
    
    ' Update UI
    If lngCount = 0 Then
        Me.lblResultCount.Caption = "No results found"
        Me.lblResultCount.ForeColor = vbRed
    ElseIf lngCount = 1 Then
        Me.lblResultCount.Caption = "Found 1 matching " & LCase(strFilter)
        Me.lblResultCount.ForeColor = RGB(0, 100, 0)
    Else
        Me.lblResultCount.Caption = "Found " & lngCount & " matching " & LCase(strFilter) & "s"
        Me.lblResultCount.ForeColor = RGB(0, 100, 0)
    End If
    
ExitHandler:
    Me.Painting = True
    DoCmd.Echo True

End Sub

Private Sub cmdClearFilters_Click()
    Me.txtSearch = Null
    Me.cboCountry = Null
    Me.cboServiceType = Null
    Me.lblResultCount.Caption = ""
    Me.FilterOn = False
    
End Sub

Private Sub UpdateHeaderStatus(strText As String)
    Me.lblStatus.Caption = strText
    Me.lblStatus.ForeColor = IIf(Left(strText, 3) = "NEW", vbBlue, vbBlack)
End Sub

Private Sub cmdCopyEmail_Click()
    If Not IsNull(Me.Email) And Len(Trim(Me.Email)) > 0 Then
        modUtilities.CopyToClipboard (Me.Email)
        modUtilities.ShowCopiedFeedback (Me.lblCopyEmail)
    Else
        MsgBox "Email is empty.", vbInformation, modGlobals.APP_NAME
    End If
End Sub

Private Sub cmdCopyPhone_Click()
    If Not IsNull(Me.Telephone) And Len(Trim(Me.Telephone)) > 0 Then
        modUtilities.CopyToClipboard (Me.Telephone)
        modUtilities.ShowCopiedFeedback (Me.lblCopyPhone)
    Else
        MsgBox "Telephone is empty.", vbInformation, modGlobals.APP_NAME
    End If
End Sub

Public Sub SyncSupplier(lngSupplierID As Long)
    If Me.Recordset.RecordCount = 0 Then Exit Sub
    
    If Me.SupplierID <> lngSupplierID Then
        Me.Recordset.FindFirst "SupplierID = " & CLng(lngSupplierID)
    End If
End Sub

Public Property Get IsDirty() As Boolean
    IsDirty = m_blnDirty
End Property

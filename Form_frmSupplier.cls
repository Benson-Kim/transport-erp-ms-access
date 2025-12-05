VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form      : frmSuppliers
' Purpose   : Supplier management – mirror of frmClients with supplier-specific logic
' Author    : Expert Access VBA Developer
' Date      : 20-Nov-2025
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Database
Option Explicit

Private Const FORM_MODE_NEW As String = "NEW"
Private Const FORM_MODE_EDIT As String = "EDIT"
Private m_strFormMode As String
Private m_blnDirty As Boolean

Private Sub cmdSave_Click()

End Sub

Private Sub Form_Open(Cancel As Integer)
    If Not modPermissions.HasPermission("PERM_VIEW_SUPPLIERS") Then
        MsgBox "You do not have permission to access Suppliers.", vbCritical, "Access Denied"
        Cancel = True
        Exit Sub
    End If
    
    Me.Filter = "IsDeleted = False"
    Me.FilterOn = True
    Me.TimerInterval = 400
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    Me.AutoHeader.Caption = "Supplier Management – " & modGlobals.APP_VERSION
    
    UpdateHeaderStatus "Ready"
    
    Call RefreshFilterDisplay
    modUtilities.SetButtonStates Me, PERM_EDIT_SUPPLIERS, PERM_DELETE_SUPPLIERS
    Call LoadQuickStatsSidebar
    m_strFormMode = FORM_MODE_EDIT
    m_blnDirty = False
    
    ' Set up search timer
    Me.TimerInterval = 300  ' 300ms delay for live search
    
    ' Header info
    
    
    ' Initial focus
    If Me.NewRecord Then UpdateHeaderStatus "NEW Supplier – Enter details below"
    
ExitHandler:
    Exit Sub
ErrorHandler:
    Call modUtilities.LogError("frmSuppliers.Form_Load", Err.Number, Err.Description)
    Resume ExitHandler
End Sub


Private Sub Form_Current()
    Call LoadQuickStatsSidebar
    modUtilities.SetButtonStates Me, PERM_EDIT_SUPPLIERS, PERM_DELETE_SUPPLIERS
    Call RefreshFilterDisplay
    
    If Me.NewRecord Then
        UpdateHeaderStatus "NEW Supplier – Enter details below"
    Else
        UpdateHeaderStatus "Supplier: " & Nz(Me.SupplierName, "")
        Me.subList_Suppliers.Form.SyncFromParent Me!SupplierID
    End If
    
End Sub


Private Sub Form_BeforeUpdate(Cancel As Integer)
'    On Error GoTo ErrorHandler
    
    If ValidateSupplierData = False Then
        Cancel = True
        Exit Sub
    End If
    
    ' Auto-format fields
    Me.VATNumber = modDatabase.FormatVATNumber(Trim(Me.VATNumber))
    If Not IsNull(Me.BankAccount) Then
        Me.BankAccount = modValidation.FormatIBAN(Me.BankAccount)
    End If
    
    Call AuditFormChanges(Me, "Suppliers", Me.SupplierID)
    
    m_blnDirty = False
    
ExitHandler:
    Exit Sub
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Validation Error"
    Cancel = True
    Resume ExitHandler
End Sub


Private Sub Form_Dirty(Cancel As Integer)
    m_blnDirty = True
End Sub


Private Function ValidateSupplierData() As Boolean
    On Error GoTo ErrorHandler
    ValidateSupplierData = False
    
    With Me
        If Len(Trim(.SupplierName & "")) < 2 Then
            MsgBox "Supplier name is required and must be at least 2 characters.", vbExclamation
            .SupplierName.SetFocus
            Exit Function
        End If
        
        If Not IsNull(DLookup("SupplierID", "Suppliers", "VATNumber = '" & Replace(.VATNumber, "'", "''") & _
                              "' AND SupplierID <> " & Nz(.SupplierID, 0))) Then
            MsgBox "This VAT number is already in use by another supplier.", vbExclamation
            .VATNumber.SetFocus
            Exit Function
        End If
        
        If Not modValidation.IsValidEmail(.Email) Then
            MsgBox "Please enter a valid e-mail address.", vbExclamation
            .Email.SetFocus
            Exit Function
        End If
        
        If Nz(.IRPFPercentage, 0) < 0 Or Nz(.IRPFPercentage, 0) > 100 Then
            MsgBox "IRPF percentage must be between 0 and 100.", vbExclamation
            .IRPFPercentage.SetFocus
            Exit Function
        End If
        
        ' Spanish VAT quick check (ES + 8 digits + optional letter)
        If Left(UCase(.VATNumber), 2) = "ES" And Not .VATNumber Like "ES########?" Then
            MsgBox "Spanish VAT number format is ES followed by 8 digits and optional letter.", vbExclamation
            .VATNumber.SetFocus
            Exit Function
        End If
    End With
    
    ValidateSupplierData = True
ExitHandler:
    Exit Function
ErrorHandler:
    Call LogError(Err.Number, Err.Description, "frmSuppliers.ValidateSupplierData")
    Resume ExitHandler
End Function


Private Sub cmdNew_Click()
    DoCmd.GoToRecord , , acNewRec
    Me.VATApplied = True
    Me.VATNumber = "ES"
    Me.IRPFPercentage = 0.15   ' most common Spanish rate
    m_strFormMode = FORM_MODE_NEW
    m_blnDirty = False
    modUtilities.SetButtonStates Me, PERM_EDIT_SUPPLIERS, PERM_DELETE_SUPPLIERS
End Sub


Private Sub cmdDuplicate_Click()
    On Error GoTo ErrorHandler

    If Me.NewRecord Or Nz(Me.SupplierID, 0) = 0 Then
        MsgBox "Please select a valid supplier to duplicate.", vbExclamation
        Exit Sub
    End If

    Dim lngOldID As Long:   lngOldID = Me.SupplierID
    Dim lngMaxID As Long:   lngMaxID = Nz(DMax("SupplierID", "Suppliers"), 0)
    Dim strDummyVAT As String
    
    strDummyVAT = "TEMP_DUP_" & lngOldID & "_" & Format(Now(), "yyyymmddhhnnss") & Int(Rnd * 10000)

    Dim strSQL As String

    strSQL = "INSERT INTO Suppliers (" & _
             "SupplierName, ContactName, AddressLine, ZipCode, City, Country, " & _
             "Telephone, Email, VATNumber, BankAccount, BankDetails, PaymentTerms, " & _
             "IRPFPercentage, TypeOfServices, VATApplied, IsDeleted, " & _
             "CreatedBy, CreatedDate" & _
             ") " & _
             "SELECT " & _
             "SupplierName & ' - Copy' AS Supplier, " & _
             "ContactName, AddressLine, ZipCode, City, Country, " & _
             "Telephone, Email, " & _
             "'" & strDummyVAT & "' AS VATNumber, " & _
             "BankAccount, BankDetails, PaymentTerms, " & _
             "IRPFPercentage, TypeOfServices, VATApplied, " & _
             "False AS IsDeleted, " & _
             g_lngUserID & " AS CreatedBy, " & _
             "Now() AS CreatedDate " & _
             "FROM Suppliers WHERE SupplierID = " & lngOldID & ";"

    modDatabase.ExecuteQuery strSQL
    
    Dim lngNewID As Long: lngNewID = Nz(DMax("SupplierID", "Suppliers"), 0)

    If lngNewID <= lngMaxID Then
        MsgBox "Duplicate failed – could not determine new record ID.", vbCritical
        Exit Sub
    End If

    Me.Requery

    Me.Recordset.FindFirst "SupplierID = " & lngNewID

    If Me.Recordset.NoMatch Then
        MsgBox "Duplicate created but could not locate the new record.", vbExclamation
    Else
        Me.Bookmark = Me.Recordset.Bookmark

        If Me.VATNumber.Enabled And Me.VATNumber.Visible Then
            Me.VATNumber.SetFocus
        End If

        MsgBox "Supplier duplicated successfully!" & vbCrLf & vbCrLf & _
               "Please enter a unique VAT Number and review all required fields.", _
               vbInformation, "Duplicate Complete"
    End If

    ' Update UI
    m_strFormMode = FORM_MODE_NEW
    m_blnDirty = False
    Call LoadQuickStatsSidebar
    UpdateHeaderStatus "DUPLICATE Supplier – Modify as needed"
    modUtilities.SetButtonStates Me, PERM_EDIT_SUPPLIERS, PERM_DELETE_SUPPLIERS

    Exit Sub

ErrorHandler:
    DoCmd.SetWarnings True
    MsgBox "Error duplicating supplier: " & Err.Number & " - " & Err.Description, vbCritical, "Duplicate Failed"
End Sub



Private Sub cmdDelete_Click()
    If Me.NewRecord Then Exit Sub
    
    Dim lngCount As Long
    lngCount = DCount("*", "Services", "SupplierID = " & Me.SupplierID & " AND IsDeleted = False")
    
    If lngCount > 0 Then
        If MsgBox("This supplier has " & lngCount & " active service(s)." & vbCrLf & _
                  "Soft-delete anyway?", vbQuestion + vbYesNo, "Confirm Delete") = vbNo Then Exit Sub
    End If
    
    Dim strConfirm As String
    strConfirm = InputBox("Type DELETE in capital letters to confirm permanent soft-deletion:", "Confirm")
    If StrComp(strConfirm, "DELETE", vbTextCompare) <> 0 Then
        MsgBox "Deletion cancelled.", vbInformation
        Exit Sub
    End If
    
    CurrentDb.Execute "UPDATE Suppliers SET IsDeleted = True, ModifiedDate = Now(), ModifiedBy = " & g_lngUserID & _
                      " WHERE SupplierID = " & Me.SupplierID, dbFailOnError
    
    Call modAudit.LogAudit("Suppliers", Me.SupplierID, IsDeleted, "False", "True", "Delete", g_lngUserID, True)

    Me.Requery
    MsgBox "Supplier has been archived.", vbInformation
End Sub


Private Sub cmdIRPF15_Click(): Me.IRPFPercentage = 15: End Sub
Private Sub cmdIRPF1_Click():  Me.IRPFPercentage = 1:  End Sub
Private Sub cmdIRPF0_Click():  Me.IRPFPercentage = 0:  End Sub


' Sidebar Quick Stats – refreshed on Current & AfterUpdate
Private Sub LoadQuickStatsSidebar()
    Dim lngID As Long: lngID = Nz(Me.SupplierID, 0)
    If lngID = 0 Then Exit Sub
    
    With Me
        .txtTotalServices.Value = DCount("*", "Services", "SupplierID = " & lngID)
        .txtActiveServices.Value = DCount("*", "Services", "SupplierID = " & lngID & " AND ServiceStatus = 'Active'")
        .txtPendingPayment.Value = Format(DSum("TotalCost", "Services", "SupplierID = " & lngID & " AND ServiceStatus IN ('Active','Completed')"), "Currency")
        .txtLastService.Value = Nz(DMax("ServiceDate", "Services", "SupplierID = " & lngID), "-")
    End With
End Sub



' Real-time search + advanced filters (Country, VAT Applied, Service Type, IRPF range)
Private Sub txtSearch_Change()
    Me.TimerInterval = 600  ' Restart timer
End Sub

Private Sub Form_Timer()
    Me.TimerInterval = 0
    ApplyFiltersAndSearch
End Sub

Private Sub cboCountry_AfterUpdate(): Call ApplyFiltersAndSearch: End Sub
Private Sub cboServiceType_AfterUpdate(): Call ApplyFiltersAndSearch: End Sub

Private Sub ApplyFiltersAndSearch()
    Dim strFilter As String
    
    If Len(Me.txtSearch & "") > 0 Then
        strFilter = "(SupplierName LIKE '*" & Me.txtSearch & "*' OR VATNumber LIKE '*" & Me.txtSearch & "*' OR Telephone LIKE '*" & Me.txtSearch & "*')"
    End If
    
    If Me.cboCountry & "" <> "" Then AddToFilter strFilter, "Country = '" & Me.cboCountry & "'"
    If Me.cboServiceType & "" <> "" Then AddToFilter strFilter, "TypeOfServices = '" & Me.cboServiceType & "'"
    
    Me.Filter = strFilter
    Me.FilterOn = (strFilter <> "")
    Call RefreshFilterDisplay
End Sub

Private Sub AddToFilter(ByRef strF As String, strNew As String)
    If strF = "" Then strF = strNew Else strF = strF & " AND " & strNew
End Sub

Private Sub cmdClearFilters_Click()
    Me.txtSearch = Null
    Me.cboCountry = Null
    Me.cboServiceType = Null
    Me.FilterOn = False
    Call RefreshFilterDisplay
End Sub

Private Sub RefreshFilterDisplay()
    DoEvents
    modPermissions.RequirePermission "PERM_VIEW_SUPPLIERS"
    Me.txtUser.Value = IIf(Len(g_strFullName) > 0, g_strFullName, "____________")
    Me.txtRole.Value = IIf(Len(g_strUserRole) > 0, g_strUserRole, "____________")
End Sub

Private Sub UpdateHeaderStatus(strText As String)
    Me.lblStatus.Caption = strText
    Me.lblStatus.ForeColor = IIf(Left(strText, 3) = "NEW", vbBlue, vbBlack)
End Sub

Private Sub cmdCopyEmail_Click()
    If Not IsNull(Me.Email) And Len(Trim(Me.Email)) > 0 Then
        Call modUtilities.CopyToClipboard(Me.Email)
        Call modUtilities.ShowCopiedFeedback(Me.lblCopyEmail)
    Else
        MsgBox "Email is empty.", vbInformation
    End If
End Sub

Private Sub cmdCopyPhone_Click()
    If Not IsNull(Me.Telephone) And Len(Trim(Me.Telephone)) > 0 Then
        Call modUtilities.CopyToClipboard(Me.Telephone)
        Call modUtilities.ShowCopiedFeedback(Me.lblCopyPhone)
    Else
        MsgBox "Telephone is empty.", vbInformation
    End If
End Sub

Public Sub SyncSupplier(lngSupplierID As Long)
    If Me.SupplierID <> lngSupplierID Then
        Me.Recordset.FindFirst "SupplierID = " & lngSupplierID
    End If
End Sub

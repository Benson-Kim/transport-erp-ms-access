VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSuppliers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form : frmSuppliers
' Purpose : Supplier management – mirror of frmClients with supplier-specific logic
' Author : Expert Access VBA Developer
' Date : 20-Nov-2025
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Compare Database
Option Explicit

Private Const FORM_MODE_NEW As String = "NEW"
Private Const FORM_MODE_EDIT As String = "EDIT"
Private Const ENTITY_TYPE As String = "Supplier"  ' Used for module calls

Private m_strFormMode As String
Private m_blnDirty As Boolean

Private Sub cmdSave_Click()
    If m_blnDirty Then
        If Me.dirty Then
            DoCmd.RunCommand acCmdSaveRecord
        End If
    End If
End Sub

Private Sub Form_Open(Cancel As Integer)
    If Not modPermissions.HasPermission(modPermissions.PERM_VIEW_SUPPLIERS) Then
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
   
    ' Me.AutoHeader.Caption = "Supplier Management – " & modGlobals.APP_VERSION
   
    UpdateHeaderStatus "Ready"
   
    Call RefreshFilterDisplay
    modUtilities.SetButtonStates Me, PERM_EDIT_SUPPLIERS, PERM_DELETE_SUPPLIERS
    Call LoadQuickStatsSidebar
    m_strFormMode = FORM_MODE_EDIT
    m_blnDirty = False
   
    ' Set up search timer
    Me.TimerInterval = 300 ' 300ms delay for live search
   
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
    Call modUtilities.SetButtonStates(Me, PERM_EDIT_SUPPLIERS, PERM_DELETE_SUPPLIERS)
    Call RefreshFilterDisplay
   
    If Me.NewRecord Then
        UpdateHeaderStatus "NEW Supplier – Enter details below"
    Else
        UpdateHeaderStatus "Supplier: " & Nz(Me.SupplierName, "")
        Me.subList_Suppliers.Form.SyncFromParent Me!SupplierID
    End If
   
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    On Error GoTo ErrorHandler
   
    Dim valResult As ValidationResult
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
   
    If Not valResult.IsValid Then
        MsgBox modClientSupplierForms.FormatValidationErrors(valResult), vbExclamation, "Validation Errors"
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

Private Sub cmdNew_Click()
    DoCmd.GoToRecord , , acNewRec
    Me.VATApplied = True
    Me.VATNumber = "ES"
    Me.IRPFPercentage = 0.15 ' most common Spanish rate
    m_strFormMode = FORM_MODE_NEW
    m_blnDirty = False
    modUtilities.SetButtonStates Me, PERM_EDIT_SUPPLIERS, PERM_DELETE_SUPPLIERS
    If Me.SupplierName.Visible And Me.SupplierName.Enabled Then Me.SupplierName.SetFocus
End Sub

Private Sub cmdDuplicate_Click()
    On Error GoTo ErrorHandler
    If Me.NewRecord Or Nz(Me.SupplierID, 0) = 0 Then
        MsgBox "Please select a valid supplier to duplicate.", vbExclamation
        Exit Sub
    End If
   
    Dim lngNewID As Long
    If modClientSupplierForms.DuplicateEntity(ENTITY_TYPE, Me.SupplierID, lngNewID) Then
        Me.Requery
        Me.Recordset.FindFirst "SupplierID = " & lngNewID
        If Not Me.Recordset.NoMatch Then
            Me.Bookmark = Me.Recordset.Bookmark
            If Me.VATNumber.Enabled And Me.VATNumber.Visible Then Me.VATNumber.SetFocus
            MsgBox "Supplier duplicated successfully!" & vbCrLf & vbCrLf & _
                   "Please enter a unique VAT Number and review all required fields.", _
                   vbInformation, "Duplicate Complete"
        Else
            MsgBox "Duplicate created but could not locate the new record.", vbExclamation
        End If
        m_strFormMode = FORM_MODE_NEW
        m_blnDirty = False
        Call LoadQuickStatsSidebar
        UpdateHeaderStatus "DUPLICATE Supplier – Modify as needed"
        modUtilities.SetButtonStates Me, PERM_EDIT_SUPPLIERS, PERM_DELETE_SUPPLIERS
        Me.subList_Suppliers.Form.Requery  ' Refresh list
    End If
    Exit Sub
ErrorHandler:
    MsgBox "Error duplicating supplier: " & Err.Number & " - " & Err.Description, vbCritical, "Duplicate Failed"
End Sub

Private Sub cmdDelete_Click()
    If Me.NewRecord Then Exit Sub
   
    If modClientSupplierForms.SafeDeleteEntity(ENTITY_TYPE, Me.SupplierID, Nz(Me.SupplierName, "")) Then
        Me.Requery
        Me.subList_Suppliers.Form.Requery  ' Refresh list
    End If
End Sub

Private Sub cmdIRPF15_Click(): Me.IRPFPercentage = 15: End Sub
Private Sub cmdIRPF1_Click(): Me.IRPFPercentage = 1: End Sub
Private Sub cmdIRPF0_Click(): Me.IRPFPercentage = 0: End Sub

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
    Me.TimerInterval = 600 ' Restart timer
End Sub

Private Sub Form_Timer()
    Me.TimerInterval = 0
    ApplyFiltersAndSearch
End Sub

Private Sub cboCountry_AfterUpdate(): Call ApplyFiltersAndSearch: End Sub
Private Sub cboServiceType_AfterUpdate(): Call ApplyFiltersAndSearch: End Sub

Private Sub ApplyFiltersAndSearch()
    On Error Resume Next
    Dim strFilter As String
   
    If Len(Me.txtSearch & "") > 0 Then
        strFilter = "(SupplierName LIKE '*" & Me.txtSearch & "*' OR VATNumber LIKE '*" & Me.txtSearch & "*' OR Telephone LIKE '*" & Me.txtSearch & "*')"
    End If
   
    If Me.cboCountry & "" <> "" Then modUtilities.AddToFilter strFilter, "Country = '" & Me.cboCountry & "'"
    If Me.cboServiceType & "" <> "" Then modUtilities.AddToFilter strFilter, "TypeOfServices = '" & Me.cboServiceType & "'"
   
    Me.Filter = strFilter
    Me.FilterOn = (strFilter <> "")
   
    With Me.subList_Suppliers.Form
        .Filter = strFilter
        .FilterOn = (strFilter <> "")
        .Requery
    End With
   
    Call RefreshFilterDisplay
End Sub


Private Sub cmdClearFilters_Click()
    Me.txtSearch = Null
    Me.cboCountry = Null
    Me.cboServiceType = Null
    Me.FilterOn = False
    Call RefreshFilterDisplay
End Sub

Private Sub RefreshFilterDisplay()
    Call modPermissions.RequirePermission(PERM_VIEW_SUPPLIERS)
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
        Me.Recordset.FindFirst "SupplierID = " & CLng(lngSupplierID)
    End If
End Sub


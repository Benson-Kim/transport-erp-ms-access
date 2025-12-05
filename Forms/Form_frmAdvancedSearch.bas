'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' FORM      : frmAdvancedSearch
' PURPOSE   : Universal advanced search for Clients and Suppliers with saved searches
' AUTHOR    : Tifamovs Saeol <tifamovs@gmail.com> [Benson Kimathi]
' CREATED      : 2025-12-04
' UPDATED      : 2025-12-05 - Improved error handling and logging
' USAGE     : DoCmd.OpenForm "frmAdvancedSearch", , , , , acDialog, "Client" or "Supplier"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Database
Option Explicit

' MODULE-LEVEL VARIABLES
Private m_strSearchType As String           ' "Client" or "Supplier"
Private m_varSelectedID As Variant          ' Return value
Private m_blnLoading As Boolean
Private m_strCurrentSavedSearch As String

' PUBLIC PROPERTIES - RETURN VALUE
Public Property Get selectedID() As Variant
    selectedID = m_varSelectedID
End Property

Public Property Get SearchType() As String
    SearchType = m_strSearchType
End Property

Private Sub cmdSave_Click()

End Sub

' FORM EVENTS - INITIALIZATION
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo ErrorHandler
    
    m_blnLoading = True
    m_varSelectedID = Null
    
    ' Get search type from OpenArgs
    If Not IsNull(Me.OpenArgs) Then
        m_strSearchType = CStr(Me.OpenArgs)
    Else
        m_strSearchType = "Client"  ' Default
    End If
    
    ' Validate search type
    If m_strSearchType <> "Client" And m_strSearchType <> "Supplier" Then
        MsgBox "Invalid search type. Must be 'Client' or 'Supplier'.", vbCritical
        Cancel = True
        Exit Sub
    End If
    
    ' Configure form appearance
    Me.Caption = "Advanced " & m_strSearchType & " Search"
    Me.lblTitle.Caption = "Search " & m_strSearchType & "s"
    
    ' Show/hide supplier-specific controls
    Me.txtIRPFFrom.Visible = (m_strSearchType = "Supplier")
    Me.txtIRPFTo.Visible = (m_strSearchType = "Supplier")
    Me.cboServiceType.Visible = (m_strSearchType = "Supplier")
    
    Exit Sub
    
ErrorHandler:
    modUtilities.LogError "frmAdvancedSearch.Form_Open", Err.number, Err.Description
    Cancel = True
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    ' Initialize controls
    Call ClearSearchCriteria
    Call LoadFilterControls
    Call LoadSavedSearches
    
    ' Configure subform
    Me.subList_SearchResults.SourceObject = ""  ' Will be set on search
    Me.lblResultCount.Caption = "Enter search criteria and click Search"
    
    ' Set focus
    On Error Resume Next
    Me.txtName.SetFocus
    On Error GoTo ErrorHandler
    
    m_blnLoading = False
    
    Exit Sub
    
ErrorHandler:
    modUtilities.LogError "frmAdvancedSearch.Form_Load", Err.number, Err.Description
    m_blnLoading = False
End Sub

' SEARCH EXECUTION
Private Sub cmdSearch_Click()
    On Error GoTo ErrorHandler
    
    Dim strSQL As String
    Dim strWhere As String
    Dim lngCount As Long
    
    ' Build WHERE clause
    strWhere = BuildWhereClause()
    
    If strWhere = "" Then
        MsgBox "Please enter at least one search criterion.", vbInformation, "No Criteria"
        Exit Sub
    End If
    
    ' Build complete SQL
    strSQL = BuildSearchSQL(strWhere)
    
    ' Create temporary query for subform
    Call CreateSearchResultsQuery(strSQL)
    
    ' Bind subform
    Me.subList_SearchResults.SourceObject = "Query.qryAdvancedSearchResults_Temp"
    
    ' Get count
    lngCount = DCount("*", "qryAdvancedSearchResults_Temp")
    
    ' Update UI
    If lngCount = 0 Then
        Me.lblResultCount.Caption = "No results found"
        Me.lblResultCount.ForeColor = vbRed
        Me.cmdSelect.Enabled = False
    ElseIf lngCount = 1 Then
        Me.lblResultCount.Caption = "Found 1 matching " & LCase(m_strSearchType)
        Me.lblResultCount.ForeColor = RGB(0, 100, 0)
        Me.cmdSelect.Enabled = True
    Else
        Me.lblResultCount.Caption = "Found " & lngCount & " matching " & LCase(m_strSearchType) & "s"
        Me.lblResultCount.ForeColor = RGB(0, 100, 0)
        Me.cmdSelect.Enabled = True
    End If
    
    Exit Sub
    
ErrorHandler:
    modUtilities.LogError "frmAdvancedSearch.cmdSearch_Click", Err.number, Err.Description
    MsgBox "Error executing search: " & Err.Description, vbCritical
End Sub

Private Function BuildWhereClause() As String
    On Error GoTo ErrorHandler
    
    Dim strWhere As String
    strWhere = "IsDeleted = False"  ' Base condition
    
    Dim tableName As String
    tableName = IIf(m_strSearchType = "Client", "Clients", "Suppliers")
    
    Dim nameField As String
    nameField = IIf(m_strSearchType = "Client", "ClientName", "SupplierName")
    
    Dim emailField As String
    emailField = IIf(m_strSearchType = "Client", "EmailBilling", "Email")
    
    ' Name contains
    If Not IsNull(Me.txtName) And Len(Trim(Me.txtName)) > 0 Then
        modUtilities.AddToFilter strWhere, _
            nameField & " LIKE '*" & Replace(Trim(Me.txtName), "'", "''") & "*'"
    End If
    
    ' VAT Number exact match
    If Not IsNull(Me.txtVATNumber) And Len(Trim(Me.txtVATNumber)) > 0 Then
        modUtilities.AddToFilter strWhere, _
            "VATNumber = '" & Replace(Trim(Me.txtVATNumber), "'", "''") & "'"
    End If
    
    ' Telephone contains
    If Not IsNull(Me.txtTelephone) And Len(Trim(Me.txtTelephone)) > 0 Then
        modUtilities.AddToFilter strWhere, _
            "Telephone LIKE '*" & Replace(Trim(Me.txtTelephone), "'", "''") & "*'"
    End If
    
    ' Email contains
    If Not IsNull(Me.txtEmail) And Len(Trim(Me.txtEmail)) > 0 Then
        modUtilities.AddToFilter strWhere, _
            emailField & " LIKE '*" & Replace(Trim(Me.txtEmail), "'", "''") & "*'"
    End If
    
    ' Country
    If Not IsNull(Me.cboCountry) And Len(Trim(Me.cboCountry)) > 0 Then
        modUtilities.AddToFilter strWhere, _
            "Country = '" & Replace(Me.cboCountry, "'", "''") & "'"
    End If
    
    ' Address contains (search in AddressLine, ZipCode, and City)
    If Not IsNull(Me.txtAddress) And Len(Trim(Me.txtAddress)) > 0 Then
        Dim addressField As String
        Dim searchText As String
    
        searchText = Replace(Trim(Me.txtAddress), "'", "''")
    
        addressField = "(AddressLine LIKE '*" & searchText & "*' " & _
                       "OR ZipCode LIKE '*" & searchText & "*' " & _
                       "OR City LIKE '*" & searchText & "*')"
    
        modUtilities.AddToFilter strWhere, addressField
    End If

    
    ' VAT Applied
    If Not IsNull(Me.cboVATApplied) And Me.cboVATApplied <> "" Then
        If Me.cboVATApplied = "Yes" Then
            modUtilities.AddToFilter strWhere, "VATApplied = True"
        ElseIf Me.cboVATApplied = "No" Then
            modUtilities.AddToFilter strWhere, "VATApplied = False"
        End If
    End If
    
    ' Supplier-specific: IRPF range
    If m_strSearchType = "Supplier" Then
        Dim dblIRPFFrom As Double
        Dim dblIRPFTo As Double
        
        If Not IsNull(Me.txtIRPFFrom) And IsNumeric(Me.txtIRPFFrom) Then
            dblIRPFFrom = CDbl(Me.txtIRPFFrom)
            modUtilities.AddToFilter strWhere, "IRPFPercentage >= " & dblIRPFFrom
        End If
        
        If Not IsNull(Me.txtIRPFTo) And IsNumeric(Me.txtIRPFTo) Then
            dblIRPFTo = CDbl(Me.txtIRPFTo)
            modUtilities.AddToFilter strWhere, "IRPFPercentage <= " & dblIRPFTo
        End If
        
        ' Service Type
        If Not IsNull(Me.cboServiceType) And Len(Trim(Me.cboServiceType)) > 0 Then
            modUtilities.AddToFilter strWhere, _
                "TypeOfServices = '" & Replace(Me.cboServiceType, "'", "''") & "'"
        End If
    End If
    
    BuildWhereClause = strWhere
    Exit Function
    
ErrorHandler:
    modUtilities.LogError "frmAdvancedSearch.BuildWhereClause", Err.number, Err.Description
    BuildWhereClause = "IsDeleted = False"
End Function

Private Function BuildSearchSQL(strWhere As String) As String
    On Error GoTo ErrorHandler
    
    Dim strSQL As String
    Dim tableName As String
    Dim pkField As String
    Dim nameField As String
    Dim emailField As String
    
    If m_strSearchType = "Client" Then
        tableName = "Clients"
        pkField = "ClientID"
        nameField = "ClientName"
        emailField = "EmailBilling"
        
        strSQL = "SELECT " & pkField & ", " & nameField & " AS Name, " & _
                 "VATNumber, Country, Telephone, " & emailField & " AS Email, " & _
                 "PaymentTerms, VATApplied " & _
                 "FROM " & tableName & " " & _
                 "WHERE " & strWhere & " " & _
                 "ORDER BY " & nameField
    Else
        tableName = "Suppliers"
        pkField = "SupplierID"
        nameField = "SupplierName"
        emailField = "Email"
        
        strSQL = "SELECT " & pkField & ", " & nameField & " AS Name, " & _
                 "VATNumber, Country, Telephone, " & emailField & " AS Email, " & _
                 "IRPFPercentage, TypeOfServices, PaymentTerms, VATApplied " & _
                 "FROM " & tableName & " " & _
                 "WHERE " & strWhere & " " & _
                 "ORDER BY " & nameField
    End If
    
    BuildSearchSQL = strSQL
    Exit Function
    
ErrorHandler:
    modUtilities.LogError "frmAdvancedSearch.BuildSearchSQL", Err.number, Err.Description
    BuildSearchSQL = ""
End Function

Private Sub CreateSearchResultsQuery(strSQL As String)
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    
    Set db = CurrentDb
    
    ' Delete existing temp query if exists
    On Error Resume Next
    db.QueryDefs.Delete "qryAdvancedSearchResults_Temp"
    On Error GoTo ErrorHandler
    
    ' Create new query
    Set qdf = db.CreateQueryDef("qryAdvancedSearchResults_Temp", strSQL)
    
    Exit Sub
    
ErrorHandler:
    modUtilities.LogError "frmAdvancedSearch.CreateSearchResultsQuery", Err.number, Err.Description
End Sub

' RESULT SELECTION
Private Sub cmdSelect_Click()
    On Error GoTo ErrorHandler
    
    ' Verify subform has data
    If Me.subList_SearchResults.Form.RecordsetClone.RecordCount = 0 Then
        MsgBox "No results to select.", vbInformation
        Exit Sub
    End If
    
    ' Get selected ID from subform
    Dim pkField As String
    pkField = IIf(m_strSearchType = "Client", "ClientID", "SupplierID")
    
    On Error Resume Next
    m_varSelectedID = Me.subList_SearchResults.Form.Controls(pkField).value
    On Error GoTo ErrorHandler
    
    If IsNull(m_varSelectedID) Or m_varSelectedID = 0 Then
        MsgBox "Please select a " & LCase(m_strSearchType) & " from the results.", vbInformation
        Exit Sub
    End If
    
    ' Close form (calling form will retrieve SelectedID property)
    Me.Visible = False
    
    Exit Sub
    
ErrorHandler:
    modUtilities.LogError "frmAdvancedSearch.cmdSelect_Click", Err.number, Err.Description
    MsgBox "Error selecting " & LCase(m_strSearchType) & ": " & Err.Description, vbCritical
End Sub

Private Sub cmdCancel_Click()
    m_varSelectedID = Null
    Me.Visible = False
End Sub

Private Sub subList_SearchResults_DblClick(Cancel As Integer)
    ' Double-click to select
    Call cmdSelect_Click
End Sub

' SAVED SEARCHES
Private Sub LoadSavedSearches()
    On Error GoTo ErrorHandler
    
    Dim strSQL As String
    
    ' Load user's saved searches + shared searches
    strSQL = "SELECT SearchID, SearchName " & _
             "FROM SavedSearches " & _
             "WHERE SearchType = '" & m_strSearchType & "' " & _
             "AND (UserID = " & modGlobals.UserID & " OR IsShared = True) " & _
             "ORDER BY SearchName"
    
    Me.cboSavedSearches.RowSource = strSQL
    Me.cboSavedSearches.Requery
    
    Exit Sub
    
ErrorHandler:
    modUtilities.LogError "frmAdvancedSearch.LoadSavedSearches", Err.number, Err.Description
End Sub

Private Sub cboSavedSearches_AfterUpdate()
    On Error GoTo ErrorHandler
    
    If IsNull(Me.cboSavedSearches) Then Exit Sub
    
    Call LoadSavedSearch(Me.cboSavedSearches)
    
    Exit Sub
    
ErrorHandler:
    modUtilities.LogError "frmAdvancedSearch.cboSavedSearches_AfterUpdate", Err.number, Err.Description
End Sub

Private Sub LoadSavedSearch(lngSearchID As Long)
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strCriteria As String
    Dim dictCriteria As Object
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
        "SELECT SearchCriteria, SearchName FROM SavedSearches WHERE SearchID = " & lngSearchID, _
        dbOpenSnapshot)
    
    If rs.EOF Then
        MsgBox "Saved search not found.", vbExclamation
        Exit Sub
    End If
    
    strCriteria = Nz(rs!SearchCriteria, "")
    m_strCurrentSavedSearch = Nz(rs!SearchName, "")
    rs.Close
    
    If strCriteria = "" Then Exit Sub
    
    ' Parse JSON criteria
    Set dictCriteria = ParseSearchCriteria(strCriteria)
    
    If dictCriteria Is Nothing Then
        MsgBox "Error loading saved search criteria.", vbExclamation
        Exit Sub
    End If
    
    ' Populate form controls
    Call PopulateCriteriaFromDict(dictCriteria)
    
    ' Auto-execute search
    Call cmdSearch_Click
    
    Exit Sub
    
ErrorHandler:
    modUtilities.LogError "frmAdvancedSearch.LoadSavedSearch", Err.number, Err.Description
End Sub

Private Sub cmdSaveSearch_Click()
    On Error GoTo ErrorHandler
    
    Dim strSearchName As String
    Dim blnIsShared As Boolean
    Dim strCriteria As String
    
    ' Get search name
    strSearchName = InputBox( _
        "Enter a name for this search:", _
        "Save Search", _
        m_strCurrentSavedSearch)
    
    If Len(Trim(strSearchName)) = 0 Then Exit Sub
    
    ' Ask if shared
    blnIsShared = (MsgBox( _
        "Make this search available to all users?", _
        vbQuestion + vbYesNo, _
        "Share Search") = vbYes)
    
    ' Serialize criteria
    strCriteria = SerializeSearchCriteria()
    
    ' Save to database
    Call SaveSearchToDatabase(strSearchName, strCriteria, blnIsShared)
    
    ' Refresh list
    Call LoadSavedSearches
    
    ' Select newly saved search
    Me.cboSavedSearches.value = DMax("SearchID", "SavedSearches", _
        "SearchName = '" & Replace(strSearchName, "'", "''") & "' AND UserID = " & modGlobals.UserID)
    
    MsgBox "Search saved successfully.", vbInformation
    
    Exit Sub
    
ErrorHandler:
    modUtilities.LogError "frmAdvancedSearch.cmdSaveSearch_Click", Err.number, Err.Description
    MsgBox "Error saving search: " & Err.Description, vbCritical
End Sub

Private Sub cmdDeleteSearch_Click()
    On Error GoTo ErrorHandler
    
    If IsNull(Me.cboSavedSearches) Then
        MsgBox "Please select a saved search to delete.", vbInformation
        Exit Sub
    End If
    
    ' Check ownership
    Dim lngOwnerID As Long
    lngOwnerID = Nz(DLookup("UserID", "SavedSearches", "SearchID = " & Me.cboSavedSearches), 0)
    
    If lngOwnerID <> modGlobals.UserID Then
        MsgBox "You can only delete your own saved searches.", vbExclamation
        Exit Sub
    End If
    
    ' Confirm deletion
    If MsgBox("Delete this saved search?", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    ' Delete
    CurrentDb.Execute "DELETE FROM SavedSearches WHERE SearchID = " & Me.cboSavedSearches, dbFailOnError
    
    ' Refresh
    Me.cboSavedSearches.value = Null
    Call LoadSavedSearches
    
    MsgBox "Search deleted.", vbInformation
    
    Exit Sub
    
ErrorHandler:
    modUtilities.LogError "frmAdvancedSearch.cmdDeleteSearch_Click", Err.number, Err.Description
End Sub

' CRITERIA SERIALIZATION (Simple pipe-delimited format)
Private Function SerializeSearchCriteria() As String
    On Error GoTo ErrorHandler
    
    Dim strResult As String
    
    ' Format: Field=Value|Field=Value|...
    If Not IsNull(Me.txtName) Then strResult = strResult & "Name=" & Me.txtName & "|"
    If Not IsNull(Me.txtVATNumber) Then strResult = strResult & "VAT=" & Me.txtVATNumber & "|"
    If Not IsNull(Me.txtTelephone) Then strResult = strResult & "Phone=" & Me.txtTelephone & "|"
    If Not IsNull(Me.txtEmail) Then strResult = strResult & "Email=" & Me.txtEmail & "|"
    If Not IsNull(Me.cboCountry) Then strResult = strResult & "Country=" & Me.cboCountry & "|"
    If Not IsNull(Me.txtAddress) Then strResult = strResult & "AddressLine=" & Me.txtAddress Or "City=" & Me.txtAddress Or "ZIPCode=" & Me.txtAddress & "|"
    If Not IsNull(Me.cboVATApplied) Then strResult = strResult & "VATApplied=" & Me.cboVATApplied & "|"
    
    If m_strSearchType = "Supplier" Then
        If Not IsNull(Me.txtIRPFFrom) Then strResult = strResult & "IRPFFrom=" & Me.txtIRPFFrom & "|"
        If Not IsNull(Me.txtIRPFTo) Then strResult = strResult & "IRPFTo=" & Me.txtIRPFTo & "|"
        If Not IsNull(Me.cboServiceType) Then strResult = strResult & "ServiceType=" & Me.cboServiceType & "|"
    End If
    
    ' Remove trailing pipe
    If Right(strResult, 1) = "|" Then strResult = Left(strResult, Len(strResult) - 1)
    
    SerializeSearchCriteria = strResult
    Exit Function
    
ErrorHandler:
    modUtilities.LogError "frmAdvancedSearch.SerializeSearchCriteria", Err.number, Err.Description
    SerializeSearchCriteria = ""
End Function

Private Function ParseSearchCriteria(strCriteria As String) As Object
    On Error GoTo ErrorHandler
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim arrPairs() As String
    Dim arrKeyVal() As String
    Dim i As Integer
    
    arrPairs = Split(strCriteria, "|")
    
    For i = LBound(arrPairs) To UBound(arrPairs)
        If InStr(arrPairs(i), "=") > 0 Then
            arrKeyVal = Split(arrPairs(i), "=")
            If UBound(arrKeyVal) >= 1 Then
                dict(Trim(arrKeyVal(0))) = Trim(arrKeyVal(1))
            End If
        End If
    Next i
    
    Set ParseSearchCriteria = dict
    Exit Function
    
ErrorHandler:
    modUtilities.LogError "frmAdvancedSearch.ParseSearchCriteria", Err.number, Err.Description
    Set ParseSearchCriteria = Nothing
End Function

Private Sub PopulateCriteriaFromDict(dict As Object)
    On Error Resume Next
    
    If dict.Exists("Name") Then Me.txtName = dict("Name")
    If dict.Exists("VAT") Then Me.txtVATNumber = dict("VAT")
    If dict.Exists("Phone") Then Me.txtTelephone = dict("Phone")
    If dict.Exists("Email") Then Me.txtEmail = dict("Email")
    If dict.Exists("Country") Then Me.cboCountry = dict("Country")
    If dict.Exists("Address") Then Me.txtAddress = dict("Address")
    If dict.Exists("VATApplied") Then Me.cboVATApplied = dict("VATApplied")
    
    If m_strSearchType = "Supplier" Then
        If dict.Exists("IRPFFrom") Then Me.txtIRPFFrom = dict("IRPFFrom")
        If dict.Exists("IRPFTo") Then Me.txtIRPFTo = dict("IRPFTo")
        If dict.Exists("ServiceType") Then Me.cboServiceType = dict("ServiceType")
    End If
    
    On Error GoTo 0
End Sub

Private Sub SaveSearchToDatabase(strName As String, strCriteria As String, blnShared As Boolean)
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SavedSearches", dbOpenDynaset)
    
    rs.AddNew
    rs!SearchName = strName
    rs!SearchType = m_strSearchType
    rs!SearchCriteria = strCriteria
    rs!UserID = modGlobals.UserID
    rs!IsShared = blnShared
    rs!CreatedDate = Now()
    rs.Update
    
    rs.Close
    
    Exit Sub
    
ErrorHandler:
    modUtilities.LogError "frmAdvancedSearch.SaveSearchToDatabase", Err.number, Err.Description
    Err.Raise Err.number, , Err.Description
End Sub

' UTILITY FUNCTIONS
Private Sub LoadFilterControls()
    On Error Resume Next
    
    Dim tableName As String
    tableName = IIf(m_strSearchType = "Client", "Clients", "Suppliers")
    
    ' Countries
    Me.cboCountry.RowSource = _
        "SELECT DISTINCT Country FROM " & tableName & " " & _
        "WHERE IsDeleted = False AND Country Is Not Null " & _
        "ORDER BY Country"
    
    ' VAT Applied
    Me.cboVATApplied.RowSource = "All;Yes;No"
    
    ' Service Type (Suppliers only)
    If m_strSearchType = "Supplier" Then
        Me.cboServiceType.RowSource = _
            "SELECT DISTINCT TypeOfServices FROM Suppliers " & _
            "WHERE IsDeleted = False AND TypeOfServices Is Not Null " & _
            "ORDER BY TypeOfServices"
    End If
End Sub

Private Sub ClearSearchCriteria()
    On Error Resume Next
    
    Me.txtName = Null
    Me.txtVATNumber = Null
    Me.txtTelephone = Null
    Me.txtEmail = Null
    Me.cboCountry = Null
    Me.txtAddress = Null
    Me.cboVATApplied = Null
    Me.txtIRPFFrom = Null
    Me.txtIRPFTo = Null
    Me.cboServiceType = Null
    Me.cboSavedSearches = Null
    
    m_strCurrentSavedSearch = ""
End Sub

Private Sub cmdClear_Click()
    Call ClearSearchCriteria
    
    ' Clear results
    Me.subList_SearchResults.SourceObject = ""
    Me.lblResultCount.Caption = "Enter search criteria and click Search"
    Me.lblResultCount.ForeColor = vbBlack
    Me.cmdSelect.Enabled = False
End Sub

' FORM CLEANUP
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    ' Clean up temporary query
    CurrentDb.QueryDefs.Delete "qryAdvancedSearchResults_Temp"
End Sub

' KEYBOARD SHORTCUTS
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If (Shift And acCtrlMask) <> 0 Then
                KeyCode = 0
                Call cmdSearch_Click
            End If
            
        Case vbKeyEscape
            KeyCode = 0
            Call cmdCancel_Click
    End Select
End Sub

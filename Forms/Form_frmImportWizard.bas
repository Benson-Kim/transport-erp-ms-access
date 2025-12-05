'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form      : frmImportWizard
' Purpose   : Multi-step wizard for importing Clients or Suppliers from Excel
' Author    : Tifamovs Saeol <tifamovs@gmail.com> [Benson Kimathi]
' Date      : 2025-11-30
' Usage     : DoCmd.OpenForm "frmImportWizard", , , , , , "Client" or "Supplier"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Database
Option Explicit

' MODULE-LEVEL VARIABLES
Private m_strEntityType As String           ' Singular entity type
Private m_strSourcePath As String           ' Selected Excel file
Private m_intCurrentStep As Integer         ' Current wizard step (1-5)
Private m_ImportResult As modDatabase.ImportResult
Private m_blnUpdateExisting As Boolean

' Wizard steps
Private Const STEP_SELECT_FILE As Integer = 1
Private Const STEP_PREVIEW As Integer = 2
Private Const STEP_VALIDATION As Integer = 3
Private Const STEP_IMPORT As Integer = 4
Private Const STEP_COMPLETE As Integer = 5

' FORM EVENTS

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    ' Get entity type from OpenArgs
    If Not IsNull(Me.OpenArgs) Then
        m_strEntityType = CStr(Me.OpenArgs)
    Else
        m_strEntityType = "Client"
    End If
    
    ' Configure form
    Me.Caption = "Import " & m_strEntityType & "s from Excel"
    Me.lblTitle.Caption = "Import " & m_strEntityType & "s - Step 1 of 5"
    
    ' Initialize
    m_intCurrentStep = STEP_SELECT_FILE
    m_blnUpdateExisting = False
    
    Call ShowStep(STEP_SELECT_FILE)
    
    Exit Sub
    
ErrorHandler:
    modUtilities.LogError "frmImportWizard.Form_Load", Err.number, Err.Description
End Sub

' STEP NAVIGATION

Private Sub ShowStep(StepNumber As Integer)
    On Error Resume Next
    
    m_intCurrentStep = StepNumber
    
    ' Update title
    Me.lblTitle.Caption = "Import " & m_strEntityType & "s - Step " & StepNumber & " of 5"
    
    ' Hide all step panels
    Me.pnlStep1.Visible = False
    Me.pnlStep2.Visible = False
    Me.pnlStep3.Visible = False
    Me.pnlStep4.Visible = False
    Me.pnlStep5.Visible = False
    
    ' Show current step
    Select Case StepNumber
        Case STEP_SELECT_FILE
            Me.pnlStep1.Visible = True
            Me.cmdBack.Enabled = False
            Me.cmdNext.Enabled = (Len(m_strSourcePath) > 0)
            Me.cmdNext.Caption = "Next >"
            
        Case STEP_PREVIEW
            Me.pnlStep2.Visible = True
            Me.cmdBack.Enabled = True
            Me.cmdNext.Enabled = True
            Me.cmdNext.Caption = "Next >"
            Call LoadPreview
            
        Case STEP_VALIDATION
            Me.pnlStep3.Visible = True
            Me.cmdBack.Enabled = True
            Me.cmdNext.Enabled = True
            Me.cmdNext.Caption = "Import"
            Call RunValidation
            
        Case STEP_IMPORT
            Me.pnlStep4.Visible = True
            Me.cmdBack.Enabled = False
            Me.cmdNext.Enabled = False
            Call ExecuteImport
            
        Case STEP_COMPLETE
            Me.pnlStep5.Visible = True
            Me.cmdBack.Enabled = False
            Me.cmdNext.Caption = "Finish"
            Me.cmdNext.Enabled = True
            Call ShowResults
    End Select
End Sub

Private Sub cmdNext_Click()
    On Error GoTo ErrorHandler
    
    Select Case m_intCurrentStep
        Case STEP_SELECT_FILE
            If Len(m_strSourcePath) = 0 Then
                MsgBox "Please select a file first.", vbExclamation
                Exit Sub
            End If
            Call ShowStep(STEP_PREVIEW)
            
        Case STEP_PREVIEW
            Call ShowStep(STEP_VALIDATION)
            
        Case STEP_VALIDATION
            Call ShowStep(STEP_IMPORT)
            
        Case STEP_COMPLETE
            DoCmd.Close acForm, Me.Name
    End Select
    
    Exit Sub
    
ErrorHandler:
    modUtilities.LogError "frmImportWizard.cmdNext_Click", Err.number, Err.Description
End Sub

Private Sub cmdBack_Click()
    If m_intCurrentStep > STEP_SELECT_FILE Then
        Call ShowStep(m_intCurrentStep - 1)
    End If
End Sub

Private Sub cmdCancel_Click()
    If MsgBox("Cancel import?", vbQuestion + vbYesNo) = vbYes Then
        DoCmd.Close acForm, Me.Name
    End If
End Sub

' STEP 1: FILE SELECTION

Private Sub cmdBrowse_Click()
    On Error GoTo ErrorHandler
    
    Dim varFileName As Office.FileDialog
    
    Set varFileName = Application.FileDialog(msoFileDialogFilePicker)
    With varFileName
        .InitialFileName = ".xlsx"
        .Title = "Select " & m_strEntityType & " Import File"
        .Filters.Clear
        .Filters.Add "Excel Files (*.xlsx;*.xls)", "*.xlsx;*.xls"
    End With
    
    If varFileName <> 0 Then
        m_strSourcePath = Application.FileDialog(msoFileDialogFilePicker).SelectedItems(1)
        Me.txtFilePath.value = m_strSourcePath
        Me.cmdNext.Enabled = True
    End If
    
    Exit Sub
    
ErrorHandler:
    ' Fallback method if FileDialog fails
'    m_strSourcePath = Application.GetOpenFilename( _
'        FileFilter:="Excel Files (*.xlsx;*.xls),*.xlsx;*.xls", _
'        Title:="Select " & m_strEntityType & " Import File")
'
    If m_strSourcePath <> "False" Then
        Me.txtFilePath.value = m_strSourcePath
        Me.cmdNext.Enabled = True
    End If
End Sub

Private Sub chkUpdateExisting_AfterUpdate()
    m_blnUpdateExisting = Me.chkUpdateExisting
End Sub

' STEP 2: PREVIEW DATA

Private Sub LoadPreview()
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim strTempTable As String
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    strTempTable = "tblImportPreview_Temp"
    
    ' Delete existing temp table
    On Error Resume Next
    DoCmd.DeleteObject acTable, strTempTable
    On Error GoTo ErrorHandler
    
    ' Import first 10 rows for preview
    DoCmd.TransferSpreadsheet _
        acImport, _
        acSpreadsheetTypeExcel12Xml, _
        strTempTable, _
        m_strSourcePath, _
        True
    
    ' Set subform source
    Me.subPreview.SourceObject = ""
    Me.subPreview.SourceObject = "Table." & strTempTable
    
    ' Count total rows
    Set rs = db.OpenRecordset(strTempTable, dbOpenSnapshot)
    rs.MoveLast
    Me.lblPreviewCount.Caption = "Showing first 10 of " & rs.RecordCount & " rows"
    rs.Close
    
    ' Show column mapping info
    Me.lblMappingInfo.Caption = _
        "Columns detected: " & db.TableDefs(strTempTable).Fields.Count & vbCrLf & _
        "Verify that column names match expected format."
    
    Exit Sub
    
ErrorHandler:
    modUtilities.LogError "frmImportWizard.LoadPreview", Err.number, Err.Description
    MsgBox "Error loading preview: " & Err.Description, vbCritical
End Sub

' STEP 3: VALIDATION

Private Sub RunValidation()
    On Error GoTo ErrorHandler
    
    Me.lblValidationStatus.Caption = "Running validation checks..."
    Me.lblValidationStatus.ForeColor = vbBlue
    DoEvents
    
    ' Perform quick validation scan
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strTempTable As String
    Dim lngTotalRows As Long
    Dim lngErrorCount As Long
    Dim colErrors As Collection
    
    Set colErrors = New Collection
    Set db = CurrentDb
    strTempTable = "tblImportPreview_Temp"
    
    Set rs = db.OpenRecordset(strTempTable, dbOpenSnapshot)
    
    If rs.EOF Then
        Me.lblValidationStatus.Caption = "No data to validate"
        Me.lblValidationStatus.ForeColor = vbRed
        Exit Sub
    End If
    
    rs.MoveLast
    lngTotalRows = rs.RecordCount
    rs.MoveFirst
    
    ' Validate first 20 rows (quick check)
    Dim i As Integer
    i = 0
    
    Do While Not rs.EOF And i < 20
        i = i + 1
        
        ' Basic validation
        If m_strEntityType = "Client" Then
            Dim validResult As modClientSupplierForms.ValidationResult
            validResult = modClientSupplierForms.ValidateClientData( _
                Null, _
                Nz(rs!ClientName, ""), _
                Nz(rs!VATNumber, ""), _
                Nz(rs!EmailBilling, ""), _
                Nz(rs!Telephone, ""), _
                Nz(rs!Address, ""), _
                Nz(rs!Country, ""))
            
            If Not validResult.isValid Then
                lngErrorCount = lngErrorCount + 1
                
                Dim errMsg As String
                errMsg = "Row " & (i + 1) & ": "
                
                Dim errItem As Variant
                For Each errItem In validResult.ErrorMessages
                    errMsg = errMsg & errItem & "; "
                Next errItem
                
                If colErrors.Count < 10 Then  ' Show max 10 errors
                    colErrors.Add errMsg
                End If
            End If
        Else
            ' Supplier validation
            Dim validResultSup As modClientSupplierForms.ValidationResult
            validResultSup = modClientSupplierForms.ValidateSupplierData( _
                Null, _
                Nz(rs!SupplierName, ""), _
                Nz(rs!VATNumber, ""), _
                Nz(rs!Email, ""), _
                Nz(rs!Telephone, ""), _
                Nz(rs!AddressLine, ""), _
                Nz(rs!Country, ""), _
                Nz(rs!IRPFPercentage, 0))
            
            If Not validResultSup.isValid Then
                lngErrorCount = lngErrorCount + 1
                
                errMsg = "Row " & (i + 1) & ": "
                For Each errItem In validResultSup.ErrorMessages
                    errMsg = errMsg & errItem & "; "
                Next errItem
                
                If colErrors.Count < 10 Then
                    colErrors.Add errMsg
                End If
            End If
        End If
        
        rs.MoveNext
    Loop
    
    rs.Close
    
    ' Display validation results
    If lngErrorCount = 0 Then
        Me.lblValidationStatus.Caption = "? All validated rows passed (" & i & " checked)"
        Me.lblValidationStatus.ForeColor = RGB(0, 128, 0)
        Me.txtValidationErrors.value = "No errors found in sample."
        Me.cmdNext.Enabled = True
    Else
        Me.lblValidationStatus.Caption = "? " & lngErrorCount & " errors found in sample"
        Me.lblValidationStatus.ForeColor = RGB(200, 0, 0)
        
        Dim strAllErrors As String
        For Each errItem In colErrors
            strAllErrors = strAllErrors & errItem & vbCrLf
        Next errItem
        
        Me.txtValidationErrors.value = strAllErrors
        
        If MsgBox("Errors detected in preview." & vbCrLf & vbCrLf & _
                  "Continue and skip error rows?", vbQuestion + vbYesNo) = vbYes Then
            Me.cmdNext.Enabled = True
        Else
            Me.cmdNext.Enabled = False
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    modUtilities.LogError "frmImportWizard.RunValidation", Err.number, Err.Description
    Me.lblValidationStatus.Caption = "Validation error: " & Err.Description
    Me.lblValidationStatus.ForeColor = vbRed
End Sub

' STEP 4: EXECUTE IMPORT

Private Sub ExecuteImport()
    On Error GoTo ErrorHandler
    
    Me.lblImportStatus.Caption = "Importing data..."
    Me.lblImportStatus.ForeColor = vbBlue
    DoEvents
    
    ' Execute import
    If m_strEntityType = "Client" Then
        m_ImportResult = modDatabase.ImportClientsFromExcel(m_strSourcePath, m_blnUpdateExisting)
    Else
        m_ImportResult = modDatabase.ImportSuppliersFromExcel(m_strSourcePath, m_blnUpdateExisting)
    End If
    
    ' Move to results
    Call ShowStep(STEP_COMPLETE)
    
    Exit Sub
    
ErrorHandler:
    modUtilities.LogError "frmImportWizard.ExecuteImport", Err.number, Err.Description
    Me.lblImportStatus.Caption = "Import failed: " & Err.Description
    Me.lblImportStatus.ForeColor = vbRed
End Sub

' STEP 5: SHOW RESULTS

Private Sub ShowResults()
    On Error Resume Next
    
    With m_ImportResult
        Me.lblResultSummary.Caption = _
            "Import Complete!" & vbCrLf & vbCrLf & _
            "Total Rows: " & .TotalRows & vbCrLf & _
            "? Successful: " & .SuccessCount & vbCrLf & _
            "? Errors: " & .ErrorCount & vbCrLf & _
            "? Skipped: " & .SkippedCount
        
        ' Show error details if any
        If .ErrorCount > 0 Or .SkippedCount > 0 Then
            Dim strErrors As String
            Dim errItem As Variant
            
            For Each errItem In .ErrorDetails
                strErrors = strErrors & errItem & vbCrLf
            Next errItem
            
            Me.txtResultErrors.value = strErrors
            Me.txtResultErrors.Visible = True
            Me.lblErrorsLabel.Visible = True
        Else
            Me.txtResultErrors.Visible = False
            Me.lblErrorsLabel.Visible = False
        End If
    End With
End Sub

Private Sub cmdExportErrors_Click()
    ' Export error log to text file
    If m_ImportResult.ErrorCount = 0 And m_ImportResult.SkippedCount = 0 Then
        MsgBox "No errors to export.", vbInformation
        Exit Sub
    End If
    
    Dim strPath As String
    strPath = CurrentProject.Path & "\ImportErrors_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
    
    Dim intFile As Integer
    intFile = FreeFile
    
    Open strPath For Output As #intFile
    
    Print #intFile, "Import Error Log - " & Now
    Print #intFile, "Entity Type: " & m_strEntityType
    Print #intFile, "Source File: " & m_strSourcePath
    Print #intFile, String(50, "=")
    Print #intFile, ""
    
    Dim errItem As Variant
    For Each errItem In m_ImportResult.ErrorDetails
        Print #intFile, errItem
    Next errItem
    
    Close #intFile
    
    MsgBox "Error log exported to:" & vbCrLf & vbCrLf & strPath, vbInformation
End Sub

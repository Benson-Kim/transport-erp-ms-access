'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MODULE: modUtilities
' PURPOSE: Miscellaneous helper functions + centralized error handling
' AUTHOR: Tifamovs Saeol <tifamovs@gmail.com> [Benson Kimathi]
' CREATED: November 17, 2025
' UPDATED: December 5, 2025
'          - Added ExportAllModules and ImportModule for dev convenience
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Compare Database
Option Explicit

Private Const MODULE_NAME As String = "modUtilities"


' PUBLIC: LogError � Centralized error logging with fallback to file

''
''' Procedure: LogError
''' Purpose   : Write comprehensive error details to ErrorLog table
'''             with automatic fallback to local text file if DB unavailable
''' Parameters:
'''   strProcName     - Name of procedure where error occurred
'''   lngErrNumber    - Err.Number
'''   strErrDesc      - Err.Description
'''   Optional strAdditionalInfo - Any context (SQL, record ID, etc.)
'''   Optional lngLine        - Line number (requires numbered lines or Erl)
''
Public Sub LogError(ByVal strProcName As String, _
                    ByVal lngErrNumber As Long, _
                    ByVal strErrDesc As String, _
                    Optional ByVal strAdditionalInfo As String = "", _
                    Optional ByVal lngLine As Long = 0)

    On Error GoTo FallbackFileLog
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim strSource As String

    
    strSource = strProcName
    If lngLine > 0 Then strSource = strSource & " (Line " & lngLine & ")"
    
    Set db = CurrentDb
    
    strSQL = "INSERT INTO ErrorLog " & _
             "(ErrorNumber, ErrorDescription, ErrorSource, ErrorLine, " & _
             "UserID, ErrorDate, WorkstationName, AdditionalInfo) VALUES " & _
             "(" & lngErrNumber & ", " & _
             "'" & Replace(strErrDesc, "'", "''") & "', " & _
             "'" & Left(Replace(strSource, "'", "''"), 100) & "', " & _
             lngLine & ", " & _
             Nz(g_lngUserID, 0) & ", " & _
             "Now(), " & _
             "'" & Left(Environ("COMPUTERNAME"), 50) & "', " & _
             "'" & Left(Replace(Nz(strAdditionalInfo, ""), "'", "''"), 65535) & "')"
        
    db.Execute strSQL, dbFailOnError
    Exit Sub
    
FallbackFileLog:
    ' If we cannot write to the database at all, fall back to local text file
    Dim strLogPath As String
    Dim intFile As Integer
    
    strLogPath = GetLocalAppDataPath() & "\RoadFreightERP_ErrorLog.txt"
    
    intFile = FreeFile
    Open strLogPath For Append As #intFile
    Print #intFile, "------ ERROR ------" & vbCrLf & _
                    "DateTime      : " & Now() & vbCrLf & _
                    "UserID        : " & g_lngUserID & vbCrLf & _
                    "Workstation   : " & Environ("COMPUTERNAME") & vbCrLf & _
                    "Procedure     : " & strProcName & IIf(lngLine > 0, " (Line " & lngLine & ")", "") & vbCrLf & _
                    "Error #       : " & lngErrNumber & vbCrLf & _
                    "Description   : " & strErrDesc & vbCrLf & _
                    "Additional    : " & strAdditionalInfo & vbCrLf & _
                    String(50, "=") & vbCrLf
    Close #intFile
End Sub

Public Sub AddToFilter(ByRef strF As String, strNew As String)
    If strF = "" Then strF = strNew Else strF = strF & " AND " & strNew
End Sub

Public Sub RestartSearchTimer(frm As Form)
    If Not frm Is Nothing Then
        DoEvents
        frm.TimerInterval = modGlobals.SEARCH_DELAY_MS
    End If
End Sub

' PRIVATE: Helper to get reliable local folder for fallback logging
Private Function GetLocalAppDataPath() As String
    Dim strPath As String
    strPath = Environ("APPDATA")
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    strPath = strPath & "RoadFreightERP\"
    
    If Dir(strPath, vbDirectory) = "" Then
        MkDir strPath
    End If
    
    GetLocalAppDataPath = strPath
End Function

' Helper � safe because not every form has every button
Private Sub SafeEnable(ctl As control, blnEnabled As Boolean)
    On Error Resume Next

    ' Store original color if not stored yet
    If ctl.Tag = "" Then ctl.Tag = CStr(ctl.ForeColor)

    Dim enabledColor As Long
    enabledColor = CLng(ctl.Tag)

    ctl.Enabled = blnEnabled
    ctl.ForeColor = IIf(blnEnabled, enabledColor, vbGrayText)

    On Error GoTo 0
End Sub

Public Sub SetButtonStates(frm As Form, _
                           Optional permEdit As String = "", _
                           Optional permDelete As String = "", _
                           Optional forceDirty As Boolean = False)

    On Error Resume Next
    
    ' Determine state
    Dim isNew As Boolean:    isNew = frm.NewRecord
    Dim hasRecord As Boolean: hasRecord = Not isNew
    Dim dirty As Boolean
    
    If forceDirty Then dirty = True Else dirty = frm.dirty
    
    Dim canEdit As Boolean:  canEdit = (permEdit <> "") And HasPermission(permEdit)
    Dim canDelete As Boolean: canDelete = (permDelete <> "") And HasPermission(permDelete)
    
    ' New / Duplicate / Delete buttons
    SafeEnable frm.cmdNew, True
    SafeEnable frm.cmdDuplicate, hasRecord And canEdit And Not dirty
    SafeEnable frm.cmdDelete, hasRecord And (Not isNew) And canDelete And Not dirty
    
    ' Save / Undo / Cancel buttons
    SafeEnable frm.cmdSave, dirty And canEdit
    SafeEnable frm.cmdCancel, dirty
    SafeEnable frm.cmdUndo, dirty And Not isNew

    ' Other buttons (export, print, refresh)
    SafeEnable frm.cmdExportExcel, hasRecord And Not dirty
    SafeEnable frm.cmdPrint, hasRecord And Not dirty
    SafeEnable frm.cmdAdvancedSearch, Not dirty
    SafeEnable frm.cmdRefresh, Not dirty
End Sub

Public Sub CopyToClipboard(ByVal strText As String)
    With CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}") ' HTML clipboard
        .SetText strText
        .PutInClipboard
    End With
    ' Optional: beep or toast notification
    Beep
End Sub

Public Sub ShowCopiedFeedback(lbl As Label)
    Dim origCaption As String
    Dim origForeColor As Long
    
    ' Save original appearance
    origCaption = lbl.Caption
    origForeColor = lbl.ForeColor
    
    ' Show success
    lbl.Caption = "Copied!"
    lbl.ForeColor = RGB(14, 76, 73)
    
    ' Force repaint
'    lbl.Repaint
    DoEvents
    
    ' Wait 1.5 seconds (non-blocking with DoEvents)
    Dim pauseTime As Double
    pauseTime = Timer + 1.5
    Do While Timer < pauseTime
        DoEvents
    Loop
    
    ' Restore original
    lbl.Caption = origCaption
    lbl.ForeColor = origForeColor
    lbl.FontBold = False
End Sub

' ------------------------------------------------------------------
' FUNCTION: NormalizeUsername
' PURPOSE : Standardize username input for consistent comparison
'           - Trims whitespace
'           - Converts to UpperCase
'           - Removes common typos (accents via StrConv if needed later)
' RETURNS : Clean uppercase username
' USAGE   : Called from login, user creation, and any username lookup
' ------------------------------------------------------------------
Public Function NormalizeUsername(ByVal strInput As String) As String
    
    On Error GoTo ErrorHandler
    
    Dim strClean As String
    
    strClean = Trim$(Nz(strInput, ""))
    
    If Len(strClean) = 0 Then
        NormalizeUsername = ""
        Exit Function
    End If
    
    ' Convert to uppercase for case-insensitive matching
    strClean = UCase$(strClean)
    
    ' Optional: Remove diacritics (e.g., Jos� ? JOSE) � uncomment if needed
    ' strClean = StrConv(strClean, vbUnicode)
    ' strClean = StrConv(strClean, vbFromUnicode)
    ' strClean = Replace(strClean, Chr(0), "")
    
    NormalizeUsername = strClean
    
    Exit Function
    
ErrorHandler:
    modUtilities.LogError MODULE_NAME & ".NormalizeUsername", Err.number, Err.Description, "Input=" & strInput
    NormalizeUsername = ""
End Function

Public Function GetPropertyBool(frm As Form, propName As String) As Boolean
    On Error Resume Next
    GetPropertyBool = frm.Controls(propName)
    GetPropertyBool = frm(propName)
End Function

' '''''''''''''''''' DISPLAY HELPER FUNCTIONS ''''''''''''''''''

Public Function FormatValidationErrors(result As ValidationResult) As String
    On Error GoTo ErrorHandler
    
    If result.isValid Then
        FormatValidationErrors = ""
        Exit Function
    End If
    
    Dim strMsg As String
    Dim errItem As Variant
    Dim intCount As Integer
    
    strMsg = "Please correct the following error"
    
    If result.ErrorMessages.Count > 1 Then
        strMsg = strMsg & "s:" & vbCrLf & vbCrLf
    Else
        strMsg = strMsg & ":" & vbCrLf & vbCrLf
    End If
    
    intCount = 0
    For Each errItem In result.ErrorMessages
        intCount = intCount + 1
        strMsg = strMsg & intCount & ". " & errItem & vbCrLf
     Next errItem
     strMsg = strMsg & vbCrLf & "Please review the highlighted fields."
     
     FormatValidationErrors = strMsg
     Exit Function
ErrorHandler:
     modUtilities.LogError MODULE_NAME & ".FormatValidationErrors", Err.number, Err.Description
     FormatValidationErrors = "An error occurred while formatting validation errors."
End Function

' '''''''''''''''''' HIGHLIGHT ERROR FIELDS (CALL AFTER MESSAGE) ''''''''''''''''''
Public Sub HighlightValidationErrors(frm As Form, result As ValidationResult)
    On Error Resume Next
    
    Dim fld As Variant
    Dim ctl As control
   
    If frm Is Nothing Then Exit Sub
    If result.isValid Then Exit Sub
    
    ' Highlight each error field
    For Each fld In result.FieldsWithErrors
        Set ctl = Nothing
        Set ctl = frm.Controls(CStr(fld))
        
        If Not ctl Is Nothing Then
            If ctl.ControlType = acTextBox Or _
               ctl.ControlType = acComboBox Or _
               ctl.ControlType = acListBox Then
               
                    ctl.BorderColor = RGB(242, 80, 34) ' Light red border
            End If
        End If
    Next fld

    Set ctl = Nothing
End Sub

' '''''''''''''''''' CLEAR ERROR HIGHLIGHTING ''''''''''''''''''
Public Sub ClearValidationHighlights(frm As Form)
    On Error Resume Next
    
    Dim ctl As control
    
    If frm Is Nothing Then Exit Sub
    
    ' Reset border colors to default (black)
    For Each ctl In frm.Controls
        If ctl.ControlType = acTextBox Or _
           ctl.ControlType = acComboBox Or _
           ctl.ControlType = acListBox Then
            ctl.BorderColor = RGB(217, 217, 217)  ' Background 1, Darker 15%
        End If
    Next ctl
    
    Set ctl = Nothing
End Sub

Public Function IsLoaded(strFormName As String) As Boolean
    Dim frm As AccessObject
    
    For Each frm In CurrentProject.AllForms
        If frm.Name = strFormName Then
            IsLoaded = (frm.IsLoaded = True)
            Exit Function
        End If
    Next frm
    
    IsLoaded = False
End Function

Public Function TableExists(tableName As String) As Boolean
    On Error Resume Next
    Dim tdf As DAO.TableDef
    Set tdf = CurrentDb.TableDefs(tableName)
    TableExists = (Err.number = 0)
End Function

Public Function GetSaveAsPath(Optional ByVal defaultName As String = "Export.xlsx") As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogSaveAs)

    With fd
        .Title = "Save File As"
        .InitialFileName = defaultName

        If .Show = -1 Then
            GetSaveAsPath = .SelectedItems(1)
        Else
            GetSaveAsPath = vbNullString
        End If
    End With
End Function

Public Sub ExportAllModules()
    Dim vbComp As VBIDE.VBComponent
    Dim exportPath As String
    
    ' Folder to export to (change as needed)
    exportPath = CurrentProject.path & "\ExportedModules\"
    
    ' Create folder if needed
    If Dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If
    
    ' Loop through all VB components
    For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
        Select Case vbComp.Type
            
            Case vbext_ct_StdModule
                vbComp.Export exportPath & vbComp.Name & ".bas"
                
            Case vbext_ct_ClassModule
                vbComp.Export exportPath & vbComp.Name & ".cls"
                
            Case vbext_ct_MSForm
                vbComp.Export exportPath & vbComp.Name & ".frm"
                
            Case vbext_ct_Document
                ' Form/report code-behind
                vbComp.Export exportPath & vbComp.Name & ".bas"
                
        End Select
    Next
    
    MsgBox "Export complete! Files saved to: " & exportPath
End Sub

Public Sub ImportModule(path As String)
    Application.VBE.ActiveVBProject.VBComponents.Import path
End Sub



' STANDARD ERROR HANDLER TEMPLATE (copy-paste into every procedure)

' Place this exact block in every Public/Private Sub or Function
'
'    On Error GoTo ErrorHandler
'    '--- MAIN CODE HERE ---------------------------------------------
'
'ExitHandler:
'    ' Cleanup code that MUST run (close recordsets, set objects = Nothing)
'    Set rs = Nothing
'    Set db = Nothing
'    Exit Sub
'
'ErrorHandler:
'    Dim strMsg As String
'    Select Case Err.Number
'        Case 1234   ' Example of recoverable error
'            strMsg = "Friendly message for user"
'            MsgBox strMsg, vbExclamation, APP_NAME
'            Resume ExitHandler
'        Case Else
'            ' Non-recoverable � log + graceful shutdown
'            LogError "ProcedureName", Err.Number, Err.Description, "Any context info", Erl
'            strMsg = "A critical error has occurred. The application will now close " & _
'                     "to prevent data corruption." & vbCrLf & vbCrLf & _
'                     "Please contact your system administrator."
'            MsgBox strMsg, vbCritical, APP_NAME & " - Critical Error"
'            DoCmd.Quit acQuitSaveNone
'    End Select
'


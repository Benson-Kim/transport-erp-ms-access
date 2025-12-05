Attribute VB_Name = "modUtilities"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MODULE: modUtilities
' PURPOSE: Miscellaneous helper functions + centralized error handling
' AUTHOR: Expert Back-End Developer (MS Access VBA Security & Reliability)
' CREATED: November 17, 2025
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Compare Database
Option Explicit

Private Const MODULE_NAME As String = "modUtilities"


' PUBLIC: LogError – Centralized error logging with fallback to file

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
    
    strSource = MODULE_NAME & "." & strProcName
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
    Print #intFile, "=== ERROR ===" & vbCrLf & _
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

' Helper – safe because not every form has every button
Private Sub SafeEnable(ctl As control, blnEnabled As Boolean)
    On Error Resume Next
    Dim origForeColor As Long
    
    origForeColor = ctl.ForeColor
    ctl.Enabled = blnEnabled
    ctl.ForeColor = IIf(blnEnabled, origForeColor, vbGrayText)
    On Error GoTo 0
End Sub

Public Sub SetButtonStates(frm As Form, _
                           Optional permEdit As String = "", _
                           Optional permDelete As String = "")

    On Error Resume Next
    
    '-------------------------------
    ' Determine state
    '-------------------------------
    Dim isNew As Boolean:    isNew = frm.NewRecord
    Dim hasRecord As Boolean: hasRecord = Not isNew
    Dim dirty As Boolean:    dirty = frm.dirty
    
    Dim canEdit As Boolean:  canEdit = (permEdit <> "") And HasPermission(permEdit)
    Dim canDelete As Boolean: canDelete = (permDelete <> "") And HasPermission(permDelete)
    
    '-------------------------------
    ' New / Duplicate / Delete buttons
    '-------------------------------
    SafeEnable frm.cmdNew, True
    SafeEnable frm.cmdDuplicate, hasRecord And canEdit
    SafeEnable frm.cmdDelete, hasRecord And (Not isNew) And canDelete
    
    '-------------------------------
    ' Save / Undo / Cancel buttons
    '-------------------------------
    If isNew Then
        ' For new records, always allow Save if user can edit
        SafeEnable frm.cmdSave, canEdit
        ' Cancel should always be enabled
        SafeEnable frm.cmdCancel, True
        ' Undo not applicable on new record
        SafeEnable frm.cmdUndo, False
    Else
        ' Existing record – follow dirty + permissions
        SafeEnable frm.cmdSave, dirty And canEdit
        SafeEnable frm.cmdUndo, dirty
        SafeEnable frm.cmdCancel, dirty
    End If

    '-------------------------------
    ' Other buttons (export, print, refresh)
    '-------------------------------
    SafeEnable frm.cmdExportExcel, hasRecord
    SafeEnable frm.cmdPrint, hasRecord
    SafeEnable frm.cmdAdvancedSearch, True
    SafeEnable frm.cmdRefresh, True
End Sub

Public Sub AddToFilter(ByRef strF As String, strNew As String)
    If strF = "" Then strF = strNew Else strF = strF & " AND " & strNew
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
'    lbl.Repaint
End Sub



' STANDARD ERROR HANDLER TEMPLATE (copy-paste into every procedure)

' Place this exact block in every Public/Private Sub or Function
'
'    On Error GoTo ErrorHandler
'    '=== MAIN CODE HERE =============================================
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
'            ' Non-recoverable – log + graceful shutdown
'            LogError "ProcedureName", Err.Number, Err.Description, "Any context info", Erl
'            strMsg = "A critical error has occurred. The application will now close " & _
'                     "to prevent data corruption." & vbCrLf & vbCrLf & _
'                     "Please contact your system administrator."
'            MsgBox strMsg, vbCritical, APP_NAME & " - Critical Error"
'            DoCmd.Quit acQuitSaveNone
'    End Select
'


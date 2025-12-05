'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MODULE    : modAudit
' PURPOSE   : Comprehensive audit trail framework with synchronous and
'          asynchronous logging, form-level change detection, and
'          Admin-only audit viewer support.
' SECURITY  : Only Admin can view. All actions are logged permanently.
' AUTHOR    : Tifamovs Saeol <tifamovs@gmail.com> [Benson Kimathi]
' CREATED   : November 17, 2025
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Compare Database
Option Explicit

Private Const MODULE_NAME As String = "modAudit"
Private Const BATCH_SIZE As Long = 50
Private Const MAX_QUEUE_SIZE As Long = 1000

' Queue for asynchronous (batched) audit logging
Private Type AuditQueueItem
    tableName As String
    RecordID As Long
    FieldName As String
    OldValue As Variant
    NewValue As Variant
    ActionType As String
    PerformedBy As Long
    Workstation As String
End Type

Private marrQueue() As AuditQueueItem
Private mlngQueueCount As Long

' INITIALIZE QUEUE (called once at app startup from modStartup)
Public Sub InitializeAuditQueue()
    ReDim marrQueue(1 To BATCH_SIZE)
    mlngQueueCount = 0
End Sub


' CORE: LogAudit � Synchronous or Queued depending on criticality
' Procedure: LogAudit
' Purpose  : Primary audit logging function � ALWAYS synchronous for
'            critical operations. Use LogAuditAsync for non-critical.
'
Public Sub LogAudit(ByVal strTableName As String, _
                    ByVal lngRecordID As Long, _
                    ByVal strFieldName As String, _
                    ByVal varOldValue As Variant, _
                    ByVal varNewValue As Variant, _
                    ByVal strActionType As String, _
                    Optional ByVal lngPerformedBy As Long = -1, _
                    Optional ByVal blnForceSync As Boolean = False)

    On Error GoTo ErrorHandler

    Dim lngUser As Long
    
    ' Determine user ID: override > current > 1 (system)
    If lngPerformedBy <> -1 Then
        lngUser = lngPerformedBy
    ElseIf modGlobals.UserID > 0 Then
        lngUser = modGlobals.UserID
    Else
        lngUser = 1  ' System/Test
    End If

    ' Critical actions (Insert/Delete/Login) or forced sync ? write immediately
    If blnForceSync Or strActionType = "Insert" Or strActionType = "Delete" Or strActionType = "PurgeOld" Then
        WriteAuditRecord strTableName, lngRecordID, strFieldName, varOldValue, varNewValue, strActionType, lngUser
    Else
        ' Queue for batch write (non-critical field updates)
        QueueAuditRecord strTableName, lngRecordID, strFieldName, varOldValue, varNewValue, strActionType, lngUser
    End If

    Exit Sub

ErrorHandler:
    modUtilities.LogError MODULE_NAME & ".LogAudit", Err.number, Err.Description, _
                "Context: Table=" & strTableName & ", Action=" & strActionType
End Sub


' WRAPPERS
Public Sub LogInsert(ByVal strTableName As String, ByVal lngRecordID As Long, Optional lngUser As Long = -1)
    LogAudit strTableName, lngRecordID, "", Null, Null, "Insert", lngUser, True
End Sub

Public Sub LogUpdate(ByVal strTableName As String, _
                     ByVal lngRecordID As Long, _
                     ByVal strFieldName As String, _
                     ByVal varOldValue As Variant, _
                     ByVal varNewValue As Variant, _
                     Optional lngUser As Long = -1)
    If Nz(varOldValue, "") <> Nz(varNewValue, "") Then
        LogAudit strTableName, lngRecordID, strFieldName, _
             varOldValue, varNewValue, "Update", lngUser, False
    End If
End Sub

Public Sub LogDelete(ByVal strTableName As String, ByVal lngRecordID As Long, Optional lngUser As Long = -1)
    LogAudit strTableName, lngRecordID, "", Null, Null, "Delete", lngUser, True
End Sub


' ASYNCHRONOUS (BATCHED) SUPPORT
Public Sub LogAuditAsync(ByVal strTableName As String, _
                         ByVal lngRecordID As Long, _
                         ByVal strFieldName As String, _
                         ByVal varOldValue As Variant, _
                         ByVal varNewValue As Variant, _
                         ByVal strActionType As String, _
                         ByVal lngUser As Long)
    LogAudit strTableName, lngRecordID, strFieldName, varOldValue, varNewValue, strActionType, lngUser, False
End Sub

Private Sub QueueAuditRecord(ByVal strTableName As String, _
                            ByVal lngRecordID As Long, _
                            ByVal strFieldName As String, _
                            ByVal varOldValue As Variant, _
                            ByVal varNewValue As Variant, _
                            ByVal strActionType As String, _
                            ByVal lngUser As Long)

    If mlngQueueCount >= MAX_QUEUE_SIZE Then FlushAuditQueue
    
    mlngQueueCount = mlngQueueCount + 1
    If mlngQueueCount > UBound(marrQueue) Then
        ReDim Preserve marrQueue(1 To UBound(marrQueue) + BATCH_SIZE)
    End If

    With marrQueue(mlngQueueCount)
        .tableName = Left(strTableName, 50)
        .RecordID = lngRecordID
        .FieldName = Left(Nz(strFieldName, ""), 50)
        .OldValue = varOldValue
        .NewValue = varNewValue
        .ActionType = Left(strActionType, 20)
        .PerformedBy = lngUser
        .Workstation = Left(Environ("COMPUTERNAME"), 50)
    End With

    ' Flush if queue is getting large
    If mlngQueueCount >= BATCH_SIZE Then FlushAuditQueue

End Sub

Public Sub FlushAuditQueue()
    If mlngQueueCount = 0 Then Exit Sub

'    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim i As Long

    Set db = CurrentDb

    modDatabase.BeginTransaction

    For i = 1 To mlngQueueCount
        Set qdf = db.CreateQueryDef("")
        qdf.sql = "INSERT INTO AuditLog (TableName, RecordID, FieldName, OldValue, NewValue, " & _
                  "ActionType, PerformedBy, PerformedDate, WorkstationName) " & _
                  "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)"

        With marrQueue(i)
            qdf.Parameters(0) = .tableName
            qdf.Parameters(1) = .RecordID
            qdf.Parameters(2) = .FieldName
            qdf.Parameters(3) = .OldValue
            qdf.Parameters(4) = .NewValue
            qdf.Parameters(5) = .ActionType
            qdf.Parameters(6) = .PerformedBy
            qdf.Parameters(7) = Now()
            qdf.Parameters(8) = .Workstation
        End With

        qdf.Execute dbFailOnError
    Next i

    modDatabase.CommitTransaction

    mlngQueueCount = 0
    Exit Sub

ErrorHandler:
    modDatabase.RollbackTransaction
    modUtilities.LogError MODULE_NAME & ".FlushAuditQueue", Err.number, Err.Description, _
                     "Count=" & mlngQueueCount
End Sub


Public Sub WriteAuditRecord(ByVal strTableName As String, _
                            ByVal lngRecordID As Long, _
                            ByVal strFieldName As String, _
                            ByVal varOldValue As Variant, _
                            ByVal varNewValue As Variant, _
                            ByVal strActionType As String, _
                            ByVal lngUser As Long)

    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    
    Set db = CurrentDb
    Set qdf = db.CreateQueryDef("")

    qdf.sql = "INSERT INTO AuditLog " & _
              "(TableName, RecordID, FieldName, OldValue, NewValue, ActionType, PerformedBy, PerformedDate, WorkstationName) " & _
              "VALUES ([pTableName], [pRecordID], [pFieldName], [pOldValue], [pNewValue], [pActionType], [pPerformedBy], [pPerformedDate], [pWorkstationName]);"

    qdf.Parameters![pTableName].value = Left(strTableName, 50)
    qdf.Parameters![pRecordID].value = lngRecordID
    qdf.Parameters![pFieldName].value = Left(Nz(strFieldName, ""), 50)
    qdf.Parameters![pOldValue].value = varOldValue
    qdf.Parameters![pNewValue].value = varNewValue
    qdf.Parameters![pActionType].value = Left(strActionType, 20)
    qdf.Parameters![pPerformedBy].value = lngUser
    qdf.Parameters![pPerformedDate].value = Now()
    qdf.Parameters![pWorkstationName].value = Left(Environ("COMPUTERNAME"), 50)
    
    qdf.Execute dbFailOnError

    Exit Sub

ErrorHandler:
    modUtilities.LogError MODULE_NAME & ".WriteAuditRecord", Err.number, Err.Description, _
                         "Table=" & strTableName & " | Action=" & strActionType
End Sub

' FORM-LEVEL AUTOMATIC AUDIT (BeforeUpdate)
'
' Procedure: AuditFormChanges
' Purpose  : Call from ANY form's Form_BeforeUpdate event
'             Automatically logs ALL changed fields
'
Public Sub AuditFormChanges(frm As Form, ByVal strTableName As String, ByVal lngRecordID As Long)
    On Error GoTo ErrorHandler

    Dim ctl As control
    Dim strField As String
    Dim varOld As Variant, varNew As Variant

    For Each ctl In frm.Controls
        If ctl.ControlType = acTextBox Or ctl.ControlType = acComboBox Or _
           ctl.ControlType = acCheckBox Or ctl.ControlType = acOptionGroup Then

            If ctl.Name Like "*_OldValue" Or ctl.Name = "AuditSkip" Then GoTo NextControl
            
            If Len(ctl.ControlSource) = 0 Then GoTo NextControl ' Skip unbound controls
            
            strField = Nz(ctl.Tag, ctl.Name)  ' Prefer Tag for real DB field name
            If strField = "" Then strField = ctl.Name

            ' Skip if no OldValue (new record) or control not bound
             If Not IsNull(ctl.OldValue) Then
                varOld = ctl.OldValue
                varNew = ctl.value

                If Nz(varOld, "") <> Nz(varNew, "") Then
                    LogAuditAsync strTableName, lngRecordID, strField, varOld, varNew, "Update", modGlobals.UserID
                End If
            End If
        End If
NextControl:
    Next ctl

    Exit Sub

ErrorHandler:
    modUtilities.LogError MODULE_NAME & ".AuditFormChanges", Err.number, Err.Description, _
                         "Form=" & frm.Name & " | RecordID=" & lngRecordID
End Sub

Public Sub PurgeOldAudits(Optional ByVal lngDaysToKeep As Long = 365)
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim strSQL As String

    Set db = CurrentDb

    If Not modPermissions.HasPermission(PERM_MANAGE_USERS) Then  ' Reuse admin perm or add new
        MsgBox "Administrator rights required.", vbCritical
        Exit Sub
    End If

    modDatabase.BeginTransaction

    strSQL = "DELETE FROM AuditLog WHERE PerformedDate < DateAdd('d', -" & lngDaysToKeep & ", Date())"
    db.Execute strSQL, dbFailOnError

    ' Log the purge (sync)
    LogAudit "AuditLog", 0, "", Null, Null, "PurgeOld", , True
    modDatabase.CommitTransaction

    Exit Sub
    
ErrorHandler:
    modDatabase.RollbackTransaction
    modUtilities.LogError MODULE_NAME & ".PurgeOldAudits", Err.number, Err.Description, "Days=" & lngDaysToKeep
    MsgBox "Purge failed: " & Err.Description, vbCritical
End Sub


' ADMIN AUDIT VIEWER FORM (frmAuditLog) - Query Source

' Create this saved query: qryAuditLog_WithUserNames
' SELECT AuditLog.*, Users.FullName, Users.Username
' FROM AuditLog LEFT JOIN Users ON AuditLog.PerformedBy = Users.UserID
' ORDER BY AuditLog.PerformedDate DESC;

' In frmAuditLog:
' - RecordSource = qryAuditLog_WithUserNames
' - Add filters: Date range, User combo (row source: SELECT UserID, FullName FROM Users WHERE IsActive=True)
' - Add Export to Excel button:
Private Sub cmdExportAudit_Click()
    DoCmd.OutputTo acOutputQuery, "qryAuditLog_WithUserNames", acFormatXLSX, , True
End Sub


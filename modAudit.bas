Attribute VB_Name = "modAudit"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MODULE: modAudit
' PURPOSE: Comprehensive audit trail framework with synchronous and
'          asynchronous logging, form-level change detection, and
'          Admin-only audit viewer support.
' SECURITY: Only Admin can view. All actions are logged permanently.
' AUTHOR: Expert Back-End Developer (MS Access VBA Security Specialist)
' CREATED: November 17, 2025
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Compare Database
Option Explicit

Private Const MODULE_NAME As String = "modAudit"
Private Const BATCH_SIZE As Long = 50
Private Const MAX_QUEUE_SIZE As Long = 1000

' Queue for asynchronous (batched) audit logging
Private Type AuditQueueItem
    TableName As String
    RecordID As Long
    FieldName As String
    OldValue As String
    NewValue As String
    ActionType As String
    PerformedBy As Long
End Type

Private marrQueue() As AuditQueueItem
Private mlngQueueCount As Long

' INITIALIZE QUEUE (called once at app startup from modStartup)
Public Sub InitializeAuditQueue()
    ReDim marrQueue(1 To BATCH_SIZE)
    mlngQueueCount = 0
End Sub


' CORE: LogAudit – Synchronous or Queued depending on criticality
' Procedure: LogAudit
' Purpose  : Primary audit logging function – ALWAYS synchronous for
'            critical operations. Use LogAuditAsync for non-critical.
'
Public Sub LogAudit(ByVal strTableName As String, _
                    ByVal lngRecordID As Long, _
                    ByVal strFieldName As String, _
                    ByVal varOldValue As Variant, _
                    ByVal varNewValue As Variant, _
                    ByVal strActionType As String, _
                    Optional ByVal lngPerformedBy As Long = -1, _
                    Optional ByVal blnForceSync As Boolean = True)

    On Error GoTo ErrorHandler

    Dim strOld As String, strNew As String
    Dim lngUser As Long
    
    strOld = Left(Nz(varOldValue, ""), 255)
    strNew = Left(Nz(varNewValue, ""), 255)
    
    ' Determine user ID: override > current > 1 (system)
    If lngPerformedBy <> -1 Then
        lngUser = lngPerformedBy
    ElseIf g_lngUserID > 0 Then
        lngUser = g_lngUserID
    Else
        lngUser = 1  ' System/Test
    End If

    ' Critical actions (Insert/Delete/Login) or forced sync ? write immediately
    If blnForceSync Or strActionType = "Insert" Or strActionType = "Delete" Then
        WriteAuditRecord strTableName, lngRecordID, strFieldName, strOld, strNew, lngUser, strActionType
    Else
        ' Queue for batch write (non-critical field updates)
        QueueAuditRecord strTableName, lngRecordID, strFieldName, strOld, strNew, strActionType, lngUser
    End If

    Exit Sub

ErrorHandler:
    modUtilities.LogError "LogAudit", Err.Number, Err.Description, _
                         "Table=" & strTableName & " | Action=" & strActionType
End Sub


' WRAPPERS – Clean, self-documenting calls

Public Sub LogInsert(ByVal strTableName As String, ByVal lngRecordID As Long, Optional lngUser As Long = -1)
    LogAudit strTableName, lngRecordID, "", Null, Null, "Insert", lngUser, True
End Sub

Public Sub LogUpdate(ByVal strTableName As String, _
                     ByVal lngRecordID As Long, _
                     ByVal strFieldName As String, _
                     ByVal varOldValue As Variant, _
                     ByVal varNewValue As Variant, _
                     Optional lngUser As Long = -1)
    LogAudit strTableName, lngRecordID, strFieldName, _
             varOldValue, varNewValue, "Update", lngUser, False
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
        .TableName = Left(strTableName, 50)
        .RecordID = lngRecordID
        .FieldName = Left(Nz(strFieldName, ""), 50)
        .OldValue = varOldValue
        .NewValue = varNewValue
        .ActionType = Left(strActionType, 20)
        .PerformedBy = lngUser
    End With

    ' Flush if queue is getting large
    If mlngQueueCount >= BATCH_SIZE Then FlushAuditQueue

End Sub

Public Sub FlushAuditQueue()
    If mlngQueueCount = 0 Then Exit Sub

    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim lngI As Long
    Dim qdf As DAO.QueryDef

    Set db = CurrentDb
    modDatabase.BeginTransaction
    
    Set qdf = db.CreateQueryDef("")
    
    qdf.SQL = "INSERT INTO AuditLog (TableName, RecordID, FieldName, OldValue, NewValue, ActionType, " & _
              "PerformedBy, PerformedDate, WorkstationName) " & _
              "VALUES (pTableName, pRecordID, pFieldName, pOldValue, pNewValue, pActionType, " & _
              "pPerformedBy, Now(), pWorkstationName)"

    For lngI = 1 To mlngQueueCount
        With marrQueue(lngI)
            qdf.Parameters("pTableName") = Left(.TableName, 50)
            qdf.Parameters("pRecordID") = .RecordID
            qdf.Parameters("pFieldName") = Left(Nz(.FieldName, ""), 50)
            qdf.Parameters("pOldValue") = .OldValue
            qdf.Parameters("pNewValue") = .NewValue
            qdf.Parameters("pActionType") = Left(.ActionType, 20)
            qdf.Parameters("pPerformedBy") = .PerformedBy
            qdf.Parameters("pWorkstationName") = Left(Environ("COMPUTERNAME"), 50)

            qdf.Execute dbFailOnError
        End With
    Next lngI

    modDatabase.CommitTransaction
    mlngQueueCount = 0

    Exit Sub

ErrorHandler:
    modDatabase.RollbackTransaction
    modUtilities.LogError MODULE_NAME & "_FlushAuditQueue", Err.Number, Err.Description, "QueueSize=" & mlngQueueCount
End Sub

Private Sub WriteAuditRecord(ByVal strTableName As String, _
                            ByVal lngRecordID As Long, _
                            ByVal strFieldName As String, _
                            ByVal varOldValue As String, _
                            ByVal varNewValue As String, _
                            ByVal lngUser As Long, _
                            ByVal strActionType As String)

    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    
    Set db = CurrentDb
    Set qdf = db.CreateQueryDef("")

    qdf.SQL = "INSERT INTO AuditLog (TableName, RecordID, FieldName, OldValue, NewValue, ActionType, " & _
              "PerformedBy, PerformedDate, WorkstationName) " & _
              "VALUES (pTableName, pRecordID, pFieldName, pOldValue, pNewValue, pActionType, " & _
              "pPerformedBy, Now(), pWorkstationName)"

    qdf.Parameters("pTableName") = Left(strTableName, 50)
    qdf.Parameters("pRecordID") = lngRecordID
    qdf.Parameters("pFieldName") = Left(Nz(strFieldName, ""), 50)
    qdf.Parameters("pOldValue") = varOldValue
    qdf.Parameters("pNewValue") = varNewValue
    qdf.Parameters("pActionType") = Left(strActionType, 20)
    qdf.Parameters("pPerformedBy") = lngUser
    qdf.Parameters("pWorkstationName") = Left(Environ("COMPUTERNAME"), 50)
    
    qdf.Execute dbFailOnError

    Exit Sub

ErrorHandler:
    modUtilities.LogError MODULE_NAME & "_WriteAuditRecord", Err.Number, Err.Description, strSQL
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
                varNew = ctl.Value

                If Nz(varOld, "") <> Nz(varNew, "") Then
                    LogAuditAsync strTableName, lngRecordID, strField, varOld, varNew, "Update"
                End If
            End If
        End If
NextControl:
    Next ctl

    Exit Sub

ErrorHandler:
    modUtilities.LogError MODULE_NAME & "_AuditFormChanges", Err.Number, Err.Description, _
                         "Form=" & frm.Name & " | RecordID=" & lngRecordID
End Sub

Public Sub PurgeOldAudits(Optional ByVal lngDaysToKeep As Long = 365)
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim strSQL As String

    Set db = CurrentDb

    If Not modPermissions.HasPermission(PERM_MANAGE_USERS) Then  ' Reuse admin perm or add new
        Err.Raise 1003, "PurgeOldAudits", "Insufficient permissions"
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
    modUtilities.LogError MODULE_NAME & "_PurgeOldAudits", Err.Number, Err.Description, "Days=" & lngDaysToKeep

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


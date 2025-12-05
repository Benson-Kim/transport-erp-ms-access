Attribute VB_Name = "modServiceNumbers"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MODULE: modServiceNumbers
' PURPOSE: Thread-safe, year-aware service & invoice number generation
'          with full transaction safety and audit trail on failure
' SECURITY: Pessimistic locking + transactions = ZERO duplicates
' AUTHOR: Expert Back-End Developer (MS Access Concurrency Specialist)
' CREATED: November 17, 2025
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Compare Database
Option Explicit

Private Const MODULE_NAME As String = "modServiceNumbers"


' PUBLIC: GenerateServiceNumber

''
''' Function : GenerateServiceNumber
''' Purpose  : Atomically generate next SRV-YYYY-NNNNN number
'''            Handles year rollover and concurrent users safely
''' Returns  : String like "SRV-2025-00001"
''' Throws   : Never – returns "" on critical failure with full audit
''
Public Function GenerateServiceNumber() As String
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim lngYear As Long
    Dim lngNext As Long
    Dim strNumber As String

    Set db = CurrentDb

    ' === BEGIN TRANSACTION + PESSIMISTIC LOCK =========================
    BeginTransaction
    Set rs = db.OpenRecordset("SystemSettings", dbOpenDynaset, dbSeeChanges + dbPessimistic)
    
    If rs.EOF Then
        GoTo MissingSettings
    End If

    With rs
        .MoveFirst
        Do While Not .EOF
            If !SettingName = "CurrentServiceYear" Then
                lngYear = Val(Nz(!SettingValue, Year(Date)))
            ElseIf !SettingName = "NextServiceNumber" Then
                lngNext = Val(Nz(!SettingValue, 1))
            End If
            .MoveNext
        Loop

        ' --- Year rollover detection ---
        If lngYear <> Year(Date) Then
            lngYear = Year(Date)
            lngNext = 1
            
            ' Update year
            .MoveFirst
            Do While Not .EOF
                If !SettingName = "CurrentServiceYear" Then
                    .Edit
                    !SettingValue = CStr(lngYear)
                    !ModifiedBy = g_lngUserID
                    !ModifiedDate = Now()
                    .Update
                ElseIf !SettingName = "NextServiceNumber" Then
                    .Edit
                    !SettingValue = "1"
                    !ModifiedBy = g_lngUserID
                    !ModifiedDate = Now()
                    .Update
                End If
                .MoveNext
            Loop
        End If

        ' --- Build number ---
        strNumber = "SRV-" & Format(lngYear, "0000") & "-" & Format(lngNext, "00000")

        ' --- Increment next number ---
        .MoveFirst
        Do While Not .EOF
            If !SettingName = "NextServiceNumber" Then
                .Edit
                !SettingValue = CStr(lngNext + 1)
                !ModifiedBy = g_lngUserID
                !ModifiedDate = Now()
                .Update
                Exit Do
            End If
            .MoveNext
        Loop

        ' Audit the generation (critical operation ? sync)
        modAudit.LogAudit "SystemSettings", 0, "NextServiceNumber", _
                          lngNext, lngNext + 1, "Increment", , True

        GenerateServiceNumber = strNumber
    End With

    CommitTransaction
    GoTo CleanExit

MissingSettings:
    RollbackTransaction
    modAudit.LogAudit "SystemSettings", 0, "", Null, Null, "MissingSettings", , True
    modUtilities.LogError "GenerateServiceNumber", 1001, _
        "Required SystemSettings records missing (CurrentServiceYear/NextServiceNumber)"
    GenerateServiceNumber = ""

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
    Exit Function

ErrorHandler:
    RollbackTransaction
    modUtilities.LogError "GenerateServiceNumber", Err.Number, Err.Description
    GenerateServiceNumber = ""
    Resume CleanExit
End Function


' PUBLIC: GenerateInvoiceNumber

''
''' Function : GenerateInvoiceNumber
''' Purpose  : Identical logic but for INV-YYYY-NNNNN
''
Public Function GenerateInvoiceNumber() As String
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim lngYear As Long
    Dim lngNext As Long
    Dim strNumber As String

    Set db = CurrentDb

    BeginTransaction
    Set rs = db.OpenRecordset("SystemSettings", dbOpenDynaset, dbSeeChanges + dbPessimistic)
    
    If rs.EOF Then
        GoTo MissingSettings
    End If

    With rs
        .MoveFirst
        Do While Not .EOF
            If !SettingName = "CurrentInvoiceYear" Then
                lngYear = Val(Nz(!SettingValue, Year(Date)))
            ElseIf !SettingName = "NextInvoiceNumber" Then
                lngNext = Val(Nz(!SettingValue, 1))
            End If
            .MoveNext
        Loop

        If lngYear <> Year(Date) Then
            lngYear = Year(Date)
            lngNext = 1

            .MoveFirst
            Do While Not .EOF
                If !SettingName = "CurrentInvoiceYear" Then
                    .Edit
                    !SettingValue = CStr(lngYear)
                    !ModifiedBy = g_lngUserID
                    !ModifiedDate = Now()
                    .Update
                ElseIf !SettingName = "NextInvoiceNumber" Then
                    .Edit
                    !SettingValue = "1"
                    !ModifiedBy = g_lngUserID
                    !ModifiedDate = Now()
                    .Update
                End If
                .MoveNext
            Loop
        End If

        strNumber = "INV-" & Format(lngYear, "0000") & "-" & Format(lngNext, "00000")

        .MoveFirst
        Do While Not .EOF
            If !SettingName = "NextInvoiceNumber" Then
                .Edit
                !SettingValue = CStr(lngNext + 1)
                !ModifiedBy = g_lngUserID
                !ModifiedDate = Now()
                .Update
                Exit Do
            End If
            .MoveNext
        Loop

        modAudit.LogAudit "SystemSettings", 0, "NextInvoiceNumber", _
                          lngNext, lngNext + 1, "Increment", , True

        GenerateInvoiceNumber = strNumber
    End With

    CommitTransaction
    GoTo CleanExit

MissingSettings:
    RollbackTransaction
    modAudit.LogAudit "SystemSettings", 0, "", Null, Null, "MissingSettings", , True
    modUtilities.LogError "GenerateInvoiceNumber", 1002, _
        "Required SystemSettings records missing (CurrentInvoiceYear/NextInvoiceNumber)"
   GenerateInvoiceNumber = ""

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
    Exit Function

ErrorHandler:
    RollbackTransaction
    modUtilities.LogError "GenerateInvoiceNumber", Err.Number, Err.Description
    GenerateInvoiceNumber = ""
    Resume CleanExit
End Function


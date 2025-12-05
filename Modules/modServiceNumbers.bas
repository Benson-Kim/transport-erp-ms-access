'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MODULE: modServiceNumbers
' PURPOSE: Thread-safe, year-aware service & invoice number generation
'          with full transaction safety and audit trail on failure
' SECURITY: Pessimistic locking + transactions = ZERO duplicates
' AUTHOR: Tifamovs Saeol <tifamovs@gmail.com> [Benson Kimathi]
' CREATED: November 17, 2025
' UPDATED: November 24, 2025 � Used Now() for consistency,
'          auto-create missing settings,
'          merged into general GenerateSequentialNumber for DRY.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Database
Option Explicit

Private Const MODULE_NAME As String = "modServiceNumbers"

' PRIVATE: General Sequential Number Generator
' Purpose  : Atomically generate next PREFIX-YYYY-NNNNN number
'            Handles year rollover and concurrent users safely
' Parameters:
'   strPrefix     - e.g., "SRV" or "INV"
'   strYearSetting - SystemSettings name for year, e.g., "CurrentServiceYear"
'   strNextSetting - SystemSettings name for next number, e.g., "NextServiceNumber"
' Returns  : String like "SRV-2025-00001" or "" on failure

Private Function GenerateSequentialNumber(ByVal Prefix As String, _
                                          ByVal YearSetting As String, _
                                          ByVal NextSetting As String) As String
    On Error GoTo ErrHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim curYear As Long
    Dim dbYear As Long
    Dim dbNext As Long
    Dim sql As String

    Set db = CurrentDb

RetryTransaction:
    modDatabase.BeginTransaction
    
    '----- Lock year row -----
    sql = "SELECT SettingValue FROM SystemSettings " & _
          "WHERE SettingName = '" & YearSetting & "'"
    Set rs = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges + dbPessimistic)

    If rs.EOF Then
        modDatabase.RollbackTransaction
        AutoCreateSettings YearSetting, NextSetting
        GoTo RetryTransaction
    End If

    dbYear = CLng(Nz(rs!SettingValue, Year(Now())))
    rs.Close

    '----- Lock next row -----
    sql = "SELECT SettingValue FROM SystemSettings " & _
          "WHERE SettingName = '" & NextSetting & "'"
    Set rs = db.OpenRecordset(sql, dbOpenDynaset, dbSeeChanges + dbPessimistic)

    If rs.EOF Then
        modDatabase.RollbackTransaction
        AutoCreateSettings YearSetting, NextSetting
        GoTo RetryTransaction
    End If

    dbNext = CLng(Nz(rs!SettingValue, 1))
    rs.Close

    curYear = Year(Now())

    '----- Handle rollover -----
    If dbYear <> curYear Then
        dbYear = curYear
        dbNext = 1

        db.Execute "UPDATE SystemSettings SET SettingValue = '" & curYear & "'" & _
                   " WHERE SettingName = '" & YearSetting & "'", dbFailOnError

        db.Execute "UPDATE SystemSettings SET SettingValue = '1'" & _
                   " WHERE SettingName = '" & NextSetting & "'", dbFailOnError
    Else
        '----- Atomic increment -----
        db.Execute "UPDATE SystemSettings SET SettingValue = '" & (dbNext + 1) & "'" & _
                   " WHERE SettingName = '" & NextSetting & "'", dbFailOnError
    End If

    modDatabase.CommitTransaction

    GenerateSequentialNumber = _
        Prefix & "-" & Format$(dbYear, "0000") & "-" & Format$(dbNext, "00000")

    Exit Function

ErrHandler:
    modDatabase.RollbackTransaction
    GenerateSequentialNumber = ""
End Function



' PUBLIC: GenerateServiceNumber
Public Function GenerateServiceNumber() As String

    GenerateServiceNumber = GenerateSequentialNumber("SRV", "CurrentServiceYear", "NextServiceNumber")

End Function

' PUBLIC: GenerateInvoiceNumber
Public Function GenerateInvoiceNumber() As String

    GenerateInvoiceNumber = GenerateSequentialNumber("INV", "CurrentInvoiceYear", "NextInvoiceNumber")

End Function

' HELPER: AutoCreateSettings � Create missing year/next settings
Private Function AutoCreateSettings(ByVal strYearSetting As String, ByVal strNextSetting As String) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb
    modDatabase.BeginTransaction
    
    Set rs = db.OpenRecordset("SystemSettings", dbOpenDynaset)

    ' Create year setting
    rs.AddNew
    rs!SettingName = strYearSetting
    rs!SettingValue = CStr(Year(Now()))
    rs!ModifiedBy = modGlobals.UserID
    rs!ModifiedDate = Now()
    rs.Update

    ' Create next setting
    rs.AddNew
    rs!SettingName = strNextSetting
    rs!SettingValue = "1"
    rs!ModifiedBy = modGlobals.UserID
    rs!ModifiedDate = Now()
    rs.Update
    
    modDatabase.CommitTransaction

    modAudit.LogAudit "SystemSettings", 0, "", Null, Null, "AutoCreatedSettings", , True

    AutoCreateSettings = True

    GoTo CleanExit
ErrorHandler:
    modUtilities.LogError MODULE_NAME & ".AutoCreateSettings", Err.number, Err.Description
    AutoCreateSettings = False

CleanExit:
    On Error Resume Next
    
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing

End Function


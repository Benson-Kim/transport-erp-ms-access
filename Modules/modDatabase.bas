'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MODULE    : modDatabase
' Purpose   : Centralize all database-related helper functions,
'          transaction management, recordset handling and common
'          data-access patterns used throughout the entire application.
' AUTHOR    : Tifamovs Saeol <tifamovs@gmail.com> [Benson Kimathi]
' CREATED   : November 17, 2025
' UPDATED   : December 4, 2025
' NOTES     : This module requires a reference to:
'             - Microsoft DAO 3.6 Object Library (or later)
'             - Microsoft VBScript Regular Expressions 5.5 (for regex functions) 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Database
Option Explicit

' PRIVATE CONSTANTS
Private Const MODULE_NAME As String = "modDatabase"

Private mlngTransCount As Long
Private mdictCache As Object


' PUBLIC FUNCTIONS

' Import result structure
Public Type ImportResult
    TotalRows As Long
    SuccessCount As Long
    ErrorCount As Long
    SkippedCount As Long
    ErrorDetails As Collection
    NewRecordIDs As Collection
End Type

' Function: GetNextID
' Purpose : Returns the next logical AutoNumber value for any table
'            (useful when the ID is needed before the record is saved)
' Parameters:
'    strTableName  - Name of the table
'    strIDField    - Name of the AutoNumber primary key field
' Returns   : Long - Next available ID (Max + 1). Returns 1 if table empty
'
Public Function GetNextID(ByVal strTableName As String, ByVal strIDField As String) As Long
    On Error GoTo ErrorHandler
    
    Dim db          As DAO.Database
    Dim rs          As DAO.Recordset
    Dim lngNextID   As Long
    Dim strSeqName As String
    
    strSeqName = strTableName & "_Sequence"
    Set db = CurrentDb
    
    modDatabase.BeginTransaction
    
    Set rs = db.OpenRecordset("SELECT * FROM SequenceSettings WHERE Name = '" & strSeqName & "'", dbOpenDynaset, dbPessimistic)

    If rs.EOF Then
        rs.AddNew
        rs!Name = strSeqName
        rs!value = 1
        lngNextID = 1
    Else
        rs.Edit
        lngNextID = Nz(rs!value, 0) + 1
        rs!value = lngNextID
    End If
    rs.Update
    
    modDatabase.CommitTransaction
    
    GetNextID = lngNextID
    
CleanExit:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    modDatabase.RollbackTransaction
    modUtilities.LogError MODULE_NAME & ".GetNextID", Err.number, Err.Description
    MsgBox "Error retrieving next ID for " & strTableName & vbCrLf & _
           Err.Description, vbCritical, APP_NAME
    GetNextID = 0
    Resume CleanExit
End Function

' Function: RecordExists
' Purpose : Generic existence check - eliminates repetitive code

Public Function RecordExists(ByVal strTableName As String, _
                            ByVal strCriteria As String) As Boolean
    On Error GoTo ErrorHandler
    
    RecordExists = (CachedDCount("*", strTableName, strCriteria) > 0)
    
CleanExit:
    Exit Function
    
ErrorHandler:
    modUtilities.LogError MODULE_NAME & ".RecordExists", Err.number, Err.Description
    RecordExists = False
    Resume CleanExit
End Function


' Function: GetSingleValue
' Purpose : Return a single field value from the first matching record

Public Function GetSingleValue(ByVal strTableName As String, _
                               ByVal strFieldName As String, _
                               ByVal strCriteria As String, _
                               Optional ByVal varDefault As Variant = Null) As Variant
    On Error GoTo ErrorHandler
    
    Dim varResult As Variant
    varResult = CachedDLookup(strFieldName, strTableName, strCriteria)
    
    If IsNull(varResult) Then
        GetSingleValue = varDefault
    Else
        GetSingleValue = varResult
    End If
    
CleanExit:
    Exit Function
    
ErrorHandler:
    modUtilities.LogError MODULE_NAME & ".GetSingleValue", Err.number, Err.Description
    GetSingleValue = varDefault
    Resume CleanExit
End Function

' Transaction Management - Begin / Commit / Rollback

Public Sub BeginTransaction()
    On Error GoTo ErrorHandler
    
    mlngTransCount = mlngTransCount + 1
    If mlngTransCount = 1 Then DBEngine.BeginTrans
    Exit Sub
    
ErrorHandler:
    modUtilities.LogError MODULE_NAME & ".BeginTransaction", Err.number, Err.Description
    MsgBox "Unable to start transaction.", vbCritical, APP_NAME
End Sub

Public Sub CommitTransaction()
    On Error GoTo ErrorHandler
    
    If mlngTransCount <= 0 Then Err.Raise 1004, "CommitTrans", "No transaction to commit"
    
    If mlngTransCount = 1 Then DBEngine.CommitTrans

    mlngTransCount = mlngTransCount - 1
    
    Exit Sub
    
ErrorHandler:
    modUtilities.LogError MODULE_NAME & ".CommitTransaction", Err.number, Err.Description
    MsgBox "Transaction commit failed.", vbCritical, APP_NAME
End Sub

Public Sub RollbackTransaction()
    On Error GoTo ErrorHandler
    
    If mlngTransCount <= 0 Then Err.Raise 1005, "RollbackTrans", "No transaction to rollback"

    If mlngTransCount = 1 Then DBEngine.Rollback

    mlngTransCount = 0  ' Reset on rollback
    
    Exit Sub
    
ErrorHandler:
    modUtilities.LogError MODULE_NAME & ".RollbackTransaction", Err.number, Err.Description
    MsgBox "Transaction rollback failed.", vbCritical, APP_NAME
End Sub

' Function: ExecuteQuery
' Purpose : Execute an action query (INSERT/UPDATE/DELETE) with
'           centralized error handling and return affected rows

Public Function ExecuteQuery(ByVal strSQL As String) As Long
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Set db = CurrentDb
    db.Execute strSQL, dbFailOnError
    ExecuteQuery = db.RecordsAffected
    
CleanExit:
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    modUtilities.LogError MODULE_NAME & ".ExecuteQuery", Err.number, Err.Description & vbCrLf & "SQL: " & strSQL
    MsgBox "Database error executing query." & vbCrLf & Err.Description, vbCritical, APP_NAME
    ExecuteQuery = -1
    Resume CleanExit
End Function


' Function: GetRecordset
' Purpose : Standardized recordset creation with sensible defaults
'           LockType: dbOpenDynaset (default) - editable
'                     dbOpenSnapshot - read-only (faster for reports)

Public Function GetRecordset(ByVal strSQL As String, _
                  Optional ByVal lngLockType As DAO.LockTypeEnum = dbOpenDynaset) As DAO.Recordset
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL, dbOpenDynaset, 0, lngLockType)
    Set GetRecordset = rs
    
CleanExit:
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    modUtilities.LogError MODULE_NAME & ".GetRecordset", Err.number, Err.Description & vbCrLf & "SQL: " & strSQL
    MsgBox "Failed to open recordset." & vbCrLf & Err.Description, vbCritical, APP_NAME
    Set GetRecordset = Nothing
    Resume CleanExit
End Function

' Function: UpdateQuerySQL
' Purpose: Update or create a saved query with new SQL
' Usage: Used by frmDeletedRecords and other dynamic query forms
Public Function UpdateQuerySQL(QueryName As String, strSQL As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    
    Set db = CurrentDb
    
    ' Try to get existing query
    On Error Resume Next
    Set qdf = db.QueryDefs(QueryName)
    On Error GoTo ErrorHandler
    
    If qdf Is Nothing Then
        Set qdf = db.CreateQueryDef(QueryName, strSQL)
    Else
        qdf.sql = strSQL
    End If
    
    UpdateQuerySQL = True
    Exit Function
    
ErrorHandler:
    modUtilities.LogError MODULE_NAME & ".UpdateQuerySQL", Err.number, Err.Description, _
        "QueryName=" & QueryName
    UpdateQuerySQL = False
End Function

' Function: QueryExists
' Purpose: Check if a saved query exists
Public Function QueryExists(QueryName As String) As Boolean
    On Error Resume Next
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    
    Set db = CurrentDb
    Set qdf = db.QueryDefs(QueryName)
    
    QueryExists = (Not qdf Is Nothing And Err.number = 0)
End Function

' Function: CreateEmptyResultQuery
' Purpose: Create placeholder query for "no results" scenarios
' Usage: Call once during database setup
Public Sub CreateEmptyResultQuery()
    On Error Resume Next
    
    Dim db As DAO.Database
    Set db = CurrentDb
    
    db.QueryDefs.Delete "EmptyResult"
    
    db.CreateQueryDef "EmptyResult", "SELECT 'No records found' AS Message"
End Sub

' Function: DeleteQuery
' Purpose: Safely delete a saved query
Public Function DeleteQuery(QueryName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Set db = CurrentDb
    
    If QueryExists(QueryName) Then
        db.QueryDefs.Delete QueryName
    End If
    
    DeleteQuery = True
    Exit Function
    
ErrorHandler:
    modUtilities.LogError MODULE_NAME & ".DeleteQuery", Err.number, Err.Description, "QueryName=" & QueryName
    DeleteQuery = False
End Function

' Function: CompactAndRepair
' Purpose: Compact and repair database (call during maintenance)
Public Function CompactAndRepair() As Boolean
    On Error GoTo ErrorHandler
    
    Dim strSourceDB As String
    Dim strTempDB As String
    
    strSourceDB = CurrentDb.Name
    strTempDB = CurrentProject.Path & "\TempCompact.accdb"
    
    DoCmd.Close acForm, "", acSaveNo
    
    DBEngine.CompactDatabase strSourceDB, strTempDB
    
    Kill strSourceDB
    Name strTempDB As strSourceDB
   
    Application.Quit acQuitSaveAll
    
    CompactAndRepair = True
    Exit Function
    
ErrorHandler:
    modUtilities.LogError MODULE_NAME & ".CompactAndRepair", Err.number, Err.Description
    CompactAndRepair = False
End Function

' STRING / DATA CLEANING & VALIDATION UTILITIES

' Function: CleanString
' Purpose : Remove leading/trailing spaces, control chars, multiple spaces

Public Function CleanString(ByVal strInput As String) As String
    Dim strTemp As String
    strTemp = Trim$(Nz(strInput, ""))
    strTemp = RegexReplace(strTemp, "[^\x09\x0A\x0D\x20-\x7E]", "")

    strTemp = RegexReplace(strTemp, "\s{2,}", " ")
    CleanString = strTemp
End Function

' Function: FormatVATNumber
' Purpose : Standardize VAT numbers (remove spaces/dots/dashes, uppercase)

Public Function FormatVATNumber(ByVal strVAT As String) As String
    Dim strClean As String
    strClean = UCase(Trim$(Nz(strVAT, "")))
    strClean = Replace(strClean, " ", "")
    strClean = Replace(strClean, "-", "")
    strClean = Replace(strClean, ".", "")
    FormatVATNumber = strClean
End Function


' Function: ParseTime
' Purpose : Convert HH.MM string (e.g. "14.30") to true TimeSerial value
' Returns : Date (time portion only) or #12:00:00 AM# on error

Public Function ParseTime(ByVal strTime As String) As Date
    On Error GoTo ErrorHandler
    
    Dim strClean As String
    strClean = Replace(Trim$(strTime), ".", ":")
    
    If IsTime(strClean) Then
        ParseTime = TimeValue(strClean)
    Else
        ParseTime = #12:00:00 AM#
    End If
    
    Exit Function
    
ErrorHandler:
    ParseTime = #12:00:00 AM#
End Function


' PRIVATE SUPPORT ROUTINES (Regex helpers - require reference to
'  Microsoft VBScript Regular Expressions 5.5)

Private Function RegexReplace(ByVal strInput As String, _
                              ByVal strPattern As String, _
                              ByVal strReplace As String) As String
    With New RegExp
        .Global = True
        .IgnoreCase = True
        .Pattern = strPattern
        RegexReplace = .Replace(strInput, strReplace)
    End With
End Function

Private Function RegexTest(ByVal strInput As String, _
                           ByVal strPattern As String, _
                           ByVal blnIgnoreCase As Boolean) As Boolean
    With New RegExp
        .Pattern = strPattern
        .IgnoreCase = blnIgnoreCase
        RegexTest = .Test(strInput)
    End With
End Function

' Function: BuildSafeInsert - Centralized parameterized INSERT builder
' Purpose : Create and execute safe INSERT with parameters to prevent injection
' Parameters:
'   strTable    - Table name
'   dictFields  - Dictionary of field names (keys) and values
' Returns   : Boolean - True if successful
'
' Usage: Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
'        dict("Field1") = "Value1": ... : BuildSafeInsert "Table", dict
'

Public Function BuildSafeInsert(ByVal strTable As String, ByVal dictFields As Object) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim strFields As String, strParams As String
    Dim vKey As Variant

    Set db = CurrentDb
    Set qdf = db.CreateQueryDef("")

    ' Build SQL
    For Each vKey In dictFields.Keys
        If strFields <> "" Then strFields = strFields & ", "
        If strParams <> "" Then strParams = strParams & ", "
        strFields = strFields & vKey
        strParams = strParams & "p" & vKey
    Next vKey

    qdf.sql = "INSERT INTO " & strTable & " (" & strFields & ") VALUES (" & strParams & ")"

    ' Set parameters
    For Each vKey In dictFields.Keys
        qdf.Parameters("p" & vKey) = dictFields(vKey)
    Next vKey

    qdf.Execute dbFailOnError

    BuildSafeInsert = True

CleanExit:
    Set qdf = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    modUtilities.LogError MODULE_NAME & ".BuildSafeInsert", Err.number, Err.Description, "Table=" & strTable
    BuildSafeInsert = False
    Resume CleanExit
End Function

' CachedDLookup - Wrapper with simple caching

Private Function CachedDLookup(ByVal strField As String, ByVal strTable As String, ByVal strCriteria As String) As Variant

    Dim strKey As String: strKey = "DLookup|" & strField & "|" & strTable & "|" & strCriteria

    If mdictCache Is Nothing Then Set mdictCache = CreateObject("Scripting.Dictionary")

    If mdictCache.Exists(strKey) Then
        CachedDLookup = mdictCache(strKey)
    Else
        CachedDLookup = DLookup(strField, strTable, strCriteria)
        mdictCache(strKey) = CachedDLookup
    End If
End Function

' CachedDCount - Wrapper with simple caching

Public Function CachedDCount(ByVal strField As String, ByVal strTable As String, Optional ByVal strCriteria As String) As Long

    If Len(Nz(strCriteria, "")) = 0 Then strCriteria = "1=1"
    
    Dim strKey As String: strKey = "DCount|" & strField & "|" & strTable & "|" & strCriteria

    If mdictCache Is Nothing Then Set mdictCache = CreateObject("Scripting.Dictionary")

    If mdictCache.Exists(strKey) Then
        CachedDCount = mdictCache(strKey)
    Else
        CachedDCount = DCount(strField, strTable, strCriteria)
        mdictCache(strKey) = CachedDCount
    End If

End Function

' ClearCache - Call when data changes to invalidate cache

Public Sub ClearCache(Optional ByVal strKeyPrefix As String = "")

    If mdictCache Is Nothing Then Exit Sub

    If strKeyPrefix = "" Then
        Set mdictCache = Nothing
    Else
        Dim vKey As Variant
        For Each vKey In mdictCache.Keys
            If Left(vKey, Len(strKeyPrefix)) = strKeyPrefix Then mdictCache.Remove vKey
        Next vKey
    End If

End Sub

' ---------------------------------------------------------------------
' Function: ExportToExcel
' Purpose: Generic export function for any table/query with filters
' RETURNS: True if successful
' ---------------------------------------------------------------------
Public Function ExportToExcel( _
    TableOrQueryName As String, _
    OutputPath As String, _
    Optional ApplyFilter As String = "", _
    Optional OpenAfterExport As Boolean = True) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim strTempQuery As String
    Dim strSQL As String
    
    Set db = CurrentDb
    strTempQuery = "qryExport_Temp_" & Format(Now(), "hhnnss")
    
    ' If filter provided, create filtered query
    If Len(ApplyFilter) > 0 Then
        ' Get base SQL from table/query
        If TableExists(TableOrQueryName) Then
            strSQL = "SELECT * FROM [" & TableOrQueryName & "] WHERE " & ApplyFilter
        Else
            ' It's a query - wrap it
            strSQL = "SELECT * FROM [" & TableOrQueryName & "] WHERE " & ApplyFilter
        End If
        
        ' Create temporary query
        On Error Resume Next
        db.QueryDefs.Delete strTempQuery
        On Error GoTo ErrorHandler
        
        Set qdf = db.CreateQueryDef(strTempQuery, strSQL)
        TableOrQueryName = strTempQuery
    End If
    
    ' Export to Excel
    DoCmd.TransferSpreadsheet _
        TransferType:=acExport, _
        SpreadsheetType:=acSpreadsheetTypeExcel12Xml, _
        tableName:=TableOrQueryName, _
        FileName:=OutputPath, _
        HasFieldNames:=True
    
    ' Clean up temp query
    On Error Resume Next
    db.QueryDefs.Delete strTempQuery
    On Error GoTo ErrorHandler
    
    ' Log export
    modAudit.LogAudit TableOrQueryName, 0, "Export", Null, _
        "Exported to: " & OutputPath, "Export", modGlobals.UserID, True
    
    ' Open file if requested
    If OpenAfterExport Then
        If MsgBox("Export completed successfully!" & vbCrLf & vbCrLf & _
                  "Path: " & OutputPath & vbCrLf & vbCrLf & _
                  "Open the file now?", vbQuestion + vbYesNo, "Export Complete") = vbYes Then
            Application.FollowHyperlink OutputPath
        End If
    End If
    
    ExportToExcel = True
    Exit Function
    
ErrorHandler:
    modUtilities.LogError MODULE_NAME & ".ExportToExcel", Err.number, Err.Description, _
        "Source=" & TableOrQueryName & " | Output=" & OutputPath
    MsgBox "Error exporting data:" & vbCrLf & vbCrLf & Err.Description, vbCritical, APP_NAME
    ExportToExcel = False
End Function



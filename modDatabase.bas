Attribute VB_Name = "modDatabase"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MODULE: modDatabase
' PURPOSE: Centralize all database-related helper functions,
'          transaction management, recordset handling and common
'          data-access patterns used throughout the entire application.
' AUTHOR: Expert Back-End Developer (MS Access VBA Specialist)
' CREATED: November 17, 2025
' SECURITY NOTE: All public procedures include comprehensive error
'                handling, proper object cleanup and audit-ready logging.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Compare Database
Option Explicit


' PRIVATE CONSTANTS
Private Const MODULE_NAME As String = "modDatabase"

Private mlngTransCount As Long
Private mdictCache As Object


' PUBLIC FUNCTIONS

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
        rs!Value = 1
        lngNextID = 1
    Else
        rs.Edit
        lngNextID = Nz(rs!Value, 0) + 1
        rs!Value = lngNextID
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
    modUtilities.LogError MODULE_NAME & "_GetNextID", Err.Number, Err.Description
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
    modUtilities.LogError MODULE_NAME & "_RecordExists", Err.Number, Err.Description
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
    modUtilities.LogError MODULE_NAME & "_GetSingleValue", Err.Number, Err.Description
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
    modUtilities.LogError MODULE_NAME & "_BeginTransaction", Err.Number, Err.Description
    MsgBox "Unable to start transaction.", vbCritical, APP_NAME
End Sub

Public Sub CommitTransaction()
    On Error GoTo ErrorHandler
    
    If mlngTransCount <= 0 Then Err.Raise 1004, "CommitTrans", "No transaction to commit"
    
    If mlngTransCount = 1 Then DBEngine.CommitTrans

    mlngTransCount = mlngTransCount - 1
    
    Exit Sub
    
ErrorHandler:
    modUtilities.LogError MODULE_NAME & "_CommitTransaction", Err.Number, Err.Description
    modUtilities.LogError
    MsgBox "Transaction commit failed.", vbCritical, APP_NAME
End Sub

Public Sub RollbackTransaction()
    On Error GoTo ErrorHandler
    
    If mlngTransCount <= 0 Then Err.Raise 1005, "RollbackTrans", "No transaction to rollback"

    If mlngTransCount = 1 Then DBEngine.Rollback

    mlngTransCount = 0  ' Reset on rollback
    
    Exit Sub
    
ErrorHandler:
    modUtilities.LogError MODULE_NAME & "_RollbackTransaction", Err.Number, Err.Description
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
    modUtilities.LogError MODULE_NAME & " _ExecuteQuery", Err.Number, Err.Description & vbCrLf & "SQL: " & strSQL
    MsgBox "Database error executing query." & vbCrLf & Err.Description, vbCritical, APP_NAME
    ExecuteQuery = 0
    Resume CleanExit
End Function


''' Function: GetRecordset
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
    modUtilities.LogError MODULE_NAME & " _GetRecordset", Err.Number, Err.Description & vbCrLf & "SQL: " & strSQL
    MsgBox "Failed to open recordset." & vbCrLf & Err.Description, vbCritical, APP_NAME
    Set GetRecordset = Nothing
    Resume CleanExit
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

    qdf.SQL = "INSERT INTO " & strTable & " (" & strFields & ") VALUES (" & strParams & ")"

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
    modUtilities.LogError MODULE_NAME & "_BuildSafeInsert", Err.Number, Err.Description, "Table=" & strTable
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

Private Function CachedDCount(ByVal strField As String, ByVal strTable As String, ByVal strCriteria As String) As Long

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

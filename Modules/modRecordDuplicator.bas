
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MODULE    : modRecordDuplicator
' PURPOSE   : Robust record duplication with selective field copying,
'               child record handling, and unique placeholder generation
' AUTHOR    : Tifamovs Saeol <tifamovs@gmail.com> [Benson Kimathi]
' CREATED   : November 18, 2025
' UPDATED   : December 2, 2025 
'    - Improved placeholder uniqueness, fixed exclude logic,
'    - Added preview function and optional fields retrieval.
' NOTES     : 
'    - Handles required and unique fields intelligently
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Database
Option Explicit

Private Const MODULE_NAME As String = "modRecordDuplicator"

' Module-level counter for generating unique placeholders
Private mlngPlaceholderCounter As Long


' PUBLIC: DuplicateRecord � MAIN FUNCTION
Public Function DuplicateRecord( _
    ByVal strTable As String, _
    ByVal lngSourceID As Long, _
    Optional dictInclude As Object = Nothing, _
    Optional dictExclude As Object = Nothing, _
    Optional dictChildren As Object = Nothing, _
    Optional ByVal blnAudit As Boolean = True) As Long
    
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rsSource As DAO.Recordset
    Dim rsTarget As DAO.Recordset
    Dim fld As DAO.Field
    Dim strSQL As String
    Dim lngNewID As Long
    Dim blnUseInclude As Boolean
    Dim colPKFields As Collection
    
    Set db = CurrentDb
    
    modDatabase.BeginTransaction
    
    Set colPKFields = GetPrimaryKeyFields(strTable)
    If colPKFields.Count > 1 Then
        Err.Raise 1006, "DuplicateRecord", "Composite primary keys not yet supported."
    End If

    Dim strPK As String: strPK = colPKFields(1)
    
    ' Open source record
    strSQL = "SELECT * FROM [" & strTable & "] WHERE [" & strPK & "] = " & lngSourceID
    Set rsSource = db.OpenRecordset(strSQL, dbOpenDynaset)

    If rsSource.EOF Then
        MsgBox "Source record not found: " & strTable & " ID=" & lngSourceID, vbExclamation, modGlobals.APP_NAME
        GoTo NotFound
    End If
    
    Set rsTarget = db.OpenRecordset(strTable, dbOpenDynaset)

    If rsTarget Is Nothing Then
        MsgBox "Cannot open target recordset for table: " & strTable, vbCritical
        GoTo CleanExit
    End If
    
    blnUseInclude = Not (dictInclude Is Nothing)
    If blnUseInclude Then
        If dictInclude.Count = 0 Then blnUseInclude = False
    End If

    With rsTarget
        .AddNew
        For Each fld In rsSource.Fields
            ' Skip AutoNumber fields
            If (fld.Attributes And dbAutoIncrField) Then
                ' Skip - will be auto-generated
                
            ' Handle Required + Unique fields (generate placeholder)
            ElseIf IsFieldRequiredAndUnique(strTable, fld.Name) Then
                Select Case UCase(fld.Name)
                    Case "SERVICENUMBER"
                        .Fields(fld.Name) = modServiceNumbers.GenerateServiceNumber()
                    Case "INVOICENUMBER"
                        .Fields(fld.Name) = modServiceNumbers.GenerateInvoiceNumber()
                    Case Else
                        .Fields(fld.Name) = GeneratePlaceholderValue(fld, rsSource(fld.Name), strTable)
                End Select
                
            ' Check if field should be skipped (includes exclude logic)
            ElseIf ShouldSkipField(fld, dictInclude, dictExclude, blnUseInclude, strTable) Then
                ' Skip - don't copy this field
                ' If it's required, Access will use default value or raise error
                
            ' Copy the field value
            Else
                On Error Resume Next
                If fld.Type = dbAttachment Then
                    CopyAttachmentField rsSource, rsTarget, fld.Name
                ElseIf fld.Type = dbLongBinary Then
                    .Fields(fld.Name) = rsSource(fld.Name)
                ElseIf fld.Type = dbMemo Then
                    If Not IsNull(rsSource.Fields(fld.Name).value) Then
                        .Fields(fld.Name).value = CStr(rsSource.Fields(fld.Name).value)
                    End If
                ElseIf Not IsNull(rsSource(fld.Name)) Then
                    .Fields(fld.Name) = rsSource(fld.Name)
                End If
                
                ' Log any errors but continue
                If Err.number <> 0 Then
                    Debug.Print "Warning: Could not copy field " & fld.Name & " - " & Err.Description
                    Err.Clear
                End If
                On Error GoTo ErrorHandler
            End If
        Next fld
        
        ' Standard audit fields
        On Error Resume Next
        .Fields("CreatedBy") = modGlobals.UserID
        .Fields("CreatedDate") = Now()
        .Fields("ModifiedBy") = modGlobals.UserID
        .Fields("ModifiedDate") = Now()
        On Error GoTo ErrorHandler
        
        .Update
        .Bookmark = .LastModified
        
        ' Get new Primary Key
        lngNewID = Nz(.Fields(strPK).value, 0)
        
        If lngNewID = 0 Then
            .Move 0, .LastModified
            lngNewID = Nz(.Fields(strPK).value, 0)
        End If
        
        If lngNewID = 0 Then
            lngNewID = Nz(DMax(strPK, strTable), 1)
        End If
    End With
    
    ' Duplicate child records
    If Not (dictChildren Is Nothing) Then
        Dim vKey As Variant
        For Each vKey In dictChildren.Keys
            Call DuplicateChildRecords(CStr(vKey), lngSourceID, lngNewID, CStr(dictChildren(vKey)))
        Next vKey
    End If
    
    modDatabase.CommitTransaction
    
    If blnAudit Then
        modAudit.LogAudit strTable, lngNewID, "", Null, "Duplicated from ID " & lngSourceID, "Duplicate", , True
    End If
    
    DuplicateRecord = lngNewID
    GoTo CleanExit

NotFound:
    modDatabase.RollbackTransaction
    DuplicateRecord = 0
    GoTo CleanExit

ErrorHandler:
    Debug.Print "ERR " & Err.number & ": " & Err.Description
    modUtilities.LogError MODULE_NAME & ".DuplicateRecord", Err.number, Err.Description
    modDatabase.RollbackTransaction
    DuplicateRecord = 0

CleanExit:
    On Error Resume Next
    If Not rsSource Is Nothing Then rsSource.Close
    If Not rsTarget Is Nothing Then rsTarget.Close
    Set rsSource = Nothing
    Set rsTarget = Nothing
    Set db = Nothing
End Function


' Check if field is both Required AND Unique
Private Function IsFieldRequiredAndUnique(strTable As String, strField As String) As Boolean
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim blnRequired As Boolean
    Dim blnUnique As Boolean
    
    On Error Resume Next
    Set db = CurrentDb
    Set tdf = db.TableDefs(strTable)
    Set fld = tdf.Fields(strField)
    
    If fld Is Nothing Then
        IsFieldRequiredAndUnique = False
        Exit Function
    End If
    
    ' Check if Required
    blnRequired = fld.Required And Not (fld.Attributes And dbAutoIncrField)
    
    ' Check if Unique
    blnUnique = IsFieldUniqueIndexed(strTable, strField)
    
    IsFieldRequiredAndUnique = (blnRequired And blnUnique)
End Function


' FIXED: Generate placeholder value with guaranteed uniqueness
Private Function GeneratePlaceholderValue(fld As DAO.Field, varOriginal As Variant, strTable As String) As Variant
    Dim strPrefix As String
    Dim strTimestamp As String
    Dim strUnique As String
    Dim lngAttempt As Long
    Dim varTestValue As Variant
    Dim blnExists As Boolean
    
    ' Increment module-level counter for uniqueness
    mlngPlaceholderCounter = mlngPlaceholderCounter + 1
    
    ' Generate timestamp with milliseconds + counter
    strTimestamp = Format(Now, "yyyymmddhhnnss") & Format(mlngPlaceholderCounter, "0000")
    
    Select Case fld.Type
        Case dbText, dbMemo
            ' For text fields, append timestamp to original or use COPY prefix
            If Not IsNull(varOriginal) And Len(Nz(varOriginal, "")) > 0 Then
                ' Calculate max prefix length (field size - timestamp - underscore)
                Dim lngMaxPrefix As Long
                lngMaxPrefix = fld.Size - Len(strTimestamp) - 1
                If lngMaxPrefix < 1 Then lngMaxPrefix = 1
                
                strPrefix = Left(CStr(varOriginal), lngMaxPrefix)
                strUnique = strPrefix & "_" & strTimestamp
            Else
                ' Ensure COPY prefix + timestamp fits in field
                If fld.Size < Len("COPY_" & strTimestamp) Then
                    strUnique = Left("COPY_" & strTimestamp, fld.Size)
                Else
                    strUnique = "COPY_" & strTimestamp
                End If
            End If
            
            ' CRITICAL: Verify uniqueness in database
            lngAttempt = 0
            Do
                On Error Resume Next
                blnExists = (DCount("*", strTable, "[" & fld.Name & "]='" & Replace(strUnique, "'", "''") & "'") > 0)
                If Err.number <> 0 Then
                    Debug.Print "GeneratePlaceholderValue: DCount error " & Err.number & " - " & Err.Description
                    blnExists = False ' Assume unique if check fails
                    Err.Clear
                End If
                On Error GoTo 0
                
                If blnExists Then
                    lngAttempt = lngAttempt + 1
                    
                    ' Safety: Break out if too many attempts
                    If lngAttempt >= 100 Then
                        Debug.Print "WARNING: GeneratePlaceholderValue exceeded 100 attempts for " & strTable & "." & fld.Name
                        Exit Do
                    End If
                    
                    mlngPlaceholderCounter = mlngPlaceholderCounter + 1
                    strTimestamp = Format(Now, "yyyymmddhhnnss") & Format(mlngPlaceholderCounter, "0000")
                    
                    If Not IsNull(varOriginal) And Len(Nz(varOriginal, "")) > 0 Then
                        strUnique = Left(CStr(varOriginal), lngMaxPrefix) & "_" & strTimestamp
                    Else
                        strUnique = Left("COPY_" & strTimestamp, fld.Size)
                    End If
                    
                    Debug.Print "GeneratePlaceholderValue: Attempt " & lngAttempt & " - Generated: " & strUnique
                End If
            Loop While blnExists And lngAttempt < 100
            
            GeneratePlaceholderValue = strUnique
            
        Case dbLong, dbInteger
            ' For numeric fields, use timestamp as number + counter
            GeneratePlaceholderValue = CLng(strTimestamp)
            
        Case dbDate
            ' For date fields, use current timestamp
            GeneratePlaceholderValue = Now()
            
        Case Else
            ' Default: use timestamp string
            GeneratePlaceholderValue = "COPY_" & strTimestamp
    End Select
    
End Function


' FIXED: ShouldSkipField - Now properly handles Required + Unique + Exclude
Private Function ShouldSkipField( _
    fld As DAO.Field, _
    dictInclude As Object, _
    dictExclude As Object, _
    blnUseInclude As Boolean, _
    strTable As String) As Boolean
    
    Dim strName As String
    strName = fld.Name
    
    ' Always skip primary key and auto-increment
    If fld.Attributes And dbAutoIncrField Then
        ShouldSkipField = True
        Exit Function
    End If
    
    ' Skip unique indexed fields (unless required - those are handled separately)
    If IsFieldUniqueIndexed(strTable, strName) And Not fld.Required Then
        ShouldSkipField = True
        Exit Function
    End If
    
    ' CRITICAL: Check exclude list BEFORE required check
    ' This allows users to explicitly exclude even required fields
    If Not (dictExclude Is Nothing) And DictionaryExists(dictExclude, strName) Then
        ' User wants to exclude this field
        If fld.Required And Not fld.AllowZeroLength Then
            ' Required field - warn but don't skip (will cause error if no default)
'            Debug.Print "WARNING: Cannot exclude required field '" & strName & "' in table '" & strTable & "'"
            ShouldSkipField = False
        Else
            ' Optional field - safe to skip
            ShouldSkipField = True
        End If
        Exit Function
    End If
    
    ' Include list logic
    If blnUseInclude Then
        ShouldSkipField = Not DictionaryExists(dictInclude, strName)
        Exit Function
    End If
    
    ' Default: don't skip
    ShouldSkipField = False
End Function


' UTILITY FUNCTIONS

Public Function CreateIncludeList(ParamArray fieldNames() As Variant) As Object
    Dim dict As Object
    Dim i As Integer
    
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    
    For i = LBound(fieldNames) To UBound(fieldNames)
        dict(CStr(fieldNames(i))) = True
    Next i
    
    Set CreateIncludeList = dict
End Function

Public Function CreateExcludeList(ParamArray fieldNames() As Variant) As Object
    Dim dict As Object
    Dim i As Integer
    
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    
    For i = LBound(fieldNames) To UBound(fieldNames)
        dict(CStr(fieldNames(i))) = True
    Next i
    
    Set CreateExcludeList = dict
End Function

Public Function CreateChildRelations() As Object
    Set CreateChildRelations = CreateObject("Scripting.Dictionary")
End Function

Public Sub AddChildRelation(dictChildren As Object, _
                           strChildTable As String, _
                           strForeignKeyField As String)
    dictChildren(strChildTable) = strForeignKeyField
End Sub

Public Function GetChildTables(strParentTable As String) As Object
    Dim db As DAO.Database
    Dim rel As DAO.Relation
    Dim dict As Object
    
    Set dict = CreateObject("Scripting.Dictionary")
    Set db = CurrentDb
    
    For Each rel In db.Relations
        If rel.Table = strParentTable Then
            dict(rel.ForeignTable) = rel.Fields(0).ForeignName
        End If
    Next rel
    
    Set GetChildTables = dict
End Function

Private Sub DuplicateChildRecords( _
    strChildTable As String, _
    lngOldParentID As Long, _
    lngNewParentID As Long, _
    strForeignKeyField As String)
    
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rsSource As DAO.Recordset
    Dim rsTarget As DAO.Recordset
    Dim fld As DAO.Field
    Dim strSQL As String
    
    Set db = CurrentDb
    
    strSQL = "SELECT * FROM [" & strChildTable & "] WHERE [" & strForeignKeyField & "] = " & lngOldParentID
    Set rsSource = db.OpenRecordset(strSQL, dbOpenSnapshot)
    
    If rsSource.EOF Then
        rsSource.Close
        Exit Sub
    End If
    
    Set rsTarget = db.OpenRecordset(strChildTable, dbOpenDynaset)
    
    If rsTarget Is Nothing Then
        rsSource.Close
        Exit Sub
    End If
    
    Dim colPKFields As Collection
    Set colPKFields = GetPrimaryKeyFields(strChildTable)
    Dim strChildPK As String
    If colPKFields.Count > 0 Then
        strChildPK = colPKFields(1)
    Else
        strChildPK = ""
    End If
    
    Do While Not rsSource.EOF
        rsTarget.AddNew
        
        For Each fld In rsSource.Fields
            ' Skip foreign key (will be set to new parent ID)
            If fld.Name <> strForeignKeyField Then
                ' Skip AutoNumber fields
                If Not (fld.Attributes And dbAutoIncrField) Then
                    ' Skip if this is the child's primary key
                    If UCase(fld.Name) <> UCase(strChildPK) Then
                        On Error Resume Next
                        
                        Dim blnHandled As Boolean
                        blnHandled = False

                        ' Handle Required + Unique fields
                        If fld.Required And IsFieldUniqueIndexed(strChildTable, fld.Name) Then
                            blnHandled = True
                            Select Case UCase(fld.Name)
                                Case "SERVICENUMBER"
                                    rsTarget(fld.Name) = modServiceNumbers.GenerateServiceNumber()
                                Case "INVOICENUMBER"
                                    rsTarget(fld.Name) = modServiceNumbers.GenerateInvoiceNumber()
                                Case Else
                                    rsTarget(fld.Name) = GeneratePlaceholderValue(fld, rsSource(fld.Name), strChildTable)
                            End Select
                        End If
                        
                        ' Copy regular fields
                        If Not blnHandled Then
                            If fld.Type = dbAttachment Then
                                CopyAttachmentField rsSource, rsTarget, fld.Name
                            ElseIf fld.Type = dbLongBinary Then
                                If Not IsNull(rsSource(fld.Name)) Then
                                    rsTarget(fld.Name) = rsSource(fld.Name)
                                End If
                            ElseIf fld.Type = dbMemo Then
                                If Not IsNull(rsSource.Fields(fld.Name).value) Then
                                    rsTarget.Fields(fld.Name).value = CStr(rsSource.Fields(fld.Name).value)
                                End If
                            ElseIf Not IsNull(rsSource(fld.Name)) Then
                                rsTarget(fld.Name) = rsSource(fld.Name)
                            End If
                        End If
                        
                        If Err.number <> 0 Then
                            Debug.Print "Child field copy warning (" & strChildTable & "." & fld.Name & "): " & Err.Description
                            Err.Clear
                        End If
                        On Error GoTo ErrorHandler
                    End If
                End If
            End If
        Next fld
        
        On Error Resume Next
        rsTarget(strForeignKeyField) = lngNewParentID
        If Err.number <> 0 Then
            rsTarget.CancelUpdate
            Err.Clear
            rsSource.MoveNext
            On Error GoTo ErrorHandler
            GoTo NextRecord
        End If
        On Error GoTo ErrorHandler
        
        rsTarget.Update
        
NextRecord:
        rsSource.MoveNext
    Loop
    
CleanExit:
    On Error Resume Next
    If Not rsSource Is Nothing Then rsSource.Close
    If Not rsTarget Is Nothing Then rsTarget.Close
    Set rsSource = Nothing
    Set rsTarget = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    Debug.Print "DuplicateChildRecords ERROR: " & Err.number & " - " & Err.Description & " | Table: " & strChildTable
    modUtilities.LogError MODULE_NAME & ".DuplicateChildRecords", Err.number, Err.Description
    Resume CleanExit
End Sub

Private Function IsFieldUniqueIndexed(strTable As String, strField As String) As Boolean
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    
    On Error Resume Next
    Set db = CurrentDb
    Set tdf = db.TableDefs(strTable)
    
    For Each idx In tdf.Indexes
        If idx.Unique Then
            For Each fld In idx.Fields
                If UCase(fld.Name) = UCase(strField) Then
                    IsFieldUniqueIndexed = True
                    Exit Function
                End If
            Next fld
        End If
    Next idx
    
    IsFieldUniqueIndexed = False
End Function

Public Function GetPrimaryKeyField(strTable As String) As String
    Dim dbFE As DAO.Database
    Dim dbBE As DAO.Database
    Dim tdfFE As DAO.TableDef
    Dim tdfBE As DAO.TableDef
    Dim idx As DAO.Index
    Dim backEndPath As String

    Set dbFE = CurrentDb
    Set tdfFE = dbFE.TableDefs(strTable)

    If Len(tdfFE.Connect) = 0 Then
        For Each idx In tdfFE.Indexes
            If idx.Primary Then
                GetPrimaryKeyField = idx.Fields(0).Name
                Exit Function
            End If
        Next
        GetPrimaryKeyField = ""
        Exit Function
    End If

    backEndPath = Mid(tdfFE.Connect, InStr(tdfFE.Connect, "DATABASE=") + 9)
    Set dbBE = DBEngine.Workspaces(0).OpenDatabase(backEndPath)
    Set tdfBE = dbBE.TableDefs(tdfFE.SourceTableName)

    For Each idx In tdfBE.Indexes
        If idx.Primary Then
            GetPrimaryKeyField = idx.Fields(0).Name
            dbBE.Close
            Exit Function
        End If
    Next

    dbBE.Close
    GetPrimaryKeyField = ""
End Function

Public Function GetPrimaryKeyFields(strTable As String) As Collection
    Dim dbFE As DAO.Database
    Dim dbBE As DAO.Database
    Dim tdfFE As DAO.TableDef
    Dim tdfBE As DAO.TableDef
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    Dim col As New Collection
    Dim backEndPath As String

    Set dbFE = CurrentDb
    Set tdfFE = dbFE.TableDefs(strTable)

    If Len(tdfFE.Connect) = 0 Then
        For Each idx In tdfFE.Indexes
            If idx.Primary Then
                For Each fld In idx.Fields
                    col.Add fld.Name
                Next fld
                Exit For
            End If
        Next idx
        
        If col.Count = 0 Then col.Add "ID"
        Set GetPrimaryKeyFields = col
        Exit Function
    End If

    backEndPath = Mid(tdfFE.Connect, InStr(tdfFE.Connect, "DATABASE=") + 9)
    Set dbBE = DBEngine.Workspaces(0).OpenDatabase(backEndPath)
    Set tdfBE = dbBE.TableDefs(tdfFE.SourceTableName)

    For Each idx In tdfBE.Indexes
        If idx.Primary Then
            For Each fld In idx.Fields
                col.Add fld.Name
            Next fld
            Exit For
        End If
    Next idx

    dbBE.Close

    If col.Count = 0 Then col.Add "ID"
    Set GetPrimaryKeyFields = col
End Function

Private Function DictionaryExists(dict As Object, key As Variant) As Boolean
    On Error Resume Next
    Dim temp As Variant
    temp = dict(key)
    DictionaryExists = (Err.number = 0)
    On Error GoTo 0
End Function

Private Sub CopyAttachmentField(rsSource As DAO.Recordset, rsTarget As DAO.Recordset, strFieldName As String)
    Dim rsSrcAttach As DAO.Recordset2
    Dim rsTgtAttach As DAO.Recordset2

    Set rsSrcAttach = rsSource.Fields(strFieldName).value
    Set rsTgtAttach = rsTarget.Fields(strFieldName).value

    Do While Not rsSrcAttach.EOF
        rsTgtAttach.AddNew
        rsTgtAttach!FileData = rsSrcAttach!FileData
        rsTgtAttach!FileName = rsSrcAttach!FileName
        rsTgtAttach!FileType = rsSrcAttach!FileType
        rsTgtAttach.Update
        rsSrcAttach.MoveNext
    Loop

    rsSrcAttach.Close
    rsTgtAttach.Close
End Sub

' UPDATED: SafeDuplicate with better error reporting
Public Function SafeDuplicate( _
    ByVal strTable As String, _
    ByVal varKey As Variant, _
    ByRef strError As String, _
    Optional dictInclude As Object, _
    Optional dictExclude As Object, _
    Optional dictChildren As Object, _
    Optional blnAudit As Boolean = True) As Variant

    Dim vResult As Variant
    strError = ""

    On Error Resume Next
    vResult = DuplicateRecord(strTable, varKey, dictInclude, dictExclude, dictChildren, blnAudit)
    
    If Err.number <> 0 Then
        strError = "Error " & Err.number & ": " & Err.Description & _
                   " [Table=" & strTable & ", ID=" & varKey & "]"
        modUtilities.LogError "SafeDuplicate", Err.number, strError
        Err.Clear
        SafeDuplicate = Null
        Exit Function
    End If
    On Error GoTo 0

    If IsNull(vResult) Or vResult = 0 Then
        strError = "Duplication returned 0 (source not found or constraint violation)"
        SafeDuplicate = Null
    Else
        SafeDuplicate = vResult
    End If
End Function

Public Function DuplicateMultipleRecords( _
    strTable As String, _
    arrSourceIDs() As Long, _
    Optional dictInclude As Object = Nothing, _
    Optional dictExclude As Object = Nothing, _
    Optional dictChildren As Object = Nothing, _
    Optional ByVal blnAudit As Boolean = True) As Collection
    
    Dim col As Collection
    Dim i As Long
    Dim lngNewID As Long
    Dim strError As String
    
    Set col = New Collection
    
    ' Reset counter at start of batch operation
    mlngPlaceholderCounter = 0
    
    For i = LBound(arrSourceIDs) To UBound(arrSourceIDs)
        strError = ""
        lngNewID = SafeDuplicate(strTable, arrSourceIDs(i), strError, dictInclude, dictExclude, dictChildren, blnAudit)
        
        If lngNewID > 0 Then
            col.Add lngNewID
        Else
            ' Log which record failed
            Debug.Print "DuplicateMultipleRecords: Failed to duplicate ID " & arrSourceIDs(i) & " - " & strError
        End If
    Next i
    
    Set DuplicateMultipleRecords = col
End Function

Public Function PreviewDuplication( _
    strTable As String, _
    Optional dictInclude As Object = Nothing, _
    Optional dictExclude As Object = Nothing) As String
    
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim strResult As String
    Dim blnUseInclude As Boolean
    
    Set db = CurrentDb
    Set tdf = db.TableDefs(strTable)
    
    blnUseInclude = Not (dictInclude Is Nothing)
    If blnUseInclude And dictInclude.Count = 0 Then blnUseInclude = False
    
    strResult = "Fields to be duplicated for [" & strTable & "]:" & vbCrLf & vbCrLf
    
    For Each fld In tdf.Fields
        If Not ShouldSkipField(fld, dictInclude, dictExclude, blnUseInclude, strTable) Then
            strResult = strResult & "  ? " & fld.Name & vbCrLf
        Else
            strResult = strResult & "  ? " & fld.Name & " (skipped)" & vbCrLf
        End If
    Next fld
    
    PreviewDuplication = strResult
End Function

Public Function GetDuplicatableFields(strTable As String) As Collection
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim col As Collection
    
    Set col = New Collection
    Set db = CurrentDb
    Set tdf = db.TableDefs(strTable)
    
    For Each fld In tdf.Fields
        If Not (fld.Attributes And dbAutoIncrField) Then
            If Not (IsFieldUniqueIndexed(strTable, fld.Name) And Not fld.Required) Then
                col.Add fld.Name
            End If
        End If
    Next fld
    
    Set GetDuplicatableFields = col
End Function

' NEW: Get list of optional (nullable) fields that can be safely excluded
Public Function GetOptionalFields(strTable As String) As Collection
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim col As Collection
    
    Set col = New Collection
    Set db = CurrentDb
    Set tdf = db.TableDefs(strTable)
    
    For Each fld In tdf.Fields
        ' Include only optional fields (not required, not AutoNumber)
        If Not (fld.Attributes And dbAutoIncrField) Then
            If Not fld.Required Then
                col.Add fld.Name
            End If
        End If
    Next fld
    
    Set GetOptionalFields = col
End Function

' NEW: Get list of required fields in a table
Public Function GetRequiredFields(strTable As String) As Collection
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim col As Collection
    
    Set col = New Collection
    Set db = CurrentDb
    Set tdf = db.TableDefs(strTable)
    
    For Each fld In tdf.Fields
        If fld.Required And Not (fld.Attributes And dbAutoIncrField) Then
            col.Add fld.Name
        End If
    Next fld
    
    Set GetRequiredFields = col
End Function


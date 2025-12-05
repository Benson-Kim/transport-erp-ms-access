Attribute VB_Name = "modRecordDuplicator"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MODULE: modRecordDuplicator
' PURPOSE: Universal record duplication with field control & child records
' FEATURES:
'   • Duplicate any record (single table or with children)
'   • Include/Exclude specific fields
'   • Auto-skip unique/indexed fields (PK, VAT, ServiceNumber, etc.)
'   • Duplicate related child records (1-to-many)
'   • Full audit trail with transactions
'   • Helper utilities for building field lists
' AUTHOR: Expert Back-End Developer
' UPDATED: November 19, 2025
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Compare Database
Option Explicit

Private Const MODULE_NAME As String = "modRecordDuplicator"


' PUBLIC: DuplicateRecord – MAIN FUNCTION

Public Function DuplicateRecord( _
    ByVal strTable As String, _
    ByVal lngSourceID As Long, _
    Optional dictInclude As Object = Nothing, _
    Optional dictExclude As Object = Nothing, _
    Optional dictChildren As Object = Nothing) As Long
    
    On Error GoTo ErrorHandler
    
    Dim ws As Workspace
    Dim db As DAO.Database
    Dim rsSource As DAO.Recordset
    Dim rsTarget As DAO.Recordset
    Dim fld As DAO.Field
    Dim strSQL As String
    Dim lngNewID As Long
    Dim blnUseInclude As Boolean
    Dim strPK As String
    Dim blnIsTrans As Boolean
    
    Set db = CurrentDb
    ws.BeginTrans
    blnIsTrans = True
    
    ' Get primary key field
    strPK = GetPrimaryKeyField(strTable)
    
    ' Open source record
    strSQL = "SELECT * FROM [" & strTable & "] WHERE [" & strPK & "] = " & lngSourceID
    Set rsSource = db.OpenRecordset(strSQL, dbOpenSnapshot)
    If rsSource.EOF Then
        Debug.Print "Source record not found: " & strTable & " ID=" & lngSourceID
        GoTo NotFound
    End If
    
    Set rsTarget = db.OpenRecordset("[" & strTable & "]", dbOpenDynaset)
    
    blnUseInclude = Not (dictInclude Is Nothing)
    If blnUseInclude And dictInclude.Count = 0 Then blnUseInclude = False
    
    With rsTarget
        .AddNew
        For Each fld In rsSource.Fields
            If Not ShouldSkipField(fld, dictInclude, dictExclude, blnUseInclude, strTable) Then
                If Not IsNull(rsSource(fld.Name)) Then
                    .Fields(fld.Name) = rsSource(fld.Name)
                End If
            End If
        Next fld
        
        ' Standard audit fields
        On Error Resume Next
        .Fields("CreatedBy") = Environ("USERNAME")
        .Fields("CreatedDate") = Now()
        .Fields("ModifiedBy") = Environ("USERNAME")
        .Fields("ModifiedDate") = Now()
        On Error GoTo ErrorHandler
        
        .Update
        .Bookmark = .LastModified
        
        ' Get new Primary Key
        If Not rsTarget.Fields(strPK) Is Nothing Then
            lngNewID = Nz(rsTarget.Fields(strPK).Value, 0)
        End If
        
        If lngNewID = 0 Then
            rsTarget.Move 0, rsTarget.LastModified
            lngNewID = Nz(rsTarget.Fields(strPK).Value, 0)
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
    
    ws.CommitTrans
    blnIsTrans = False
    DuplicateRecord = lngNewID
    GoTo CleanExit

NotFound:
    
    If blnIsTrans Then ws.Rollback
    blnIsTrans = False
    DuplicateRecord = 0
    GoTo CleanExit

ErrorHandler:
    If blnIsTrans Then ws.Rollback
    blnIsTrans = False
    DuplicateRecord = 0

CleanExit:
    On Error Resume Next
    If Not rsSource Is Nothing Then rsSource.Close
    If Not rsTarget Is Nothing Then rsTarget.Close
    Set rsSource = Nothing
    Set rsTarget = Nothing
    If blnIsTrans = False Then Set ws = Nothing
    Set db = Nothing
End Function


' UTILITY: Create Include Dictionary

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


' UTILITY: Create Exclude Dictionary

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


' UTILITY: Create Child Relationships Dictionary

Public Function CreateChildRelations() As Object
    Set CreateChildRelations = CreateObject("Scripting.Dictionary")
End Function

Public Sub AddChildRelation(dictChildren As Object, _
                           strChildTable As String, _
                           strForeignKeyField As String)
    dictChildren(strChildTable) = strForeignKeyField
End Sub


' UTILITY: Get All Field Names (excluding auto/unique)

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
            If Not IsFieldUniqueIndexed(strTable, fld.Name) Then
                col.Add fld.Name
            End If
        End If
    Next fld
    
    Set GetDuplicatableFields = col
End Function


' UTILITY: Get Child Tables for a Parent Table

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


' UTILITY: Duplicate Multiple Records

Public Function DuplicateMultipleRecords( _
    strTable As String, _
    arrSourceIDs() As Long, _
    Optional dictInclude As Object = Nothing, _
    Optional dictExclude As Object = Nothing, _
    Optional dictChildren As Object = Nothing) As Collection
    
    Dim col As Collection
    Dim i As Long
    Dim lngNewID As Long
    
    Set col = New Collection
    
    For i = LBound(arrSourceIDs) To UBound(arrSourceIDs)
        lngNewID = DuplicateRecord(strTable, arrSourceIDs(i), dictInclude, dictExclude, dictChildren)
        If lngNewID > 0 Then
            col.Add lngNewID
        End If
    Next i
    
    Set DuplicateMultipleRecords = col
End Function


' UTILITY: Preview Fields That Will Be Duplicated

Public Function PreviewDuplication(strTable As String, _
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


' PRIVATE: Should we skip this field?

Private Function ShouldSkipField(fld As DAO.Field, _
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
    
    ' Skip known unique indexed fields
    If IsFieldUniqueIndexed(strTable, strName) Then
        ShouldSkipField = True
        Exit Function
    End If
    
    ' Include/Exclude logic
    If blnUseInclude Then
        ShouldSkipField = Not DictionaryExists(dictInclude, strName)
    Else
        If Not (dictExclude Is Nothing) Then
            ShouldSkipField = DictionaryExists(dictExclude, strName)
        Else
            ShouldSkipField = False
        End If
    End If
End Function


' PRIVATE: Duplicate child records (1-to-many)

Private Sub DuplicateChildRecords(strChildTable As String, _
                                 lngOldParentID As Long, _
                                 lngNewParentID As Long, _
                                 strForeignKeyField As String)
    
    Dim db As DAO.Database
    Dim rsSource As DAO.Recordset
    Dim rsTarget As DAO.Recordset
    Dim fld As DAO.Field
    Dim strSQL As String
    
    Set db = CurrentDb
    strSQL = "SELECT * FROM [" & strChildTable & "] WHERE [" & strForeignKeyField & "] = " & lngOldParentID
    Set rsSource = db.OpenRecordset(strSQL, dbOpenSnapshot)
    If rsSource.EOF Then Exit Sub
    
    Set rsTarget = db.OpenRecordset("[" & strChildTable & "]", dbOpenDynaset)
    
    Do While Not rsSource.EOF
        rsTarget.AddNew
        For Each fld In rsSource.Fields
            If fld.Name <> strForeignKeyField And Not (fld.Attributes And dbAutoIncrField) Then
                If Not IsNull(rsSource(fld.Name)) Then
                    rsTarget(fld.Name) = rsSource(fld.Name)
                End If
            End If
        Next fld
        rsTarget(strForeignKeyField) = lngNewParentID
        rsTarget.Update
        rsSource.MoveNext
    Loop
    
    rsSource.Close
    rsTarget.Close
End Sub


' HELPER: Check if field has unique index

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


' HELPER: Get primary key field name

Private Function GetPrimaryKeyField(strTable As String) As String
    Dim tdf As DAO.TableDef
    Dim idx As DAO.Index
    
    Set tdf = CurrentDb.TableDefs(strTable)
    For Each idx In tdf.Indexes
        If idx.Primary Then
            GetPrimaryKeyField = idx.Fields(0).Name
            Exit Function
        End If
    Next idx
    
    GetPrimaryKeyField = "ID"
End Function


' HELPER: Safe dictionary exists check

Private Function DictionaryExists(dict As Object, key As Variant) As Boolean
    On Error Resume Next
    Dim temp As Variant
    temp = dict(key)
    DictionaryExists = (Err.Number = 0)
    On Error GoTo 0
End Function


' USAGE EXAMPLES

Sub Example_SimpleDuplicate()
    ' Duplicate entire record (auto-skips unique fields)
    Dim lngNewID As Long
    Debug.Print "Testing simple duplicate of Clients ID=65..."
    lngNewID = DuplicateRecord("Clients", 65)
    If lngNewID > 0 Then
        Debug.Print "SUCCESS! New Client ID: " & lngNewID
    Else
        Debug.Print "FAILED - returned 0"
    End If
End Sub

Sub Example_ExcludeFields()
    ' Duplicate but exclude specific fields
    Dim dictExclude As Object
    Set dictExclude = CreateExcludeList("EmailBilling", "Notes", "PhoneMain")
    
    Dim lngNewID As Long
    lngNewID = DuplicateRecord("Clients", 65, , dictExclude)
    Debug.Print "New Client ID: " & lngNewID
End Sub

Sub Example_IncludeOnlySpecific()
    ' Copy ONLY specific fields
    Dim dictInclude As Object
    Set dictInclude = CreateIncludeList("ClientName", "Address", "City", "PostalCode")
    
    Dim lngNewID As Long
    lngNewID = DuplicateRecord("Clients", 65, dictInclude)
    Debug.Print "New Client ID: " & lngNewID
End Sub

Sub Example_WithChildren()
    ' Duplicate client with all services and contacts
    Dim dictChildren As Object
    Set dictChildren = CreateChildRelations()
    
    Call AddChildRelation(dictChildren, "Services", "ClientID")
    Call AddChildRelation(dictChildren, "ClientContacts", "ClientID")
    
    Dim lngNewID As Long
    lngNewID = DuplicateRecord("Clients", 65, , , dictChildren)
    Debug.Print "New Client ID: " & lngNewID
End Sub

Sub Example_PreviewDuplication()
    ' Preview what will be duplicated
    Dim dictExclude As Object
    Set dictExclude = CreateExcludeList("Notes", "EmailBilling")
    
    Debug.Print PreviewDuplication("Clients", , dictExclude)
End Sub

Sub Example_GetDuplicatableFields()
    ' Get list of all fields that CAN be duplicated
    Dim col As Collection
    Dim v As Variant
    
    Set col = GetDuplicatableFields("Clients")
    For Each v In col
        Debug.Print v
    Next v
End Sub

Sub Example_AutoDetectChildren()
    ' Automatically detect child tables
    Dim dictChildren As Object
    Set dictChildren = GetChildTables("Clients")
    
    Dim lngNewID As Long
    lngNewID = DuplicateRecord("Clients", 65, , , dictChildren)
    Debug.Print "New Client ID with all children: " & lngNewID
End Sub

Sub Example_DuplicateMultiple()
    ' Duplicate multiple records at once
    Dim arrIDs(1 To 3) As Long
    arrIDs(1) = 65
    arrIDs(2) = 72
    arrIDs(3) = 88
    
    Dim colNewIDs As Collection
    Set colNewIDs = DuplicateMultipleRecords("Clients", arrIDs)
    
    Dim v As Variant
    For Each v In colNewIDs
        Debug.Print "Created Client ID: " & v
    Next v
End Sub



Option Compare Database
Option Explicit

Private blnSyncing As Boolean

Private Const BASE_SQL As String = _
    "SELECT SupplierID, SupplierName, VATNumber, Telephone, Email, IsDeleted, Country, TypeOfServices " & _
    "FROM Suppliers WHERE IsDeleted = False"


Private Sub Form_Current()
    If blnSyncing Then Exit Sub
    If Me.NewRecord Then Exit Sub
    If Me.Parent.dirty Then Exit Sub
    
    On Error Resume Next
    Me.Parent.SyncSupplier Me!SupplierID
    On Error GoTo 0
 
End Sub

Private Sub Form_Load()
    ' Set initial RecordSource
    Me.RecordSource = BASE_SQL
    Me.Requery
End Sub

'Public Sub SyncFromParent(lngSupplierID As Long)
'    blnSyncing = True
'
'    On Error Resume Next
'    Me.Recordset.FindFirst "SupplierID = " & CLng(lngSupplierID)
'    On Error GoTo 0
'
'    If Not Me.NewRecord Then
'        Me.SelTop = Me.CurrentRecord
'    End If
'
'    blnSyncing = False
'End Sub

Public Sub SyncFromParent(lngSupplierID As Long)
    blnSyncing = True
    
    If Me.Recordset.RecordCount > 0 Then
        Me.Recordset.FindFirst "SupplierID = " & CLng(lngSupplierID)
        If Not Me.Recordset.NoMatch Then Me.Bookmark = Me.Recordset.Bookmark
    End If
    
    blnSyncing = False
End Sub


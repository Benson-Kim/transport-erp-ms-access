VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_subList_Suppliers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private blnSyncing As Boolean
Private Const BASE_SQL As String = _
    "SELECT SupplierID, SupplierName, VATNumber, Telephone, Email, IsDeleted " & _
    "FROM Suppliers WHERE IsDeleted = False"

Private Sub Form_Current()
    If blnSyncing Then Exit Sub

    If Not Me.NewRecord Then
        Me.Parent.SyncSupplier Me!SupplierID
    End If
End Sub

Private Sub Form_Load()
    ' Set initial RecordSource
    Me.RecordSource = BASE_SQL
End Sub

Private Sub txtSearch_Change()
    Me.TimerInterval = 600  ' Restart timer
End Sub

Private Sub Form_Timer()
    Me.TimerInterval = 0
    ApplySearchFilter
End Sub

Public Sub SyncFromParent(lngSupplierID As Long)
    blnSyncing = True
    
    Me.Recordset.FindFirst "SupplierID = " & lngSupplierID
    
    If Not Me.NewRecord Then
        Me.SelTop = Me.CurrentRecord
    End If
    
    blnSyncing = False
End Sub

Private Sub ApplySearchFilter()
    Dim strSearchTerm As String
    Dim strSQL As String

    strSearchTerm = Trim$(Nz(Me.txtSearch, ""))

    ' If empty ? restore all rows
    If Len(strSearchTerm) = 0 Then
        Me.RecordSource = BASE_SQL
        DoEvents
        Me.Requery
        If Not Me.Parent.NewRecord Then Call SyncFromParent(Me.Parent!SupplierID)
        Exit Sub
    End If

    ' Filter SQL (search multiple fields)
    strSQL = BASE_SQL & _
        " AND (" & _
        "SupplierName LIKE '*" & strSearchTerm & "*' OR " & _
        "VATNumber LIKE '*" & strSearchTerm & "*' OR " & _
        "Telephone LIKE '*" & strSearchTerm & "*' OR " & _
        "Email LIKE '*" & strSearchTerm & "*'" & _
        ")"

    Me.RecordSource = strSQL
    Me.Requery

    ' Resync active selection
    If Not Me.Parent.NewRecord Then Call SyncFromParent(Me.Parent!SupplierID)
End Sub



Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Esc = 27
    If KeyCode = 27 Then
        Me.txtSearch = ""
        ApplySearchFilter
        KeyCode = 0
    End If

    ' Ctrl+F = 6 (Shift: Ctrl=2)
    If KeyCode = vbKeyF And (Shift And acCtrlMask) <> 0 Then
        Me.txtSearch.SetFocus
        KeyCode = 0
    End If
End Sub




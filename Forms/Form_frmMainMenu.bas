
Option Compare Database
Option Explicit

Private Sub cmdSuppliers_Click()
    Me.frmtab.value = 2
    frmtab_Change
End Sub

Private Sub Form_Load()
    
    InitializeGlobalVariables
    
    Me.txtUser.value = IIf(Len(g_strFullName) > 0, g_strFullName, "____________")
    Me.txtRole.value = IIf(Len(g_strUserRole) > 0, g_strUserRole, "____________")
End Sub

Private Sub frmtab_Change()

    Select Case Me.frmtab.value
    
        Case 0
            Me.SubfrmContainer.SourceObject = "frmServices"
        Case 1
            Me.SubfrmContainer.SourceObject = "frmClients"
        Case 2
            Me.SubfrmContainer.SourceObject = "frmSuppliers"
        Case 3
            Me.SubfrmContainer.SourceObject = "frmInvoices"
        Case 4
            Me.SubfrmContainer.SourceObject = "frmReports"
        Case 5
            Me.SubfrmContainer.SourceObject = "frmSettings"
        Case Else
            Me.SubfrmContainer.SourceObject = "frmServices"
        
        Me.Requery
    End Select
    
End Sub


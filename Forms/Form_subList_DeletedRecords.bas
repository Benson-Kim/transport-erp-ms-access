

Option Compare Database
Option Explicit

Private Sub Form_Current()
    If Me.Parent.dirty Then Exit Sub
    
    On Error Resume Next
    If Not Me.Parent Is Nothing Then
        Me.Parent.SubformRowChanged
    End If
    On Error GoTo 0
 
End Sub


Option Compare Database
Option Explicit
Public Function ScrollTestWindow(ByVal viewName As String) As LongPtr
    With Forms("Reports")!subReports
        If viewName <> "Initial" Then
            .Width = 12000
            Forms("Reports").cmdGridView_Click
        End If
        ScrollTestWindow = .Form.hWnd
    End With
End Function

Option Compare Database
Option Explicit
Private Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal window As LongPtr, ByVal message As Long, ByVal key As LongPtr, ByVal data As LongPtr) As LongPtr

Public Function PrepareCancelTest(ByVal openArgument As String) As LongPtr
    Dim entry As Object, index As Long
    Const prefix As String = "103"
    DoCmd.OpenForm "Report Entry", , , , , , openArgument
    Set entry = Forms("Report Entry")
    entry!txtLocation.SetFocus
    entry!txtLocation.SelStart = 0
    entry!txtLocation.SelLength = Len(entry!txtLocation.Text)
    For index = 1 To Len(prefix)
        SendMessage GetFocus(), &H102, AscW(Mid$(prefix, index, 1)), 1
    Next
    If Not entry!lstLocationMatches.Visible Then Err.Raise 520, , "Suggestions did not open."
    PrepareCancelTest = entry.hWnd
End Function

Public Function TryCancelEntry(ByVal closeFromWindow As Boolean) As Boolean
    Dim entry As Object
    Set entry = Forms("Report Entry")
    entry!cmdCancel.SetFocus
    If closeFromWindow Then
        DoCmd.Close acForm, "Report Entry"
    Else
        entry.cmdCancel_Click
    End If
    TryCancelEntry = Not CurrentProject.AllForms("Report Entry").IsLoaded
End Function

Public Function CancelTestPreservedDraft() As Boolean
    With Forms("Report Entry")
        CancelTestPreservedDraft = (.txtLocation.Value = "103" And Not .lstLocationMatches.Visible)
    End With
End Function

Public Function EntryEditableTestID() As String
    EntryEditableTestID = Nz(DMax("ReportID", "DailyReports", "created_at >= DateAdd('h', -24, Now())"), "")
End Function

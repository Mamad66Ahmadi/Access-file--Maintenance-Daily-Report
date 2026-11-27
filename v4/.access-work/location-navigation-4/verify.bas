Option Compare Database
Option Explicit

Private Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
Private Declare PtrSafe Function SendMessageW Lib "user32" (ByVal window As LongPtr, ByVal message As Long, ByVal key As LongPtr, ByVal data As LongPtr) As LongPtr

Private Sub EnterText(ByVal inputText As String)
    Dim position As Long
    For position = 1 To Len(inputText)
        SendMessageW GetFocus(), &H102, AscW(Mid$(inputText, position, 1)), 1
        DoEvents
    Next position
End Sub

Public Function VerifyQuietPicker() As String
    On Error GoTo Failed
    Dim entry As Object, tagID As Long, canonical As String, rows As DAO.Recordset
    DoCmd.OpenForm "Report Entry", , , , , , "NEW"
    Set entry = Forms("Report Entry")
    entry!txtLocation.SetFocus
    DoEvents
    If entry!lstLocationMatches.Visible Or entry!lstLocationMatches.ListCount <> 0 Then Err.Raise 513, , "Matches loaded on open. Visible: " & entry!lstLocationMatches.Visible & "; Rows: " & entry!lstLocationMatches.ListCount & "; Source: " & entry!lstLocationMatches.RowSource & "; Text: " & Nz(entry!txtLocation.Value, "") & "; OpenArgs: " & Nz(entry.OpenArgs, "")
    EnterText "ZZZZINVALIDTAG12345"
    If entry!lstLocationMatches.Visible Then Err.Raise 530, , "An invalid tag showed results."
    entry!cboDepartment.SetFocus
    If entry!lblStatus.Caption <> "Please type a valid location tag." Then Err.Raise 531, , "Invalid input did not show the validation error."
    entry!cboDepartment = DMin("DepartmentID", "Departments")
    entry!grpPriority = DMin("PriorityID", "Priority")
    If entry.SaveEntry(False) Then Err.Raise 534, , "An invalid tag was saved."
    entry!txtLocation.SetFocus
    entry!txtLocation.SelStart = 0
    entry!txtLocation.SelLength = Len(entry!txtLocation.Text)
    EnterText "1"
    DoEvents
    If entry!txtLocation.Text <> "1" Then Err.Raise 523, , "The first key did not reach the location field."
    If entry!lstLocationMatches.Visible Or entry!lstLocationMatches.ListCount <> 0 Then Err.Raise 514, , "Matches loaded after one character."
    EnterText "0"
    DoEvents
    If entry!txtLocation.Text <> "10" Then Err.Raise 524, , "The second key did not reach the location field."
    If entry!lstLocationMatches.Visible Or entry!lstLocationMatches.ListCount <> 0 Then Err.Raise 515, , "Matches loaded after two characters."
    EnterText "3"
    DoEvents
    If Not entry!lstLocationMatches.Visible Or entry!lstLocationMatches.ListCount < 1 Or entry!lstLocationMatches.ListCount > 50 Then Err.Raise 516, , "Third key did not open a bounded list. Rows: " & entry!lstLocationMatches.ListCount & "; SQL: " & entry!lstLocationMatches.RowSource & "; OnChange: " & entry!txtLocation.OnChange
    EnterText Chr$(8)
    DoEvents
    If entry!lstLocationMatches.Visible Or entry!lstLocationMatches.ListCount <> 0 Then Err.Raise 517, , "Backspace did not clear the list."
    EnterText "3"
    If entry!lstLocationMatches.ColumnCount <> 2 Or entry!lstLocationMatches.ColumnWidths <> "0;4200" Then Err.Raise 532, , "Only the location tag should be visible."
    If entry!lstLocationMatches.ListCount < 3 Then Err.Raise 518, , "Not enough matches to test arrow navigation."
    tagID = entry!lstLocationMatches.ItemData(1)
    canonical = entry!lstLocationMatches.Column(1, 1)
    SendMessageW GetFocus(), &H100, vbKeyDown, 1
    DoEvents
    If entry.ActiveControl.Name <> "lstLocationMatches" Then Err.Raise 519, , "Down arrow did not focus the results."
    SendMessageW GetFocus(), &H100, vbKeyDown, 1
    SendMessageW GetFocus(), &H100, vbKeyDown, 1
    SendMessageW GetFocus(), &H100, vbKeyUp, 1
    DoEvents
    If Not entry!lstLocationMatches.Visible Or entry!lstLocationMatches.Value <> tagID Then Err.Raise 533, , "Arrow keys did not move through the open list."
    SendMessageW GetFocus(), &H100, vbKeyReturn, 1
    DoEvents
    If entry!lstLocationMatches.Visible Or entry!txtLocation.Value <> canonical Then Err.Raise 520, , "Enter did not select the official tag."
    entry!cboDepartment = DMin("DepartmentID", "Departments")
    entry!grpPriority = DMin("PriorityID", "Priority")
    entry!txtDescription = "Quiet picker verification"
    entry!txtWorkOrder = "123"
    entry!txtPermit = "123"
    If Not entry.SaveEntry(True) Then Err.Raise 521, , "Could not save the selected tag."
    Set rows = CurrentDb.OpenRecordset("SELECT TOP 1 LocationTagID, LocationTag FROM DailyReports ORDER BY ReportID DESC", dbOpenSnapshot)
    If rows!LocationTagID <> tagID Or rows!LocationTag <> canonical Then Err.Raise 522, , "Saved tag differs from selection."
    rows.Close
    VerifyQuietPicker = "PASS"
    Exit Function
Failed:
    VerifyQuietPicker = Err.Number & ": " & Err.Description
End Function

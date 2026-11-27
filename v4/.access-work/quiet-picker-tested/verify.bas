Option Compare Database
Option Explicit

Public Function VerifyQuietPicker() As String
    On Error GoTo Failed
    Dim entry As Object, tagID As Long, canonical As String, rows As DAO.Recordset
    DoCmd.OpenForm "Report Entry", , , , , , "NEW"
    Set entry = Forms("Report Entry")
    entry!txtLocation.SetFocus
    If entry!lstLocationMatches.Visible Or entry!lstLocationMatches.ListCount <> 0 Then Err.Raise 513, , "Matches loaded on open."
    SendKeys "1", True
    DoEvents
    If entry!txtLocation.Text <> "1" Then Err.Raise 523, , "The first key did not reach the location field."
    If entry!lstLocationMatches.Visible Or entry!lstLocationMatches.ListCount <> 0 Then Err.Raise 514, , "Matches loaded after one character."
    SendKeys "0", True
    DoEvents
    If entry!txtLocation.Text <> "10" Then Err.Raise 524, , "The second key did not reach the location field."
    If entry!lstLocationMatches.Visible Or entry!lstLocationMatches.ListCount <> 0 Then Err.Raise 515, , "Matches loaded after two characters."
    SendKeys "3", True
    DoEvents
    If Not entry!lstLocationMatches.Visible Or entry!lstLocationMatches.ListCount < 1 Or entry!lstLocationMatches.ListCount > 50 Then Err.Raise 516, , "Third key did not open a bounded list. Text: " & entry!txtLocation.Text & "; Status: " & entry!lblStatus.Caption
    SendKeys "{BACKSPACE}", True
    DoEvents
    If entry!lstLocationMatches.Visible Or entry!lstLocationMatches.ListCount <> 0 Then Err.Raise 517, , "Backspace did not clear the list."
    SendKeys "3K101", True
    DoEvents
    If entry!lstLocationMatches.ListCount < 1 Then Err.Raise 518, , "Normalized search did not find tags."
    tagID = entry!lstLocationMatches.ItemData(0)
    canonical = entry!lstLocationMatches.Column(1, 0)
    SendKeys "{DOWN}", True
    DoEvents
    If entry.ActiveControl.Name <> "lstLocationMatches" Then Err.Raise 519, , "Down arrow did not focus the results."
    SendKeys "{ENTER}", True
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

Option Compare Database
Option Explicit

Public Function VerifyPicker() As String
    On Error GoTo Failed
    Dim entry As Object, database As DAO.Database, rows As DAO.Recordset
    Dim tagID As Long, expectedTag As String, legacyID As Long
    Set database = CurrentDb
    Set rows = database.OpenRecordset("SELECT TOP 1 ReportID FROM DailyReports WHERE LocationTagID Is Null AND LocationTag Is Not Null", dbOpenSnapshot)
    If Not rows.EOF Then legacyID = rows!ReportID
    rows.Close
    If legacyID > 0 Then
        DoCmd.OpenForm "Report Entry", , , , , , CStr(legacyID)
        Set entry = Forms("Report Entry")
        If entry!txtLocation.Value <> 0 Or entry!txtLocation.ListCount <> 1 Then Err.Raise 513, , "Legacy location did not display."
        DoCmd.Close acForm, "Report Entry"
    End If
    DoCmd.OpenForm "Report Entry", , , , , , "NEW"
    Set entry = Forms("Report Entry")
    If entry!txtLocation.ListCount <> 0 Then Err.Raise 514, , "Opening the form loaded tags."
    entry!txtLocation.RowSource = entry.LocationSearchRows("10")
    If entry!txtLocation.ListCount <> 0 Then Err.Raise 515, , "A short search loaded tags."
    entry!txtLocation.RowSource = entry.LocationSearchRows("103K101")
    If entry!txtLocation.ListCount < 1 Or entry!txtLocation.ListCount > 50 Then Err.Raise 516, , "Search did not return a bounded list."
    tagID = entry!txtLocation.ItemData(0)
    expectedTag = entry!txtLocation.Column(1, 0)
    entry!txtLocation = tagID
    entry!cboDepartment = DMin("DepartmentID", "Departments")
    entry!grpPriority = DMin("PriorityID", "Priority")
    entry!txtDescription = "Location picker verification"
    entry!txtWorkOrder = "123"
    entry!txtPermit = "123"
    If Not entry.SaveEntry(True) Then Err.Raise 517, , "Could not save the selected tag."
    Set rows = database.OpenRecordset("SELECT TOP 1 LocationTagID, LocationTag FROM DailyReports ORDER BY ReportID DESC", dbOpenSnapshot)
    If rows!LocationTagID <> tagID Or rows!LocationTag <> expectedTag Then Err.Raise 518, , "The selection did not save both ID and canonical text."
    rows.Close
    VerifyPicker = "PASS"
    Exit Function
Failed:
    VerifyPicker = Err.Number & ": " & Err.Description
End Function

Option Compare Database
Option Explicit

Public Function VerifyPrint() As String
    On Error GoTo Failed
    Dim reportsForm As Object, printView As Object, rows As DAO.Recordset
    Dim department As Long, expected As Long
    DoCmd.OpenForm "Reports"
    Set reportsForm = Forms("Reports")
    Set rows = CurrentDb.OpenRecordset("SELECT TOP 1 DepartmentID, Count(*) AS Total FROM qryDailyReportGrid GROUP BY DepartmentID ORDER BY Count(*) ASC")
    If rows.EOF Then Err.Raise 513, , "No reports available for the filter check."
    department = rows!DepartmentID
    expected = rows!Total
    rows.Close
    reportsForm!cboDepartment = department
    reportsForm.ApplyFilters
    reportsForm!cboSort = "LocationTag"
    reportsForm.SortGrid False
    If Not reportsForm.OpenDailyPrint() Then Err.Raise 514, , "The print command did not open the report."
    Set printView = Reports("Daily Reports Print")
    If printView.Orientation <> 1 Or printView.Printer.Orientation <> 2 Or printView.Printer.PaperSize <> 9 Then Err.Raise 515, , "Report is not right-to-left A4 landscape."
    If printView.Width + printView.Printer.LeftMargin + printView.Printer.RightMargin > 16837 Then Err.Raise 516, , "Report exceeds the A4 page width."
    Set rows = CurrentDb.OpenRecordset(printView.RecordSource, dbOpenSnapshot)
    rows.MoveLast
    If rows.RecordCount <> expected Or rows!DepartmentID <> department Or InStr(printView.RecordSource, "[LocationTag] ASC") = 0 Then Err.Raise 517, , "Printed records or sort differ from the filtered grid."
    rows.Close
    DoCmd.OutputTo acOutputReport, "Daily Reports Print", "PDFFormat(*.pdf)", CurrentProject.Path & "\sample-print.pdf", False
    If Not printView.ClosePrintView() Then Err.Raise 518, , "Back action failed."
    If CurrentProject.AllReports("Daily Reports Print").IsLoaded Or Not CurrentProject.AllForms("Reports").IsLoaded Then Err.Raise 519, , "Back did not return to Reports."
    DoCmd.Close acForm, "Reports", acSaveNo
    VerifyPrint = "PASS"
    Exit Function
Failed:
    VerifyPrint = Err.Number & ": " & Err.Description
End Function

Option Compare Database
Option Explicit

Public Function VerifyStyledReports() As String
    Dim reportsForm As Object, rowID As Long
    Set reportsForm = Forms("Reports")
    With reportsForm
        If .subReports.Form.Name <> "Report Grid Centered" Then Err.Raise 513
        rowID = .subReports.Form!colReportID
        .cmdGridView_Click
        If .subReports.Form.Name <> "Report Grid" Or .subReports.Form!colReportID <> rowID Then Err.Raise 514
        .subReports.Form.Filter = "ReportID=" & rowID
        .subReports.Form.FilterOn = True
        .cmdGridView_Click
        If .subReports.Form.Name <> "Report Grid Centered" Or Not .subReports.Form.FilterOn Then Err.Raise 515
        If .subReports.Form!colReportID <> rowID Then Err.Raise 516
        .subReports.Form.FilterOn = False
        .SyncGridActions
        If Not .ApplyFilters() Then Err.Raise 517
    End With
    VerifyStyledReports = "PASS"
End Function

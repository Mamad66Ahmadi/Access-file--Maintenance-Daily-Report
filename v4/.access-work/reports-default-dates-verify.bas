Option Compare Database
Option Explicit

Public Function VerifyDefaultDates() As Boolean
    Dim reportsForm As Object, sql As String
    DoCmd.OpenForm "Reports"
    Set reportsForm = Forms("Reports")
    With reportsForm
        If .txtFrom <> DateAdd("d", -29, Date) Or .txtTo <> Date Then Err.Raise 513, , "Incorrect defaults"
        sql = CurrentDb.QueryDefs("qryDailyReportPage").SQL
        If InStr(sql, Format$(.txtFrom, "mm\/dd\/yyyy")) = 0 Then Err.Raise 514, , "Start date was not applied"
        If InStr(sql, Format$(Date + 1, "mm\/dd\/yyyy")) = 0 Then Err.Raise 515, , "End date was not applied"
        .txtFrom = DateAdd("d", -60, Date)
        If Not .ApplyFilters() Then Err.Raise 516, , "Could not apply the custom range"
        If InStr(CurrentDb.QueryDefs("qryDailyReportPage").SQL, Format$(.txtFrom, "mm\/dd\/yyyy")) = 0 Then Err.Raise 517, , "Custom start date was not applied"
        .ClearFilters
        If Not IsNull(.txtFrom) Or Not IsNull(.txtTo) Then Err.Raise 518, , "Clear Filters did not clear the dates"
    End With
    VerifyDefaultDates = True
End Function

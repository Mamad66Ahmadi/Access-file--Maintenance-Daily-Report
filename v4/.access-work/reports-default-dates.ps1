param([switch]$Install)
$ErrorActionPreference = 'Stop'
$front = Join-Path (Split-Path $PSScriptRoot) 'Maintenance Daily Reports.accdb'
$stage = Join-Path $PSScriptRoot 'reports-default-dates.accdb'
$fingerprint = Join-Path $PSScriptRoot 'reports-default-dates.sha256'
if ($Install) {
    if (Test-Path ([IO.Path]::ChangeExtension($front, 'laccdb'))) { throw 'The front end is still open.' }
    if ((Get-FileHash $front).Hash -ne (Get-Content $fingerprint)) { throw 'The front end changed during verification.' }
    $backup = Join-Path $PSScriptRoot ('Maintenance Daily Reports - before default dates ' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.accdb')
    [IO.File]::Replace($stage, $front, $backup)
    'Installed and backed up the front end.'
    return
}
$source = [IO.File]::Open($front, 'Open', 'Read', 'ReadWrite')
try {
    $target = [IO.File]::Create($stage)
    try { $source.CopyTo($target) } finally { $target.Dispose() }
} finally { $source.Dispose() }
$original = (Get-FileHash $stage).Hash
$app = New-Object -ComObject Access.Application
try {
    $app.AutomationSecurity = 1
    $app.OpenCurrentDatabase($stage, $true)
    $formPath = Join-Path $PSScriptRoot 'reports-default-dates.form.txt'
    $app.SaveAsText(2, 'Reports', $formPath)
    $form = [IO.File]::ReadAllText($formPath)
    $before = [regex]'(?m)(^    rowsPerPage = 50\r?\n(?:\r?\n)?)    RefreshGrid\r?$'
    if ($before.Matches($form).Count -ne 1) { throw 'Cannot find the Reports startup filter.' }
    $form = $before.Replace($form, ('${1}    RecentDays 30' + "`r"))
    [IO.File]::WriteAllText($formPath, $form, [Text.Encoding]::Unicode)
    $app.LoadFromText(2, 'Reports', $formPath)
    $verifyPath = Join-Path $PSScriptRoot 'reports-default-dates-verify.bas'
    @'
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
'@ | Set-Content -LiteralPath $verifyPath -Encoding ASCII
    $app.LoadFromText(5, 'VerifyDefaultDatesBuild', $verifyPath)
    if (!$app.Eval('VerifyDefaultDates()')) { throw 'Default date verification failed.' }
    $app.DoCmd.Close(2, 'Reports', 2)
    $app.DoCmd.DeleteObject(5, 'VerifyDefaultDatesBuild')
    $app.CloseCurrentDatabase()
    Set-Content -LiteralPath $fingerprint -Value $original -Encoding ASCII
    'Verified the startup date range, custom From date, and Clear Filters in Microsoft Access.'
} finally { $app.Quit(2) }

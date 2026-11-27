param([switch]$Inspect, [switch]$Install)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
$front = Join-Path $root 'Maintenance Daily Reports.accdb'
$work = Join-Path $PSScriptRoot 'picker-work'
$null = New-Item -ItemType Directory -Path $work -Force
$stage = Join-Path $work 'Maintenance Daily Reports - location picker.accdb'
$formFile = Join-Path $work 'entry.form.txt'
if ($Inspect) {
    Copy-Item -LiteralPath $front -Destination $stage -Force
    $app = New-Object -ComObject Access.Application
    try {
        $app.OpenCurrentDatabase($stage, $true)
        $app.SaveAsText(2, 'Report Entry', $formFile)
        $db = $app.CurrentDb()
        $db.TableDefs.Item('DailyReports').Connect -replace 'PWD=[^;]+','PWD=[hidden]'
        $db.Close()
        $app.CloseCurrentDatabase()
    } finally { $app.Quit(2) }
    return
}

function Require($condition, $message) { if (!$condition) { throw $message } }
$fingerprint = Join-Path $work 'verified-original.sha256'
if ($Install) {
    $lock = $null
    try {
        $lockPath = [IO.Path]::ChangeExtension($front,'laccdb')
        if (Test-Path -LiteralPath $lockPath) { $lock = [IO.File]::Open($lockPath, 'Open', 'ReadWrite', 'None') }
        $handle = [IO.File]::Open($front, 'Open', 'ReadWrite', 'None')
        try { Require ((Get-FileHash -InputStream $handle).Hash -eq (Get-Content -LiteralPath $fingerprint)) 'The original changed during verification.' }
        finally { $handle.Dispose() }
        $backup = Join-Path $PSScriptRoot ('Maintenance Daily Reports - before location picker ' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.accdb')
        [IO.File]::Replace($stage, $front, $backup)
    } finally { if ($lock) { $lock.Dispose() } }
    'Installed the v4 front end and saved a backup.'
    return
}
function Replace-Exact([string]$needle, [string]$replacement) {
    $needle = $needle -replace '\r?\n', "`r`n"
    $replacement = $replacement -replace '\r?\n', "`r`n"
    Require ($script:form.Contains($needle)) "Missing expected form code: $needle"
    $script:form = $script:form.Replace($needle, $replacement)
}
function Attach($db, $name, $connect) {
    $table = $db.CreateTableDef($name)
    $table.Connect = $connect
    $table.SourceTableName = $name
    $db.TableDefs.Append($table)
}
$original = (Get-FileHash -LiteralPath $front).Hash
Copy-Item -LiteralPath $front -Destination $stage -Force
Require ((Get-FileHash -LiteralPath $stage).Hash -eq $original) 'Front end changed while copying.'
$app = New-Object -ComObject Access.Application
try {
    $app.AutomationSecurity = 1
    $app.OpenCurrentDatabase($stage, $true)
    $db = $app.CurrentDb()
    $link = $db.TableDefs.Item('DailyReports').Connect
    $data = Join-Path $root 'Maintenance Data.accdb'
    Require ($link -match '^(.+?DATABASE=)(.+)$' -and $Matches[2] -eq $data) 'DailyReports does not point at the v4 data file.'
    $testData = Join-Path $work 'test-data.accdb'
    Copy-Item -LiteralPath $data -Destination $testData -Force
    $testLink = $Matches[1] + $testData
    $linked = @()
    foreach ($table in $db.TableDefs) {
        if ($table.Connect -eq $link) {
            $linked += $table.Name
            $table.Connect = $testLink
            $table.RefreshLink()
        }
    }
    Attach $db 'LocationTags' $testLink
    $linked += 'LocationTags'
    $app.SaveAsText(2, 'Report Entry', $formFile)
    $script:form = [IO.File]::ReadAllText($formFile)
    $block = [regex]::Matches($form, '(?ms)^                Begin TextBox\r?\n.*?^                End\r?$') | Where-Object Value -Match 'Name ="txtLocation"'
    Require ($null -ne $block) 'Location control not found.'
    $combo = @'
                Begin ComboBox
                    LimitToList = NotDefault
                    AutoExpand = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ColumnCount =4
                    ListRows =12
                    ListWidth =9450
                    Left =5600
                    Top =3890
                    Width =4500
                    Height =420
                    FontSize =10
                    TabIndex =4
                    BorderColor =14475734
                    ForeColor =3158558
                    Name ="txtLocation"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LocationTagID, LocationTag, ObjectType, LongTag FROM LocationTags WHERE LocationTagID=-1"
                    ColumnWidths ="0;2400;1900;5100"
                    AfterUpdate ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="Type 3 characters, then select a location tag"
                End
'@
    Replace-Exact $block.Value $combo
    $script:form = $form -replace 'Caption ="Location tag \* / ', 'Caption ="Location tag * (type 3+ chars) / '
    Replace-Exact 'Private originalValues(0 To 11) As Variant' "Private originalValues(0 To 12) As Variant`nPrivate updatingLocation As Boolean"
    Replace-Exact '        Me.txtLocation = records!LocationTag' '        LoadLocation records!LocationTagID, records!LocationTag'
    Replace-Exact '    Dim fieldName As Variant, progressValue As Variant' "    Dim fieldName As Variant, progressValue As Variant, tagID As Variant, tagText As Variant"
    Replace-Exact '    For Each fieldName In Array("txtWorkOrder", "txtPermit", "txtLocation")' @'
    tagID = Me.txtLocation.Value
    If Not IsNull(tagID) Then
        If CLng(tagID) > 0 Then tagText = DLookup("LocationTag", "LocationTags", "LocationTagID=" & CLng(tagID))
    End If
    If IsNull(tagID) Or (Nz(tagID, 0) <> 0 And IsNull(tagText)) Or (editedID = 0 And Nz(tagID, 0) = 0) Then
        MsgBox "Select a location tag from the list.", vbExclamation, "Location tag"
        Me.txtLocation.SetFocus
        Exit Function
    End If
    For Each fieldName In Array("txtWorkOrder", "txtPermit")
'@
    Replace-Exact '    records!LocationTag = NullIfBlank(Me.txtLocation)' @'
    If CLng(tagID) > 0 Then
        records!LocationTagID = tagID
        records!LocationTag = tagText
    End If
'@
    Replace-Exact '        Me.txtLocation = Null' '        LoadLocation Null'
    Replace-Exact '    Me.txtLocation = values(3)' '    LoadLocation IIf(Nz(values(3), 0) = 0, Null, values(3))'
    Replace-Exact '"Progress", "ActionList")' '"Progress", "ActionList", "LocationTagID")'
    Replace-Exact '    For fieldIndex = 0 To 11' '    For fieldIndex = 0 To 12'
    $vba = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'location-picker.vba') -Raw
    [IO.File]::WriteAllText($formFile, $form + "`r`n`r`n" + ($vba -replace '\r?\n', "`r`n"), [Text.Encoding]::Unicode)
    $app.LoadFromText(2, 'Report Entry', $formFile)

    $verification = @'
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
'@
    $verifyFile = Join-Path $work 'verify.bas'
    [IO.File]::WriteAllText($verifyFile, ($verification -replace '\r?\n', "`r`n") + "`r`n", [Text.Encoding]::ASCII)
    $app.LoadFromText(5, 'VerifyPickerBuild', $verifyFile)
    $result = $app.Eval('VerifyPicker()')
    Require ($result -eq 'PASS') "Picker verification failed: $result"
    $app.DoCmd.DeleteObject(5, 'VerifyPickerBuild')
    $db = $app.CurrentDb()
    foreach ($name in $linked) {
        $table = $db.TableDefs.Item($name)
        $table.Connect = $link
        $table.RefreshLink()
    }
    $db.Close()
    $app.CloseCurrentDatabase()
    Set-Content -LiteralPath $fingerprint -Value $original -Encoding ASCII
    'Verified an empty initial list, limited normalized search, legacy display, and saving an actual selection.'
} finally { $app.Quit(2) }

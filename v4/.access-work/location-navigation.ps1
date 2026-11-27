param([switch]$Install, [string]$Run = '2')
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
$front = Join-Path $root 'Maintenance Daily Reports.accdb'
$work = Join-Path $PSScriptRoot ('location-navigation-' + $Run)
$null = New-Item -ItemType Directory -Path $work -Force
$stage = Join-Path $work 'location-navigation.accdb'
$fingerprint = Join-Path $work 'verified-original.sha256'
function Require($condition, $message) { if (!$condition) { throw $message } }
if ($Install) {
    $lock = $null
    try {
        $lockPath = [IO.Path]::ChangeExtension($front, 'laccdb')
        if (Test-Path -LiteralPath $lockPath) { $lock = [IO.File]::Open($lockPath, 'Open', 'ReadWrite', 'None') }
        $handle = [IO.File]::Open($front, 'Open', 'ReadWrite', 'None')
        try { Require ((Get-FileHash -InputStream $handle).Hash -eq (Get-Content -LiteralPath $fingerprint)) 'The front end changed during verification.' }
        finally { $handle.Dispose() }
        $backup = Join-Path $PSScriptRoot ('Maintenance Daily Reports - before location navigation ' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.accdb')
        [IO.File]::Replace($stage, $front, $backup)
    } finally { if ($lock) { $lock.Dispose() } }
    'Installed and backed up the front end.'
    return
}
function Rewrite([string]$needle, [string]$replacement) {
    $needle = $needle -replace '\r?\n', "`r`n"
    $replacement = $replacement -replace '\r?\n', "`r`n"
    Require ($script:form.Contains($needle)) "Missing form text: $needle"
    $script:form = $script:form.Replace($needle, $replacement)
}
function CopyShared($source, $destination) {
    $sharing = [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete
    $inputStream = [IO.File]::Open($source, 'Open', 'Read', $sharing)
    try {
        $outputStream = [IO.File]::Create($destination)
        try { $inputStream.CopyTo($outputStream) } finally { $outputStream.Dispose() }
    } finally { $inputStream.Dispose() }
}
CopyShared $front $stage
$original = (Get-FileHash -LiteralPath $stage).Hash
$testData = Join-Path $work 'test-data.accdb'
CopyShared (Join-Path $root 'Maintenance Data.accdb') $testData
$app = New-Object -ComObject Access.Application
try {
    $app.AutomationSecurity = 1
    $app.OpenCurrentDatabase($stage, $true)
    $db = $app.CurrentDb()
    $link = $db.TableDefs.Item('DailyReports').Connect
    $testLink = $link -replace '(?<=DATABASE=).+$', $testData
    $linked = @()
    foreach ($table in $db.TableDefs) {
        if ($table.Connect -eq $link) {
            $linked += $table.Name
            $table.Connect = $testLink
            $table.RefreshLink()
        }
    }
    $formFile = Join-Path $work 'entry.form.txt'
    $app.SaveAsText(2, 'Report Entry', $formFile)
    $script:form = [IO.File]::ReadAllText($formFile)
    $list = ([regex]::Matches($form, '(?ms)^                Begin ListBox\r?\n.*?^                End\r?$') | Where-Object Value -Match 'Name ="lstLocationMatches"').Value
    Require ($list -match 'ColumnCount =4') 'The location list was not found in the expected format.'
    $updatedList = $list -replace '(?m)(ColumnCount =)4\b', '${1}2' -replace '(?m)(Left =)1100\b','${1}5600' -replace '(?m)(Width =)9000\b','${1}4500' -replace '(?m)(LayoutCachedLeft =)1100\b','${1}5600' -replace '0;2070;2100;4500','0;4200'
    Rewrite $list $updatedList
    Rewrite 'LocationTagID, LocationTag, ObjectType, LongTag FROM LocationTags' 'LocationTagID, LocationTag FROM LocationTags'
    Rewrite '    locationID = selectedID' @'
    Me.lstLocationMatches.RowSource = LocationRows()
    locationID = selectedID
'@
    Rewrite 'Select a location tag from the search results before saving.' 'Please type a valid location tag.'
    Rewrite @'
Private Sub txtLocation_AfterUpdate()
    MarkChanged
End Sub
'@ @'
Private Sub txtLocation_AfterUpdate()
    MarkChanged
    If Len(Trim$(Nz(Me.txtLocation.Value, ""))) > 0 And IsNull(ChosenLocationID()) Then
        If Me.lstLocationMatches.ListCount = 0 Then Me.lblStatus.Caption = "Please type a valid location tag."
    End If
End Sub
'@
    Rewrite '    sql = LocationSearchRows(Me.txtLocation.text)' @'
    If Me.lblStatus.Caption = "Please type a valid location tag." Then Me.lblStatus.Caption = ""
    sql = LocationSearchRows(Me.txtLocation.Text)
'@
    Rewrite @'
Private Sub lstLocationMatches_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
'@ @'
Private Sub lstLocationMatches_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim index As Long
    Select Case KeyCode
        Case vbKeyDown, vbKeyUp
            index = Me.lstLocationMatches.ListIndex
            If KeyCode = vbKeyDown Then
                index = index + 1
            Else
                index = index - 1
            End If
            KeyCode = 0
            If index >= 0 And index < Me.lstLocationMatches.ListCount Then
                Me.lstLocationMatches.Value = Me.lstLocationMatches.ItemData(index)
            End If
'@
    [IO.File]::WriteAllText($formFile, $form, [Text.Encoding]::Unicode)
    $app.LoadFromText(2, 'Report Entry', $formFile)
    $harness = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'quiet-picker-results\verify.bas'))
    $harness = $harness.Replace('    If entry!lstLocationMatches.Visible Or entry!lstLocationMatches.ListCount <> 0 Then Err.Raise 513, , "Matches loaded on open."', @'
    DoEvents
    If entry!lstLocationMatches.Visible Or entry!lstLocationMatches.ListCount <> 0 Then Err.Raise 513, , "Matches loaded on open. Visible: " & entry!lstLocationMatches.Visible & "; Rows: " & entry!lstLocationMatches.ListCount & "; Source: " & entry!lstLocationMatches.RowSource & "; Text: " & Nz(entry!txtLocation.Value, "") & "; OpenArgs: " & Nz(entry.OpenArgs, "")
'@)
    $harness = $harness.Replace('    EnterText "1"', @'
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
'@)
    $from = $harness.IndexOf('    EnterText "3K101"')
    $to = $harness.IndexOf('    SendMessageW GetFocus(), &H100, vbKeyReturn, 1', $from)
    Require ($from -ge 0 -and $to -gt $from) 'Could not locate the keyboard test.'
    $harness = $harness.Substring(0,$from) + (@'
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
'@ -replace '\r?\n', "`r`n") + "`r`n" + $harness.Substring($to)
    $verifyFile = Join-Path $work 'verify.bas'
    [IO.File]::WriteAllText($verifyFile, $harness, [Text.Encoding]::ASCII)
    $app.LoadFromText(5, 'VerifyQuietPickerBuild', $verifyFile)
    $result = $app.Eval('VerifyQuietPicker()')
    Require ($result -eq 'PASS') "Picker verification failed: $result"
    $app.DoCmd.DeleteObject(5, 'VerifyQuietPickerBuild')
    $db = $app.CurrentDb()
    foreach ($name in $linked) {
        $table = $db.TableDefs.Item($name)
        $table.Connect = $link
        $table.RefreshLink()
    }
    $db.Close()
    $app.CloseCurrentDatabase()
    Set-Content -LiteralPath $fingerprint -Value $original -Encoding ASCII
    'Verified that invalid tags are rejected, only the tag is visible, and Up, Down, and Enter select and save the expected ID.'
} finally { $app.Quit(2) }

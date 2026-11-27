param([switch]$Install)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
$front = Join-Path $root 'Maintenance Daily Reports.accdb'
$work = Join-Path $PSScriptRoot 'quiet-picker-results'
$null = New-Item -ItemType Directory -Path $work -Force
$stage = Join-Path $work 'location-quiet.accdb'
$fingerprint = Join-Path $work 'verified-original.sha256'
function Require($condition, $message) { if (!$condition) { throw $message } }
if ($Install) {
    $lock = $null
    try {
        $lockPath = [IO.Path]::ChangeExtension($front,'laccdb')
        if (Test-Path -LiteralPath $lockPath) { $lock = [IO.File]::Open($lockPath, 'Open', 'ReadWrite', 'None') }
        $handle = [IO.File]::Open($front, 'Open', 'ReadWrite', 'None')
        try { Require ((Get-FileHash -InputStream $handle).Hash -eq (Get-Content -LiteralPath $fingerprint)) 'The original changed during verification.' }
        finally { $handle.Dispose() }
        $backup = Join-Path $PSScriptRoot ('Maintenance Daily Reports - before quiet location search ' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.accdb')
        [IO.File]::Replace($stage, $front, $backup)
    } finally { if ($lock) { $lock.Dispose() } }
    'Installed the verified front end and saved a backup.'
    return
}
function Rewrite([string]$needle, [string]$replacement) {
    $needle = $needle -replace '\r?\n', "`r`n"
    $replacement = $replacement -replace '\r?\n', "`r`n"
    Require ($script:form.Contains($needle)) "Missing form text: $needle"
    $script:form = $script:form.Replace($needle, $replacement)
}
$snapshot = Join-Path $work 'input.accdb'
Copy-Item -LiteralPath $front -Destination $snapshot -Force
$original = (Get-FileHash -LiteralPath $snapshot).Hash
Copy-Item -LiteralPath $snapshot -Destination $stage -Force
Require ((Get-FileHash -LiteralPath $stage).Hash -eq $original) 'Front end changed while copying.'
$app = New-Object -ComObject Access.Application
try {
    $app.AutomationSecurity = 1
    $app.OpenCurrentDatabase($stage, $true)
    $db = $app.CurrentDb()
    $link = $db.TableDefs.Item('DailyReports').Connect
    $testData = Join-Path $work 'quiet-test-data.accdb'
    Copy-Item -LiteralPath (Join-Path $root 'Maintenance Data.accdb') -Destination $testData -Force
    $testLink = $link -replace '(?<=DATABASE=).+$', $testData
    $linked = @()
    foreach ($table in $db.TableDefs) {
        if ($table.Connect -eq $link) {
            $linked += $table.Name
            $table.Connect = $testLink
            $table.RefreshLink()
        }
    }
    Require ($linked -contains 'LocationTags') 'LocationTags is not linked to the report data file.'
    $formFile = Join-Path $work 'entry.form.txt'
    $app.SaveAsText(2, 'Report Entry', $formFile)
    $script:form = [IO.File]::ReadAllText($formFile)
    $block = [regex]::Matches($form, '(?ms)^                Begin ComboBox\r?\n.*?^                End\r?$') | Where-Object Value -Match 'Name ="txtLocation"'
    Require ($null -ne $block) 'Location combo box not found.'
    Rewrite $block.Value @'
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    Left =5600
                    Top =3890
                    Width =4500
                    Height =420
                    FontSize =10
                    TabIndex =4
                    BorderColor =14475734
                    ForeColor =3158558
                    Name ="txtLocation"
                    AfterUpdate ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    OnKeyDown ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="Type to search and select a location tag"
                End
'@
    $panel = @'
                Begin ListBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    ColumnCount =4
                    Left =1100
                    Top =4310
                    Width =9000
                    Height =1850
                    FontSize =9
                    Name ="lstLocationMatches"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT LocationTagID, LocationTag, ObjectType, LongTag FROM LocationTags WHERE LocationTagID=-1"
                    ColumnWidths ="0;2070;2100;4500"
                    OnClick ="[Event Procedure]"
                    OnKeyDown ="[Event Procedure]"
                    FontName ="Tahoma"
                End
'@
    $ending = [regex]::Match($form, '(?m)^            End\r?\n        End\r?\n    End\r?\nEnd\r?\nCodeBehindForm')
    Require $ending.Success 'Could not locate the end of the form controls.'
    $script:form = $form.Insert($ending.Index, ($panel -replace '\r?\n', "`r`n") + "`r`n")
    Rewrite 'Location tag * (type 3+ chars)' 'Location tag * (search)'
    Rewrite 'Private updatingLocation As Boolean' "Private locationID As Variant`nPrivate locationLabel As Variant"
    Rewrite "    Me.txtDate = Date`n    Me.chkActionList = False" "    Me.txtDate = Date`n    Me.chkActionList = False`n    LoadLocation Null"
    Rewrite '    tagID = Me.txtLocation.value' '    tagID = ChosenLocationID()'
    Rewrite '        MsgBox "Select a location tag from the list.", vbExclamation, "Location tag"' '        Me.lblStatus.Caption = "Select a location tag from the search results before saving."'
    Rewrite 'Me.txtLocation.value, Me.cboDepartment.value' 'ChosenLocationID(), Me.cboDepartment.value'
    Rewrite '    Me.cmdCancel.SetFocus' "    Me.cmdCancel.SetFocus`n    HideLocationMatches"
    Rewrite @'
Private Sub Form_Timer()
    If EditExpired(entryCreatedAt) Then LockEntry entryCreatedAt
End Sub
'@ @'
Private Sub Form_Timer()
    If editedID <> 0 Then
        If EditExpired(entryCreatedAt) Then LockEntry entryCreatedAt
    End If
    If Me.lstLocationMatches.Visible Then
        Select Case Me.ActiveControl.Name
            Case "txtLocation", "lstLocationMatches"
            Case Else: HideLocationMatches
        End Select
    End If
End Sub
'@
    $oldLoader = [regex]::Match($form, '(?ms)^Private Sub LoadLocation\b.*?^End Sub\r?\n\r?\n').Value
    Require ([bool]$oldLoader) 'Could not find the combo loading code.'
    $script:form = $form.Replace($oldLoader, '')
    $oldEvents = [regex]::Match($form, '(?ms)^Private Sub txtLocation_Change\(\).*?^Private Sub txtLocation_NotInList\b.*?^End Sub').Value
    Require ([bool]$oldEvents) 'Could not find the combo event code.'
    Rewrite $oldEvents (Get-Content -LiteralPath (Join-Path $PSScriptRoot 'location-quiet.vba') -Raw)
    [IO.File]::WriteAllText($formFile, $form, [Text.Encoding]::Unicode)
    $app.LoadFromText(2, 'Report Entry', $formFile)

    $verify = @'
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
    If entry!lstLocationMatches.Visible Or entry!lstLocationMatches.ListCount <> 0 Then Err.Raise 513, , "Matches loaded on open."
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
    EnterText "3K101"
    DoEvents
    If entry!lstLocationMatches.ListCount < 1 Then Err.Raise 518, , "Normalized search did not find tags."
    tagID = entry!lstLocationMatches.ItemData(0)
    canonical = entry!lstLocationMatches.Column(1, 0)
    SendMessageW GetFocus(), &H100, vbKeyDown, 1
    DoEvents
    If entry.ActiveControl.Name <> "lstLocationMatches" Then Err.Raise 519, , "Down arrow did not focus the results."
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
'@
    $verifyFile = Join-Path $work 'verify.bas'
    [IO.File]::WriteAllText($verifyFile, ($verify -replace '\r?\n', "`r`n") + "`r`n", [Text.Encoding]::ASCII)
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
    'Verified character-by-character changes, backspace, keyboard selection, and saving.'
} finally { $app.Quit(2) }

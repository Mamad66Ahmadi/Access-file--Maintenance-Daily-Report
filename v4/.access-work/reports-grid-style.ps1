param([switch]$Inspect, [switch]$Install)
$ErrorActionPreference = 'Stop'
$front = Join-Path (Split-Path $PSScriptRoot) 'Maintenance Daily Reports.accdb'
$stage = Join-Path $PSScriptRoot 'reports-grid-style.accdb'
$source = Join-Path $PSScriptRoot 'reports-grid-before.txt'
$after = Join-Path $PSScriptRoot 'reports-grid-after.txt'
$fingerprint = Join-Path $PSScriptRoot 'reports-grid-original.sha256'

if ($Install) {
    if (Test-Path -LiteralPath ([IO.Path]::ChangeExtension($front, 'laccdb'))) { throw 'Close the front end in Access before installing.' }
    if ((Get-FileHash -LiteralPath $front).Hash -ne (Get-Content $fingerprint)) { throw 'The front end changed during verification.' }
    $backup = Join-Path $PSScriptRoot ('Maintenance Daily Reports - before Reports styling ' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.accdb')
    [IO.File]::Replace($stage, $front, $backup)
    'Installed the styled Reports grid and backed up the front end.'
    return
}

function Color([string]$hex) {
    return [Convert]::ToInt32(($hex.Substring(4,2) + $hex.Substring(2,2) + $hex.Substring(0,2)), 16)
}

function Snapshot([string]$name) {
    Add-Type -AssemblyName System.Windows.Forms, System.Drawing
    $app.DoCmd.OpenForm('Reports')
    if ($name -eq 'reports-grid-after.png') {
        if ($app.Eval('VerifyStyledReports()') -ne 'PASS') { throw "Reports view verification failed." }
    }
    for ($i = $app.Forms.Count - 1; $i -ge 0; $i--) {
        $openName = $app.Forms.Item($i).Name
        if ($openName -and $openName -ne 'Reports') { $app.DoCmd.Close(2, $openName, 2) }
    }
    if (-not ('ReportsGridWindow' -as [type])) {
        Add-Type 'using System; using System.Runtime.InteropServices; public static class ReportsGridWindow { [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int command); [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd); }'
    }
    $handle = [IntPtr][long]$app.hWndAccessApp()
    $null = [ReportsGridWindow]::ShowWindow($handle, 3)
    $null = [ReportsGridWindow]::SetForegroundWindow($handle)
    Start-Sleep -Milliseconds 850
    $screen = [Windows.Forms.Screen]::PrimaryScreen.Bounds
    $size = New-Object Drawing.Size($screen.Width, $screen.Height)
    $bitmap = New-Object Drawing.Bitmap($size.Width, $size.Height)
    $graphics = [Drawing.Graphics]::FromImage($bitmap)
    try {
        $graphics.CopyFromScreen($screen.Left, $screen.Top, 0, 0, $size)
        $bitmap.Save((Join-Path $PSScriptRoot $name), [Drawing.Imaging.ImageFormat]::Png)
    } finally {
        $graphics.Dispose()
        $bitmap.Dispose()
        $app.DoCmd.Close(2, 'Reports', 2)
    }
}

Copy-Item -LiteralPath $front -Destination $stage -Force
$original = (Get-FileHash -LiteralPath $stage).Hash
$app = New-Object -ComObject Access.Application
try {
    $app.AutomationSecurity = 1
    $app.OpenCurrentDatabase($stage, $true)
    $app.SaveAsText(2, 'Report Grid', $source)
    if ($Inspect) {
        $app.SaveAsText(2, 'Reports', (Join-Path $PSScriptRoot 'reports-parent-before.txt'))
        return
    }
    Snapshot 'reports-grid-before.png'
    $app.DoCmd.OpenForm('Report Grid', 1, [Type]::Missing, [Type]::Missing, [Type]::Missing, 1)
    $grid = $app.Forms.Item('Report Grid')
    $grid.DatasheetFontName = 'Tahoma'
    $grid.DatasheetFontHeight = 11
    $grid.DatasheetForeColor = Color '26374B'
    $grid.DatasheetBackColor = Color 'FFFFFF'
    $grid.DatasheetAlternateBackColor = Color 'F1F5FA'
    $grid.DatasheetGridlinesColor = Color 'D6E0EA'
    $grid.DatasheetGridlinesBehavior = 3
    foreach ($control in $grid.Controls) {
        if ($control.ControlType -eq 109) { $control.TopMargin = 180 }
    }
    $palette = @(
        @('DFEAF9','23486D'), @('DDF1EA','205747'), @('F8EDCF','644912'), @('EBE4F6','513772'),
        @('F6E2E7','722F43'), @('DCEFF1','1D535D'), @('F8E6D9','713D24'), @('E6EBF1','354B60')
    )
    $conditions = $grid.Controls.Item('colDepartmentName').FormatConditions
    if ($conditions.Count -ne $palette.Count) { throw "Unexpected number of department color rules: $($conditions.Count)" }
    for ($index = 0; $index -lt $palette.Count; $index++) {
        $condition = $conditions.Item($index)
        $condition.BackColor = Color $palette[$index][0]
        $condition.ForeColor = Color $palette[$index][1]
        $condition.FontBold = $true
    }
    $app.DoCmd.Close(2, 'Report Grid', 1)
    $app.SaveAsText(2, 'Report Grid', $after)
    $centered = Join-Path $PSScriptRoot 'reports-grid-centered.txt'
    $definition = [IO.File]::ReadAllText($after)
    $definition = $definition.Replace('    DefaultView =2', "    DefaultView =1`r`n    Orientation =1")
    $definition = $definition.Replace('    ScrollBars =0', '    ScrollBars =3')
    $definition = $definition.Replace('        Begin Section', "        Begin FormHeader`r`n            Height =440`r`n            Name =`"GridHeader`"`r`n        End`r`n        Begin Section")
    $definition = $definition -replace '(?m)^    Width =\d+', '    Width =27800'
    [IO.File]::WriteAllText($centered, $definition, [Text.Encoding]::Unicode)
    $app.LoadFromText(2, 'Report Grid Centered', $centered)
    $app.DoCmd.OpenForm('Report Grid Centered', 1)
    $grid = $app.Forms.Item('Report Grid Centered')
    $grid.Section(1).BackColor = Color '263C56'
    $grid.Section(0).BackColor = Color 'FFFFFF'
    $grid.Section(0).AlternateBackColor = Color 'F1F5FA'
    $grid.RecordSelectors = $false
    $grid.DividingLines = $false
    $columns = @(
        @('colReportID', 'ID', 750), @('colReportDate', 'Date', 1500),
        @('colJalaliDate', 'Jalali date', 1560), @('colLocationTag', 'Location tag', 2130),
        @('colDepartmentName', 'Department', 1880), @('colDescription', ('Description / ' + [char]0x0634 + [char]0x0631 + [char]0x062D), 8200),
        @('colProjectName', 'Project', 1900), @('colActualStart', 'Actual start', 2200),
        @('colPriority', 'Priority', 1400), @('colProgress', 'Progress %', 1320),
        @('colActionList', 'Action list', 1350), @('colWorkOrderNumber', 'WO no.', 1520),
        @('colPermitNumber', 'Permit no.', 1500)
    )
    foreach ($control in $grid.Controls) {
        if ($control.ControlType -eq 100 -or $control.Name -in 'colDepartmentID','colDescriptionSort') {
            $control.Visible = $false
            $control.Move(0, 0, 100, 100)
        }
    }
    $totalWidth = ($columns | ForEach-Object { [int]$_[2] } | Measure-Object -Sum).Sum
    $left = $totalWidth
    foreach ($column in $columns) {
        $control = $grid.Controls.Item($column[0])
        $width = [int]$column[2]
        $left -= $width
        $label = $app.CreateControl('Report Grid Centered', 100, 1, '', '', $left, 80, $width, 300)
        $label.Caption = $column[1]
        $label.FontName = 'Tahoma'
        $label.FontSize = 10
        $label.FontBold = $true
        $label.ForeColor = Color 'FFFFFF'
        $label.TextAlign = 2
        if ($control.ControlType -eq 109) {
            $control.Move(($left + 35), 140, ($width - 70), 380)
            $control.FontName = 'Tahoma'
            $control.FontSize = 11
            $control.ForeColor = Color '26374B'
            $control.BorderStyle = 0
            $control.SpecialEffect = 0
            $control.TopMargin = 55
            $control.BottomMargin = 20
            $control.LeftMargin = 50
            $control.RightMargin = 50
            $control.BackStyle = [int]($column[0] -eq 'colDepartmentName')
            $control.BackColor = Color 'E6EBF1'
            $control.Locked = $true
        } else {
            $control.Move(($left + [int](($width - 235) / 2)), 240, 235, 300)
        }
    }
    $grid.Width = $totalWidth + 15
    $grid.Section(0).Height = 660
    if ($grid.Section(0).Height -ne 660) { throw "Unexpected detail height: $($grid.Section(0).Height)" }
    $app.DoCmd.Close(2, 'Report Grid Centered', 1)
    $parentPath = Join-Path $PSScriptRoot 'reports-parent-styled.txt'
    $app.SaveAsText(2, 'Reports', $parentPath)
    $parent = [IO.File]::ReadAllText($parentPath)
    $parent = $parent.Replace('SourceObject ="Form.Report Grid"', 'SourceObject ="Form.Report Grid Centered"')
    $parent = $parent.Replace('    Me.subReports.Form.RowHeight = 660', '')
    $button = @'
                Begin CommandButton
                    Left =14600
                    Top =3360
                    Width =1730
                    Height =460
                    FontSize =9
                    Name ="cmdGridView"
                    Caption ="Column menus"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Switch between the centered grid and the Access column menus"
                End
'@
    $parent = $parent.Replace("Left =14730`r`n                    Top =3480`r`n                    Width =3450", "Left =16410`r`n                    Top =3480`r`n                    Width =1770")
    $parent = $parent.Replace('                Begin Subform', ($button -replace '\r?\n', "`r`n") + "`r`n                Begin Subform")
    $handler = @'

Public Sub cmdGridView_Click()
    SyncGridActions
    With Me.subReports
        If .Form.Name = "Report Grid Centered" Then
            .SourceObject = "Form.Report Grid"
            .Form.RowHeight = 660
            Me.cmdGridView.Caption = "Centered rows"
        Else
            .SourceObject = "Form.Report Grid Centered"
            Me.cmdGridView.Caption = "Column menus"
        End If
        .Form.Filter = headerCriteria
        .Form.FilterOn = (Len(headerCriteria) > 0)
    End With
End Sub
'@
    $parent += "`r`n" + ($handler -replace '\r?\n', "`r`n") + "`r`n"
    [IO.File]::WriteAllText($parentPath, $parent, [Text.Encoding]::Unicode)
    $app.LoadFromText(2, 'Reports', $parentPath)
    $verifyPath = Join-Path $PSScriptRoot 'reports-grid-verify.bas'
    @'
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
'@ | Set-Content -LiteralPath $verifyPath -Encoding ASCII
    $app.LoadFromText(5, 'VerifyStyledReportsBuild', $verifyPath)
    Snapshot 'reports-grid-after.png'
    $app.DoCmd.DeleteObject(5, 'VerifyStyledReportsBuild')
    $app.CloseCurrentDatabase()
    Set-Content -LiteralPath $fingerprint -Value $original -Encoding ASCII
} finally { $app.Quit(2) }

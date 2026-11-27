param([switch]$Install)
$ErrorActionPreference = 'Stop'
$front = Join-Path (Split-Path $PSScriptRoot) 'Maintenance Daily Reports.accdb'
$stage = Join-Path $PSScriptRoot 'reports-scroll-right.accdb'
$fingerprint = Join-Path $PSScriptRoot 'reports-scroll-original.sha256'
if ($Install) {
    if (Test-Path -LiteralPath ([IO.Path]::ChangeExtension($front, 'laccdb'))) { throw 'The front end is still open.' }
    if ((Get-FileHash -LiteralPath $front).Hash -ne (Get-Content $fingerprint)) { throw 'The front end changed during verification.' }
    $backup = Join-Path $PSScriptRoot ('Maintenance Daily Reports - before scroll position ' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.accdb')
    [IO.File]::Replace($stage, $front, $backup)
    'Installed and backed up the front end.'
    return
}
Copy-Item -LiteralPath $front -Destination $stage -Force
$original = (Get-FileHash -LiteralPath $stage).Hash
Add-Type @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public static class ReportsScroll {
    public delegate bool Visitor(IntPtr hwnd, IntPtr unused);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr hwnd, Visitor visitor, IntPtr unused);
    [DllImport("user32.dll", CharSet=CharSet.Auto)] static extern int GetClassName(IntPtr hwnd, StringBuilder name, int max);
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr hwnd, int message, IntPtr code, IntPtr source);
    [StructLayout(LayoutKind.Sequential)] public struct Rect { public int left, top, right, bottom; }
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr hwnd, out Rect rect);
    public static bool Horizontal(IntPtr hwnd) { Rect rect; GetWindowRect(hwnd,out rect); return rect.right-rect.left > rect.bottom-rect.top; }
    [StructLayout(LayoutKind.Sequential)] public struct ScrollInfo { public int size, mask, min, max, page, pos, track; }
    [DllImport("user32.dll")] public static extern bool GetScrollInfo(IntPtr hwnd, int bar, ref ScrollInfo info);
    public static string ClassName(IntPtr hwnd) { var name = new StringBuilder(100); GetClassName(hwnd,name,100); return name.ToString(); }
    public static string Inspect(IntPtr hwnd) {
        var line = new StringBuilder();
        for(int bar=0;bar<=2;bar++){
            var info = new ScrollInfo { size=28, mask=7 };
            if(GetScrollInfo(hwnd,bar,ref info))
                line.AppendFormat(" bar {0} {1} {2} page={3} pos={4}",bar,info.min,info.max,info.page,info.pos);
        }
        return ClassName(hwnd) + " " + hwnd + line;
    }
    public static IntPtr[] Children(IntPtr hwnd) {
        var windows = new List<IntPtr>();
        EnumChildWindows(hwnd, (child, unused) => { windows.Add(child); return true; }, IntPtr.Zero);
        return windows.ToArray();
    }
}
'@
$app = New-Object -ComObject Access.Application
try {
    $app.AutomationSecurity = 1
    $app.OpenCurrentDatabase($stage, $true)
    $verify = Join-Path $PSScriptRoot 'reports-scroll-verify.bas'
    @'
Option Compare Database
Option Explicit
Public Function ScrollTestWindow(ByVal viewName As String) As LongPtr
    With Forms("Reports")!subReports
        If viewName <> "Initial" Then
            .Width = 12000
            Forms("Reports").cmdGridView_Click
        End If
        ScrollTestWindow = .Form.hWnd
    End With
End Function
'@ | Set-Content -LiteralPath $verify -Encoding ASCII
    $app.LoadFromText(5, 'VerifyReportsScrollBuild', $verify)
    $formPath = Join-Path $PSScriptRoot 'reports-scroll-form.txt'
    $app.SaveAsText(2, 'Reports', $formPath)
    $form = [IO.File]::ReadAllText($formPath)
    function Rewrite([string]$before, [string]$after) {
        $before = $before -replace '\r?\n', "`r`n"
        $after = $after -replace '\r?\n', "`r`n"
        if (!$script:form.Contains($before)) { throw "Missing form text: $before" }
        $script:form = $script:form.Replace($before,$after)
    }
    Rewrite 'Private activeCriteria As String' @'
Private Type GridScrollInfo
    cbSize As Long
    mask As Long
    minimum As Long
    maximum As Long
    page As Long
    position As Long
    trackPosition As Long
End Type

Private Type GridWindowRect
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal window As LongPtr, ByVal relation As Long) As LongPtr
Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameW" (ByVal window As LongPtr, ByVal buffer As LongPtr, ByVal capacity As Long) As Long
Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal window As LongPtr, ByRef area As GridWindowRect) As Long
Private Declare PtrSafe Function GetScrollInfo Lib "user32" (ByVal window As LongPtr, ByVal bar As Long, ByRef info As GridScrollInfo) As Long
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal window As LongPtr, ByVal message As Long, ByVal command As LongPtr, ByVal source As LongPtr) As LongPtr

Private Const FirstChild As Long = 5
Private Const NextSibling As Long = 2
Private Const ScrollBarControl As Long = 2
Private Const ScrollRangePagePosition As Long = 7
Private Const HorizontalScroll As Long = &H114
Private Const ThumbPosition As Long = 4

Private activeCriteria As String
'@
    Rewrite @'
    DoCmd.Maximize
    FitHeaderRTL
End Sub
'@ @'
    DoCmd.Maximize
    FitHeaderRTL
    PositionGridAtRight
End Sub
'@
    Rewrite @'
        .Form.FilterOn = (Len(headerCriteria) > 0)
    End With
End Sub
'@ @'
        .Form.FilterOn = (Len(headerCriteria) > 0)
    End With
    PositionGridAtRight
End Sub

Private Sub PositionGridAtRight()
    Dim formWindow As LongPtr, child As LongPtr
    Dim className As String, area As GridWindowRect, info As GridScrollInfo
    Dim endpoint As Long
    formWindow = Me.subReports.Form.hWnd
    child = GetWindow(formWindow, FirstChild)
    Do While child <> 0
        className = String$(40, vbNullChar)
        If Left$(className, GetClassName(child, StrPtr(className), Len(className))) = "NUIScrollbar" Then
            If GetClientRect(child, area) <> 0 And area.right > area.bottom Then
                info.cbSize = LenB(info)
                info.mask = ScrollRangePagePosition
                If GetScrollInfo(child, ScrollBarControl, info) <> 0 And info.page > 0 Then
                    endpoint = info.maximum - info.page
                    If Me.subReports.Form.CurrentView = 2 Then endpoint = endpoint + 1
                    If endpoint > info.minimum And endpoint <> info.position Then
                        SendMessage formWindow, HorizontalScroll, CLngPtr(endpoint) * 65536 + ThumbPosition, child
                    End If
                End If
                Exit Sub
            End If
        End If
        child = GetWindow(child, NextSibling)
    Loop
End Sub
'@
    [IO.File]::WriteAllText($formPath, $form, [Text.Encoding]::Unicode)
    $app.LoadFromText(2, 'Reports', $formPath)
    $app.DoCmd.OpenForm('Reports')
    foreach ($view in 'Initial', 'Datasheet', 'Centered') {
        $hwnd = [IntPtr][long]$app.Eval("ScrollTestWindow('$view')")
        $bars = @([ReportsScroll]::Children($hwnd) | Where-Object { [ReportsScroll]::ClassName($_) -eq 'NUIScrollbar' -and [ReportsScroll]::Horizontal($_) })
        if ($bars.Count -ne 1) { throw "$view horizontal scrollbar count: $($bars.Count)" }
        $state = [ReportsScroll]::Inspect($bars[0])
        if ($state -notmatch 'bar 2 (\d+) (\d+) page=(\d+) pos=(\d+)') { throw "Unknown scrollbar state: $state" }
        $expected = [Math]::Max([int]$Matches[1], [int]$Matches[2] - [int]$Matches[3] + [int]($view -eq 'Datasheet'))
        if ([int]$Matches[4] -ne $expected -and [int]$Matches[3] -gt 0) { throw "$view scrollbar not at the right edge: $state (expected $expected)" }
        "Verified $view : $state"
    }
    $null = [ReportsScroll]::SendMessage($hwnd, 0x114, [IntPtr]4, $bars[0])
    $manualState = [ReportsScroll]::Inspect($bars[0])
    Start-Sleep -Milliseconds 1200
    if ($manualState -eq $state -or [ReportsScroll]::Inspect($bars[0]) -ne $manualState) {
        throw 'Manual horizontal scrolling did not remain in place.'
    }
    $app.DoCmd.Close(2, 'Reports', 2)
    $app.DoCmd.DeleteObject(5, 'VerifyReportsScrollBuild')
    $app.CloseCurrentDatabase()
    Set-Content -LiteralPath $fingerprint -Value $original -Encoding ASCII
} finally { $app.Quit(2) }

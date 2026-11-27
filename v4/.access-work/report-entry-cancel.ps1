param([switch]$Install)
$ErrorActionPreference = 'Stop'
$front = Join-Path (Split-Path $PSScriptRoot) 'Maintenance Daily Reports.accdb'
$stage = Join-Path $PSScriptRoot 'report-entry-cancel.accdb'
$fingerprint = Join-Path $PSScriptRoot 'report-entry-cancel.sha256'
if ($Install) {
    if (Test-Path ([IO.Path]::ChangeExtension($front, 'laccdb'))) { throw 'The front end is still open.' }
    if ((Get-FileHash $front).Hash -ne (Get-Content $fingerprint)) { throw 'The front end changed during verification.' }
    $backup = Join-Path $PSScriptRoot ('Maintenance Daily Reports - before Cancel fix ' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.accdb')
    [IO.File]::Replace($stage, $front, $backup)
    'Installed the Cancel fix and backed up the front end.'
    return
}
Copy-Item -LiteralPath $front -Destination $stage -Force
$original = (Get-FileHash $stage).Hash
$app = New-Object -ComObject Access.Application
try {
    $app.AutomationSecurity = 1
    $app.OpenCurrentDatabase($stage, $true)
    $formPath = Join-Path $PSScriptRoot 'report-entry-cancel.form.txt'
    $app.SaveAsText(2, 'Report Entry', $formPath)
    $form = [IO.File]::ReadAllText($formPath)
    $oldBlock = [regex]::Match($form, '(?ms)^Private Sub cmdCancel_Click\(\)\r?\n.*?^End Sub\r?\n\r?\nPrivate Sub Form_Unload\(Cancel As Integer\)\r?\n.*?^End Sub')
    $message = [regex]::Match($oldBlock.Value, 'MsgBox\("([^"]+)".*?, "([^"]+)"\)')
    if (!$message.Success) { throw 'Cannot find the existing Cancel prompt.' }
    $declaration = @'
Private Declare PtrSafe Function EntryMessageBox Lib "user32" Alias "MessageBoxW" (ByVal owner As LongPtr, ByVal text As LongPtr, ByVal title As LongPtr, ByVal flags As Long) As Long

Private editedID As Long
'@
    $form = $form.Replace('Private editedID As Long', ($declaration -replace '\r?\n', "`r`n"))
    $newBlock = @'
Public Sub cmdCancel_Click()
    If Not ConfirmDiscard Then Exit Sub
    hasChanges = False
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = Not ConfirmDiscard()
End Sub

Private Function ConfirmDiscard() As Boolean
    ConfirmDiscard = True
    If Not hasChanges Then Exit Function

    If Me.lstLocationMatches.Visible Then
        Me.cmdCancel.SetFocus
        HideLocationMatches
    End If
    Me.TimerInterval = 0
    ConfirmDiscard = (EntryMessageBox(Me.hWnd, StrPtr("{MESSAGE}"), StrPtr("{TITLE}"), _
        vbQuestion Or vbYesNo Or vbDefaultButton2 Or vbMsgBoxRtlReading Or vbMsgBoxRight) = vbYes)
    If Not ConfirmDiscard And editedID <> 0 And Not Me.txtLocation.Locked Then Me.TimerInterval = 1000
End Function
'@
    $newBlock = $newBlock.Replace('{MESSAGE}', $message.Groups[1].Value).Replace('{TITLE}', $message.Groups[2].Value)
    $form = $form.Replace($oldBlock.Value, ($newBlock -replace '\r?\n', "`r`n"))
    [IO.File]::WriteAllText($formPath, $form, [Text.Encoding]::Unicode)
    $app.LoadFromText(2, 'Report Entry', $formPath)

    $modulePath = Join-Path $PSScriptRoot 'report-entry-cancel-verify.bas'
    @'
Option Compare Database
Option Explicit
Private Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal window As LongPtr, ByVal message As Long, ByVal key As LongPtr, ByVal data As LongPtr) As LongPtr

Public Function PrepareCancelTest(ByVal openArgument As String) As LongPtr
    Dim entry As Object, index As Long
    Const prefix As String = "103"
    DoCmd.OpenForm "Report Entry", , , , , , openArgument
    Set entry = Forms("Report Entry")
    entry!txtLocation.SetFocus
    entry!txtLocation.SelStart = 0
    entry!txtLocation.SelLength = Len(entry!txtLocation.Text)
    For index = 1 To Len(prefix)
        SendMessage GetFocus(), &H102, AscW(Mid$(prefix, index, 1)), 1
    Next
    If Not entry!lstLocationMatches.Visible Then Err.Raise 520, , "Suggestions did not open."
    PrepareCancelTest = entry.hWnd
End Function

Public Function TryCancelEntry(ByVal closeFromWindow As Boolean) As Boolean
    Dim entry As Object
    Set entry = Forms("Report Entry")
    entry!cmdCancel.SetFocus
    If closeFromWindow Then
        DoCmd.Close acForm, "Report Entry"
    Else
        entry.cmdCancel_Click
    End If
    TryCancelEntry = Not CurrentProject.AllForms("Report Entry").IsLoaded
End Function

Public Function CancelTestPreservedDraft() As Boolean
    With Forms("Report Entry")
        CancelTestPreservedDraft = (.txtLocation.Value = "103" And Not .lstLocationMatches.Visible)
    End With
End Function

Public Function EntryEditableTestID() As String
    EntryEditableTestID = Nz(DMax("ReportID", "DailyReports", "created_at >= DateAdd('h', -24, Now())"), "")
End Function
'@ | Set-Content -LiteralPath $modulePath -Encoding ASCII
    $app.LoadFromText(5, 'VerifyReportEntryCancelBuild', $modulePath)
    $entryArgs = @('NEW')
    $recent = [string]$app.Eval('EntryEditableTestID()')
    if ($recent) { $entryArgs += $recent }
    foreach ($argument in $entryArgs) {
        $entryWindow = [long]$app.Eval("PrepareCancelTest('$argument')")
        foreach ($answer in 7,6) {
            $worker = Start-Job -ArgumentList $entryWindow,$answer -ScriptBlock {
                param([long]$owner, [int]$answer)
                $ErrorActionPreference = 'Stop'
                Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class CancelDialog {
    public delegate bool Visitor(IntPtr hwnd, IntPtr data);
    [DllImport("user32.dll")] static extern bool EnumWindows(Visitor callback, IntPtr data);
    [DllImport("user32.dll")] static extern IntPtr GetWindow(IntPtr hwnd, int flag);
    [DllImport("user32.dll")] static extern IntPtr GetDlgItem(IntPtr hwnd, int id);
    [DllImport("user32.dll")] static extern bool PostMessage(IntPtr hwnd, int message, IntPtr command, IntPtr source);
    public static bool Reply(long owner, int answer) {
        bool replied = false;
        EnumWindows((hwnd, unused) => {
            if (GetWindow(hwnd, 4).ToInt64() != owner) return true;
            var button = GetDlgItem(hwnd, answer);
            if (button == IntPtr.Zero) return true;
            replied = PostMessage(hwnd, 0x111, (IntPtr)answer, button);
            return !replied;
        }, IntPtr.Zero);
        return replied;
    }
}
"@
                for ($attempt = 0; $attempt -lt 160; $attempt++) {
                    if ([CancelDialog]::Reply($owner, $answer)) { return 'answered' }
                    Start-Sleep -Milliseconds 50
                }
                throw 'The owned discard prompt did not appear.'
            }
            $fromWindow = $answer -eq 6 -and $argument -ne 'NEW'
            $closed = [bool]$app.Eval("TryCancelEntry($fromWindow)")
            if ((Receive-Job $worker -Wait -AutoRemoveJob) -ne 'answered') { throw 'Could not answer the discard prompt.' }
            if ($closed -ne ($answer -eq 6)) { throw "Unexpected close result for $argument / button $answer." }
            if (!$closed -and !$app.Eval('CancelTestPreservedDraft()')) { throw 'Draft was not preserved after No.' }
        }
        "Verified both answers for Report Entry ($argument)."
    }
    $app.DoCmd.DeleteObject(5, 'VerifyReportEntryCancelBuild')
    $app.CloseCurrentDatabase()
    Set-Content -LiteralPath $fingerprint -Value $original -Encoding ASCII
} finally { $app.Quit(2) }

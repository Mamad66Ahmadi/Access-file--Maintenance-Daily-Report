$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
$live = Join-Path $root 'Maintenance Data.accdb'
$stage = Join-Path $PSScriptRoot 'object-type-staged.accdb'
$backup = Join-Path $PSScriptRoot ('Maintenance Data - before object type ' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.accdb')
$types = [Collections.Generic.Dictionary[string,string]]::new([StringComparer]::OrdinalIgnoreCase)
Import-Csv -LiteralPath (Join-Path $root '07 objects to import.csv') | ForEach-Object {
    $types.Add($_.Long_Tag.Trim(), $_.Object_Type.Trim())
}
$hash = (Get-FileHash -LiteralPath $live).Hash
Copy-Item -LiteralPath $live -Destination $stage
if ((Get-FileHash -LiteralPath $stage).Hash -ne $hash) { throw 'The data file changed while copying.' }
$app = New-Object -ComObject Access.Application
try {
    $app.OpenCurrentDatabase($stage, $true)
    $db = $app.CurrentDb()
    $db.Execute('ALTER TABLE LocationTags ADD COLUMN ObjectType TEXT(100)', 128)
    $engine = $app.DBEngine
    $workspace = $engine.Workspaces.Item(0)
    $rows = $db.OpenRecordset('LocationTags', 1)
    $written = 0
    try {
        $path = $rows.Fields.Item('LongTag')
        $objectType = $rows.Fields.Item('ObjectType')
        $workspace.BeginTrans()
        try {
            while (!$rows.EOF) {
                $key = [string]$path.Value
                if ($types.ContainsKey($key) -and $types[$key]) {
                    $rows.Edit()
                    $objectType.Value = $types[$key]
                    $rows.Update()
                    $written++
                }
                $rows.MoveNext()
            }
            $workspace.CommitTrans()
        } catch { $workspace.Rollback(); throw }
    } finally { $rows.Close() }
    $rows = $db.OpenRecordset('SELECT LongTag, ObjectType FROM LocationTags', 4)
    $verified = 0
    try {
        $path = $rows.Fields.Item(0)
        $objectType = $rows.Fields.Item(1)
        while (!$rows.EOF) {
            $key = [string]$path.Value
            $actual = $objectType.Value
            if ($null -eq $actual -or $actual -eq [DBNull]::Value) { $actual = '' }
            if ($types.ContainsKey($key)) {
                if ($actual -cne $types[$key]) { throw "Type mismatch for $key." }
                $verified++
            }
            $rows.MoveNext()
        }
    } finally { $rows.Close() }
    if ($verified -ne $types.Count) { throw 'Not every CSV path matched an existing location tag.' }
    $null = $db.Relations.Item('FK_DailyReports_LocationTags')
    $db.Close()
    $app.CloseCurrentDatabase()
    "Verified $verified CSV tags, with $written object types and $($verified - $written) blanks."
} finally { $app.Quit(2) }

$lockPath = [IO.Path]::ChangeExtension($live, 'laccdb')
$lock = $null
try {
    if (Test-Path -LiteralPath $lockPath) { $lock = [IO.File]::Open($lockPath, 'Open', 'ReadWrite', 'None') }
    $handle = [IO.File]::Open($live, 'Open', 'ReadWrite', 'None')
    try {
        if ((Get-FileHash -InputStream $handle).Hash -ne $hash) { throw 'The live database changed during verification.' }
    } finally { $handle.Dispose() }
    [IO.File]::Replace($stage, $live, $backup)
} finally { if ($lock) { $lock.Dispose() } }
'Installed the verified data file and backed up the original.'

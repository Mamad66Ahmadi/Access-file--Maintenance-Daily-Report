param([switch]$Install)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
$database = Join-Path $root 'Maintenance Data.accdb'
$work = Join-Path $PSScriptRoot 'location-import'
$null = New-Item -ItemType Directory -Path $work -Force
$stage = Join-Path $work 'Maintenance Data - with Location Tags.accdb'
$hashFile = Join-Path $work 'verified-original.sha256'
$review = Join-Path $work 'Location tags to review.csv'
function Require($condition, $message) { if (!$condition) { throw $message } }
function Install-Database {
    $lock = [IO.Path]::ChangeExtension($database,'laccdb')
    $lockHandle = $null
    try {
        if (Test-Path $lock) { $lockHandle = [IO.File]::Open($lock, 'Open', 'ReadWrite', 'None') }
        $databaseHandle = [IO.File]::Open($database, 'Open', 'ReadWrite', 'None')
        try {
            Require ((Get-FileHash -InputStream $databaseHandle).Hash -eq (Get-Content $hashFile)) 'The data file changed during verification. Run the import again from the current file.'
        } finally { $databaseHandle.Dispose() }
        $backup = Join-Path $PSScriptRoot ('Maintenance Data - before location tags ' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.accdb')
        [IO.File]::Replace($stage, $database, $backup)
        'Installed in Maintenance Data.accdb with a backup.'
    } finally { if ($lockHandle) { $lockHandle.Dispose() } }
}
if ($Install) { Install-Database; return }

function Key([string]$tag) { ($tag -replace '[\s-]', '').ToUpperInvariant() }
function Scalar($db, $sql) {
    $records = $db.OpenRecordset($sql, 4)
    try { $records.Fields.Item(0).Value } finally { $records.Close() }
}
function Lookup([string]$field, [string]$value) {
    $literal = $value.Replace("'", "''")
    $records = $db.OpenRecordset("SELECT TOP 2 LocationTagID FROM LocationTags WHERE [$field] = '$literal'", 4)
    try {
        if ($records.EOF) { return 0 }
        $id = $records.Fields.Item(0).Value
        $records.MoveNext()
        if (!$records.EOF) { return -1 }
        $id
    } finally { $records.Close() }
}

$csv = Join-Path $root '07 objects to import.csv'
$tags = @(Import-Csv -LiteralPath $csv | ForEach-Object {
    $tag = $_.Object_Tag.Trim()
    $path = $_.Long_Tag.Trim()
    $key = Key $tag
    Require ($tag -and $path -and $key -and $path.Length -le 100 -and $tag.Length -le 100) 'Invalid tag or path in the source CSV.'
    [pscustomobject]@{LocationTag=$tag; LongTag=$path; SearchKey=$key}
})
Require ($tags.Count -gt 0) 'No tags to import.'
$original = (Get-FileHash $database).Hash
Copy-Item -LiteralPath $database -Destination $stage -Force
Require ((Get-FileHash $stage).Hash -eq $original) 'The data file changed while it was being copied.'

$app = New-Object -ComObject Access.Application
try {
    $app.OpenCurrentDatabase($stage, $true)
    $db = $app.CurrentDb()
    $before = Scalar $db 'SELECT Count(*) FROM DailyReports'
    $db.Execute('CREATE TABLE LocationTags (LocationTagID COUNTER CONSTRAINT PK_LocationTags PRIMARY KEY, LocationTag TEXT(100) NOT NULL, LongTag TEXT(100) NOT NULL, SearchKey TEXT(100) NOT NULL)', 128)
    $db.Execute('CREATE UNIQUE INDEX UX_LocationTags_LongTag ON LocationTags (LongTag)', 128)
    $db.Execute('CREATE INDEX IX_LocationTags_LocationTag ON LocationTags (LocationTag)', 128)
    $db.Execute('CREATE INDEX IX_LocationTags_SearchKey ON LocationTags (SearchKey)', 128)
    $engine = $app.DBEngine
    $workspace = $engine.Workspaces.Item(0)
    $writer = $db.OpenRecordset('LocationTags', 1)
    try {
        $tagField = $writer.Fields.Item('LocationTag')
        $pathField = $writer.Fields.Item('LongTag')
        $keyField = $writer.Fields.Item('SearchKey')
        $workspace.BeginTrans()
        try {
            foreach ($row in $tags) {
                $writer.AddNew()
                $tagField.Value = $row.LocationTag
                $pathField.Value = $row.LongTag
                $keyField.Value = $row.SearchKey
                $writer.Update()
            }
            $workspace.CommitTrans()
        } catch { $workspace.Rollback(); throw }
    } finally { $writer.Close() }
    Require ((Scalar $db 'SELECT Count(*) FROM LocationTags') -eq $tags.Count) 'Import count did not match the source CSV.'
    $db.Execute('ALTER TABLE DailyReports ADD COLUMN LocationTagID LONG', 128)

    $macro = Join-Path $work 'report-audit-original.xml'
    $temporary = Join-Path $work 'report-audit-import.xml'
    $app.SaveAsText(12, 'DailyReports', $macro)
    $rules = [xml](Get-Content -LiteralPath $macro -Raw)
    $ns = New-Object System.Xml.XmlNamespaceManager($rules.NameTable)
    $ns.AddNamespace('a', $rules.DocumentElement.NamespaceURI)
    $stamp = $rules.SelectSingleNode('//a:Action[a:Argument[@Name="Field" and text()="ModifiedAt"] and a:Argument[@Name="Value" and text()="Now()"]]', $ns)
    Require ($null -ne $stamp) 'Could not identify the report edit timestamp rule.'
    $null = $stamp.ParentNode.RemoveChild($stamp)
    $rules.Save($temporary)
    $app.LoadFromText(12, 'DailyReports', $temporary)
    $db = $app.CurrentDb()
    $reportRows = $db.OpenRecordset('SELECT ReportID, LocationTag FROM DailyReports', 4)
    $unmatched = @()
    $matched = 0
    try {
        while (!$reportRows.EOF) {
            $reportId = $reportRows.Fields.Item('ReportID').Value
            $legacy = ([string]$reportRows.Fields.Item('LocationTag').Value).Trim()
            if ($legacy) {
                $id = Lookup 'LongTag' $legacy
                if ($id -eq 0) { $id = Lookup 'LocationTag' $legacy }
                if ($id -eq 0) { $id = Lookup 'SearchKey' (Key $legacy) }
                if ($id -gt 0) {
                    $db.Execute("UPDATE DailyReports SET LocationTagID=$id WHERE ReportID=$reportId", 128)
                    $matched++
                } else {
                    $reason = if ($id -eq -1) { 'Multiple matches' } else { 'No match' }
                    $unmatched += [pscustomobject]@{ReportID=$reportId; LocationTag=$legacy; Reason=$reason}
                }
            }
            $reportRows.MoveNext()
        }
    } finally { $reportRows.Close() }
    $app.LoadFromText(12, 'DailyReports', $macro)
    $db = $app.CurrentDb()
    $db.Execute('CREATE INDEX IX_DailyReports_LocationTagID ON DailyReports (LocationTagID)', 128)
    $db.Execute('ALTER TABLE DailyReports ADD CONSTRAINT FK_DailyReports_LocationTags FOREIGN KEY (LocationTagID) REFERENCES LocationTags (LocationTagID)', 128)
    $relation = $db.Relations.Item('FK_DailyReports_LocationTags')
    Require ($relation.Table -eq 'LocationTags' -and $relation.ForeignTable -eq 'DailyReports' -and ($relation.Attributes -band 2) -eq 0) 'The relationship is not enforced.'
    if ($before -gt 0) {
        $first = Scalar $db 'SELECT Min(ReportID) FROM DailyReports'
        $rejected = $false
        try { $db.Execute("UPDATE DailyReports SET LocationTagID=-1 WHERE ReportID=$first",128) } catch { $rejected = $true }
        Require $rejected 'The relationship accepted an invalid location ID.'
    }
    Require ((Scalar $db 'SELECT Count(*) FROM DailyReports') -eq $before) 'The number of reports changed.'
    Require ((Scalar $db 'SELECT Count(*) FROM DailyReports WHERE LocationTagID Is Not Null') -eq $matched) 'Backfill count did not match.'
    $db.Close()
    $app.CloseCurrentDatabase()
    if ($unmatched.Count) { $unmatched | Export-Csv -LiteralPath $review -NoTypeInformation -Encoding UTF8 }
    Set-Content -LiteralPath $hashFile -Value $original -Encoding ASCII
    "Verified $($tags.Count) tags, $before reports, $matched linked reports, $($unmatched.Count) nonblank tags to review."
} finally { $app.Quit(2) }

Install-Database

<#
.SYNOPSIS
    Round-trip a GPMC migration table (.migtable) through CSV so it can be bulk-edited in Excel.

.DESCRIPTION
    -ToCsv    .migtable -> CSV with one row per entry: Type, Source, Action, Destination
    -FromCsv  CSV -> .migtable (UTF-16, the format GPMC and Import-GPO expect)

    Action values (what each entry becomes in the target):
      SameAsSource   keep it as-is (built-ins like Administrators, Authenticated Users)
      RelativeName   same account name in the target domain: grp@ad.engr.wisc.edu -> grp@ad.wisc.edu
      None           drop it (deleted accounts, ENGR-only admins)
      Explicit       use the value in the Destination column
    A filled-in Destination is used whenever Action is blank or SameAsSource, so you can just
    type destinations. Blank Destination and blank Action = SameAsSource.

    Bulk edits in Excel, for example:
      - Filter Source on "@ad.engr.wisc.edu", set Action = RelativeName, fill down
      - Filter Source beginning "S-1-5-21-" (deleted accounts), set Action = None
      - Filter Type = UNCPath, put the new share paths in Destination
      - Use a formula for renamed groups, e.g. ="ENGR-"&LEFT(B2,FIND("@",B2)-1)&"@ad.wisc.edu",
        then paste as values before saving
    Save as CSV UTF-8.

    -FromCsv -Validate (run on the target side, needs the ActiveDirectory module) checks that
    every RelativeName / Explicit principal exists in the target domain and lists the ones that
    don't, so you find them before the import does.

.EXAMPLE
    .\MigTable-Csv.ps1 -ToCsv   -MigTable .\MigrationTable-template.migtable -Csv .\migtable.csv
    # ...edit migtable.csv in Excel, save as CSV UTF-8...
    .\MigTable-Csv.ps1 -FromCsv -Csv .\migtable.csv -MigTable .\MigrationTable.migtable
    .\MigTable-Csv.ps1 -FromCsv -Csv .\migtable.csv -MigTable .\MigrationTable.migtable -Validate
#>
[CmdletBinding(DefaultParameterSetName = 'ToCsv')]
param(
    [Parameter(ParameterSetName = 'ToCsv', Mandatory = $true)]   [switch]$ToCsv,
    [Parameter(ParameterSetName = 'FromCsv', Mandatory = $true)] [switch]$FromCsv,
    [Parameter(Mandatory = $true)] [string]$MigTable,
    [Parameter(Mandatory = $true)] [string]$Csv,
    [Parameter(ParameterSetName = 'FromCsv')] [switch]$Validate,
    [Parameter(ParameterSetName = 'FromCsv')] [string]$TargetDomain = 'ad.wisc.edu'
)

$MigNS = 'http://www.microsoft.com/GroupPolicy/GPOOperations/MigrationTable'
$PrincipalTypes = 'User', 'Computer', 'LocalGroup', 'GlobalGroup', 'UniversalGroup', 'Unknown'

function Read-MigTable([string]$Path) {
    # GPMC writes UTF-16, but copies are often re-saved as UTF-8/ASCII with the UTF-16
    # declaration left in, which XmlDocument refuses. Decode by BOM, then drop the declaration.
    $bytes = [System.IO.File]::ReadAllBytes((Resolve-Path $Path))
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) { $text = [Text.Encoding]::Unicode.GetString($bytes, 2, $bytes.Length - 2) }
    elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) { $text = [Text.Encoding]::BigEndianUnicode.GetString($bytes, 2, $bytes.Length - 2) }
    elseif ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { $text = [Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3) }
    else { $text = [Text.Encoding]::UTF8.GetString($bytes) }
    $text = $text -replace '^\s*<\?xml[^>]*\?>', ''
    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml($text)
    return $doc
}

if ($ToCsv) {
    $doc = Read-MigTable $MigTable
    $nsm = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $nsm.AddNamespace('m', $MigNS)
    $rows = foreach ($m in $doc.SelectNodes('//m:Mapping', $nsm)) {
        $action = 'SameAsSource'; $dest = ''
        if ($m.SelectSingleNode('m:DestinationNone', $nsm)) { $action = 'None' }
        elseif ($m.SelectSingleNode('m:DestinationByRelativeName', $nsm)) { $action = 'RelativeName' }
        elseif ($d = $m.SelectSingleNode('m:Destination', $nsm)) { $action = 'Explicit'; $dest = $d.InnerText }
        [PSCustomObject]@{
            Type        = $m.SelectSingleNode('m:Type', $nsm).InnerText
            Source      = $m.SelectSingleNode('m:Source', $nsm).InnerText
            Action      = $action
            Destination = $dest
        }
    }
    $rows | Sort-Object Type, Source | Export-Csv $Csv -NoTypeInformation -Encoding UTF8
    Write-Host "$(@($rows).Count) entries -> $Csv"
    $rows | Group-Object Type | Sort-Object Name | ForEach-Object { Write-Host ("  {0,-15} {1}" -f $_.Name, $_.Count) }
    return
}

# ---- FromCsv
$rows = @(Import-Csv $Csv)
if (-not $rows.Count) { throw "CSV is empty: $Csv" }
foreach ($c in 'Type', 'Source') {
    if ($rows[0].PSObject.Properties.Name -notcontains $c) { throw "CSV has no '$c' column" }
}

$problems = @()
$settings = New-Object System.Xml.XmlWriterSettings
$settings.Indent = $true
$settings.Encoding = [Text.Encoding]::Unicode          # UTF-16 LE with BOM, as GPMC writes it
$w = [System.Xml.XmlWriter]::Create($MigTable, $settings)
$w.WriteStartDocument()
$w.WriteStartElement('MigrationTable', $MigNS)
$w.WriteAttributeString('xmlns', 'xsd', $null, 'http://www.w3.org/2001/XMLSchema')
$w.WriteAttributeString('xmlns', 'xsi', $null, 'http://www.w3.org/2001/XMLSchema-instance')
$counts = @{}
$seen = @{}
$line = 1
foreach ($r in $rows) {
    $line++
    $type = ([string]$r.Type).Trim(); $src = ([string]$r.Source).Trim()
    $action = ([string]$r.Action).Trim(); $dest = ([string]$r.Destination).Trim()
    if (-not $type -or -not $src) { $problems += "line ${line}: blank Type or Source - skipped"; continue }
    $key = "$type|$src".ToLower()
    if ($seen.ContainsKey($key)) { $problems += "line ${line}: duplicate of line $($seen[$key]) ($src) - skipped"; continue }
    $seen[$key] = $line
    # A filled-in Destination wins over SameAsSource (the value -ToCsv pre-fills), so typing
    # destinations without touching Action works. With None / RelativeName it's reported.
    if (-not $action -or ($dest -and $action -match '^(?i)sameassource$')) { $action = if ($dest) { 'Explicit' } else { 'SameAsSource' } }
    elseif ($dest -and $action -match '^(?i)(none|relativename)$') {
        $problems += "line ${line}: Destination '$dest' ignored because Action is '$action' ($src)"
    }
    switch -Regex ($action) {
        '^(?i)sameassource$' { $action = 'SameAsSource' }
        '^(?i)relativename$' { $action = 'RelativeName' }
        '^(?i)none$'         { $action = 'None' }
        '^(?i)explicit$'     { $action = 'Explicit' }
        default { $problems += "line ${line}: unknown Action '$action' for $src - kept as SameAsSource"; $action = 'SameAsSource' }
    }
    if ($action -eq 'Explicit' -and -not $dest) { $problems += "line ${line}: Explicit with no Destination for $src - kept as SameAsSource"; $action = 'SameAsSource' }
    if ($action -eq 'RelativeName' -and $type -eq 'UNCPath') { $problems += "line ${line}: RelativeName isn't valid for a UNC path ($src) - kept as SameAsSource"; $action = 'SameAsSource' }

    $w.WriteStartElement('Mapping', $MigNS)
    $w.WriteElementString('Type', $MigNS, $type)
    $w.WriteElementString('Source', $MigNS, $src)
    switch ($action) {
        'SameAsSource' { $w.WriteStartElement('DestinationSameAsSource', $MigNS); $w.WriteEndElement() }
        'RelativeName' { $w.WriteStartElement('DestinationByRelativeName', $MigNS); $w.WriteEndElement() }
        'None'         { $w.WriteStartElement('DestinationNone', $MigNS); $w.WriteEndElement() }
        'Explicit'     { $w.WriteElementString('Destination', $MigNS, $dest) }
    }
    $w.WriteEndElement()
    $counts[$action] = 1 + [int]$counts[$action]

    # Optional: does the target principal exist?
    if ($Validate -and $type -in $PrincipalTypes -and $action -in 'RelativeName', 'Explicit') {
        $name = if ($action -eq 'Explicit') { $dest } else { $src }
        $sam = ($name -replace '@.*$', '') -replace '^.*\\', ''
        if (-not $script:adLoaded) { Import-Module ActiveDirectory -ErrorAction Stop; $script:adLoaded = $true }
        $found = Get-ADObject -Server $TargetDomain -LDAPFilter "(sAMAccountName=$($sam -replace '([\\*()])', '\$1'))" -ErrorAction SilentlyContinue
        if (-not $found) { $problems += "line ${line}: '$sam' not found in $TargetDomain (Source $src, $action)" }
    }
}
$w.WriteEndElement()
$w.WriteEndDocument()
$w.Close()

Write-Host "$($seen.Count) entries -> $MigTable"
$counts.GetEnumerator() | Sort-Object Name | ForEach-Object { Write-Host ("  {0,-13} {1}" -f $_.Name, $_.Value) }
if ($problems) {
    Write-Host "`n$($problems.Count) issue(s):" -ForegroundColor Yellow
    $problems | ForEach-Object { Write-Host "  $_" -ForegroundColor Yellow }
}

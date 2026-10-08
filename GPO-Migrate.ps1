<#
.SYNOPSIS
    ENGR -> CAD GPO migration: export the in-scope GPOs from ad.engr.wisc.edu, import them into
    ad.wisc.edu as "ENGR-<name>", and link them to the mapped target OUs.

.DESCRIPTION
    Two modes, run on different sides:

      -Mode Export   (source domain, read-only against AD)
          1. Reads the in-scope list (GPO-InScope.csv).
          2. Finds every OU / domain-root link for those GPOs, with link order, enabled
             and enforced state (links.csv). Site links are recorded but not migrated.
          3. Backs up each GPO, writes an HTML settings report, records WMI filter and
             security filtering (manifest.csv).
          4. Writes a migration-table template listing every security principal and UNC
             path the GPOs reference (MigrationTable-template.migtable).

      -Mode Import   (target domain)
          1. Imports each GPO under its target name (ENGR- prefix, or TargetGPOName if set).
             An existing GPO with that name is NOT overwritten unless -OverwriteExisting.
             A same-named GPO that this process didn't create is never touched.
          2. Maps each source link to a target OU through OU-Map.csv and creates the link,
             keeping the source enabled/enforced state and the relative order of the
             migrated links on each OU (they're added after any links already there).
          3. Links for GPOs with security filtering or a WMI filter are created DISABLED,
             because neither carries over on import. Enable them after fixing the filter.
          4. Never changes a link that already exists, so re-running is safe.
          5. -OnlyLinked skips GPOs that have no mapped link yet, so CAD doesn't fill up with
             unlinked GPOs. Re-run later as more OUs are mapped; only the new ones are added.

    Always do a -DryRun first on each side. It does every lookup and writes the result
    files, but changes nothing in AD and backs nothing up.

.INPUTS
    GPO-InScope.csv  - save the workbook's "Group Policy Inventory" sheet as CSV (UTF-8).
                       Columns used: GPOName, InScope (rows with "Yes" are migrated),
                       TargetGPOName (optional name override). With no InScope column,
                       every row is treated as in scope.
    OU-Map.csv       - save the workbook's "OU Structure Map" sheet as CSV (UTF-8).
                       Columns used: "ENGR OU" (source DN) and "CAD OU" (target DN).
                       Rows with a blank or #N/A target are treated as unmapped.
    Column names are set in $Columns below if the sheets change.

.EXAMPLE
    # Source side
    .\GPO-Migrate.ps1 -Mode Export -DryRun
    .\GPO-Migrate.ps1 -Mode Export
    # copy C:\Temp\GPO_Migration to the target side, review manifest.csv / links.csv
    # Target side
    .\GPO-Migrate.ps1 -Mode Import -DryRun
    .\GPO-Migrate.ps1 -Mode Import
    .\GPO-Migrate.ps1 -Mode Import -LinkState Disabled   # link everything disabled, enable per wave

.NOTES
    PowerShell 5.1+, RSAT Group Policy Management + AD module.
    Export needs read access to the source GPOs; Import needs rights to create GPOs in the
    target domain and link them on the target OUs.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Export', 'Import')]
    [string]$Mode,

    [string]$WorkRoot     = 'C:\Temp\GPO_Migration',
    [string]$InScopeCsv   = '',          # default: <WorkRoot>\GPO-InScope.csv
    [string]$OUMapCsv     = '',          # default: <WorkRoot>\OU-Map.csv
    [string]$SourceDomain = 'ad.engr.wisc.edu',
    [string]$TargetDomain = 'ad.wisc.edu',
    [string]$NamePrefix   = 'ENGR-',

    # AsSource = keep each link's source enabled/disabled state
    # Disabled = create every link disabled (enable them per wave)
    [ValidateSet('AsSource', 'Disabled')]
    [string]$LinkState = 'AsSource',

    [switch]$OverwriteExisting,          # re-import settings into a GPO this process created earlier
    [switch]$OnlyLinked,                 # import only GPOs with at least one mapped OU link; the rest wait for a later run
    [switch]$DryRun
)

# ==============================================================
#  CONFIG
# ==============================================================
$Columns = @{
    GPOName    = 'GPOName'        # in-scope list
    InScope    = 'InScope'
    TargetName = 'TargetGPOName'
    SourceOU   = 'ENGR OU'        # OU map
    TargetOU   = 'CAD OU'
}
$InScopeValue   = 'Yes'
$MarkerPrefix   = '[ENGR-MIG'     # stamped in the target GPO description: [ENGR-MIG {source guid}]

if (-not $InScopeCsv) { $InScopeCsv = Join-Path $WorkRoot 'GPO-InScope.csv' }
if (-not $OUMapCsv)   { $OUMapCsv   = Join-Path $WorkRoot 'OU-Map.csv' }
$BackupPath        = Join-Path $WorkRoot 'Backups'
$ReportPath        = Join-Path $WorkRoot 'Reports'
$ManifestFile      = Join-Path $WorkRoot 'manifest.csv'
$LinksFile         = Join-Path $WorkRoot 'links.csv'
$NotFoundFile      = Join-Path $WorkRoot 'not-found.csv'
$MigTableTemplate  = Join-Path $WorkRoot 'MigrationTable-template.migtable'
$MigTable          = Join-Path $WorkRoot 'MigrationTable.migtable'   # used by Import if present
$ImportResults     = Join-Path $WorkRoot 'import_results.csv'
$LinkResults       = Join-Path $WorkRoot 'link_results.csv'
$LogFile           = Join-Path $WorkRoot 'migration.log'

# ==============================================================
#  HELPERS
# ==============================================================
function Write-Log {
    param([string]$Message, [ValidateSet('INFO', 'WARN', 'ERROR', 'OK')][string]$Level = 'INFO')
    $line = "[{0}][{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    $color = @{ INFO = 'Cyan'; WARN = 'Yellow'; ERROR = 'Red'; OK = 'Green' }[$Level]
    Write-Host $line -ForegroundColor $color
    Add-Content -Path $LogFile -Value $line
}

function Ensure-Directory([string]$Path) {
    if (-not (Test-Path $Path)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
}

function Import-RequiredModules {
    foreach ($m in 'GroupPolicy', 'ActiveDirectory') {
        if (-not (Get-Module -ListAvailable -Name $m)) {
            throw "PowerShell module '$m' is not installed (RSAT)."
        }
        Import-Module $m -ErrorAction Stop
    }
}

function Test-Blank($v) { return ($null -eq $v) -or ([string]$v).Trim() -eq '' }

function Get-TargetName([string]$SourceName, [string]$Override) {
    if (-not (Test-Blank $Override)) { return $Override.Trim() }
    if ($SourceName.StartsWith($NamePrefix, [StringComparison]::OrdinalIgnoreCase)) { return $SourceName }
    return "$NamePrefix$SourceName"
}

# True when security filtering makes the GPO apply to fewer objects than "everyone":
# Authenticated Users isn't allowed to apply it, or someone is denied. Extra allowed groups
# alongside Authenticated Users don't narrow anything.
function Test-NarrowFiltering([string]$Filtering) {
    $entries = @($Filtering -split ';\s*' | Where-Object { $_ })
    $everyone = @($entries | Where-Object { $_ -notlike 'DENY:*' -and $_ -match '(^|\\)Authenticated Users$' }).Count -gt 0
    $denies = @($entries | Where-Object { $_ -like 'DENY:*' }).Count -gt 0
    return (-not $everyone) -or $denies
}

function Get-SafeFileName([string]$Name) { return ($Name -replace '[\\/:\*\?"<>\|]', '_') }

# ==============================================================
#  EXPORT  (source domain)
# ==============================================================
function Invoke-Export {
    Write-Log "===== EXPORT ($SourceDomain) $(if ($DryRun) {'- DRY RUN'}) ====="
    Ensure-Directory $BackupPath
    Ensure-Directory $ReportPath

    if (-not (Test-Path $InScopeCsv)) { throw "In-scope list not found: $InScopeCsv" }
    $rows = @(Import-Csv $InScopeCsv)
    if ($rows.Count -eq 0) { throw "In-scope list is empty: $InScopeCsv" }
    $cols = $rows[0].PSObject.Properties.Name
    if ($cols -notcontains $Columns.GPOName) { throw "In-scope list has no '$($Columns.GPOName)' column" }
    $hasScopeCol = $cols -contains $Columns.InScope
    $hasTargetCol = $cols -contains $Columns.TargetName
    if (-not $hasScopeCol) { Write-Log "No '$($Columns.InScope)' column - every row is treated as in scope" 'WARN' }

    # In-scope GPO names (trimmed, de-duplicated, case-insensitive)
    $wanted = [ordered]@{}
    foreach ($r in $rows) {
        $name = ([string]$r.($Columns.GPOName)).Trim()
        if (-not $name) { continue }
        if ($hasScopeCol -and ([string]$r.($Columns.InScope)).Trim() -ne $InScopeValue) { continue }
        $override = if ($hasTargetCol) { [string]$r.($Columns.TargetName) } else { '' }
        if (-not $wanted.Contains($name.ToLower())) { $wanted[$name.ToLower()] = @{ Name = $name; Override = $override } }
    }
    Write-Log "In-scope GPOs in list: $($wanted.Count) (of $($rows.Count) rows)"
    if ($wanted.Count -eq 0) { Write-Log "Nothing marked '$InScopeValue' - nothing to export" 'WARN'; return }

    # Resolve GPOs
    $gpos = @{}          # guid(lower) -> @{Gpo; Override}
    $notFound = @()
    foreach ($w in $wanted.Values) {
        try {
            $g = Get-GPO -Name $w.Name -Domain $SourceDomain -ErrorAction Stop
            $gpos[$g.Id.ToString().ToLower()] = @{ Gpo = $g; Override = $w.Override }
        } catch {
            $notFound += [PSCustomObject]@{ GPOName = $w.Name; Error = $_.Exception.Message }
            Write-Log "Not found in $SourceDomain : '$($w.Name)'" 'WARN'
        }
    }
    $notFound | Export-Csv $NotFoundFile -NoTypeInformation -Encoding UTF8
    if ($gpos.Count -eq 0) { Write-Log "None of the in-scope GPOs were found" 'ERROR'; return }

    # ---- Links: every container whose gPLink names an in-scope GPO
    $domainDN = (Get-ADDomain -Server $SourceDomain).DistinguishedName
    $guidPattern = ($gpos.Keys | ForEach-Object { [regex]::Escape($_) }) -join '|'
    $links = New-Object System.Collections.Generic.List[object]

    $containers = @(Get-ADObject -Server $SourceDomain -SearchBase $domainDN -SearchScope Subtree `
                        -LDAPFilter '(gPLink=*)' -Properties gPLink, canonicalName)
    Write-Log "Containers with GPO links in $SourceDomain : $($containers.Count)"
    foreach ($c in $containers) {
        if ($c.gPLink -notmatch $guidPattern) { continue }
        $inh = Get-GPInheritance -Target $c.DistinguishedName -Domain $SourceDomain
        foreach ($l in $inh.GpoLinks) {
            $gid = $l.GpoId.ToString().ToLower()
            if (-not $gpos.ContainsKey($gid)) { continue }
            $links.Add([PSCustomObject]@{
                SourceGPOName     = $gpos[$gid].Gpo.DisplayName
                SourceGPOGuid     = $gid
                LinkType          = if ($c.DistinguishedName -eq $domainDN) { 'DomainRoot' } else { 'OU' }
                SourceOU          = $c.DistinguishedName
                SourceOUCanonical = $c.canonicalName
                LinkOrder         = $l.Order
                LinkEnabled       = [bool]$l.Enabled
                Enforced          = [bool]$l.Enforced
                BlockInheritance  = [bool]$inh.GpoInheritanceBlocked
            })
        }
    }

    # Site links (recorded, not migrated)
    try {
        $cfg = (Get-ADRootDSE -Server $SourceDomain).configurationNamingContext
        $sites = @(Get-ADObject -Server $SourceDomain -SearchBase "CN=Sites,$cfg" -LDAPFilter '(gPLink=*)' -Properties gPLink)
        foreach ($s in $sites) {
            foreach ($gid in $gpos.Keys) {
                if ($s.gPLink -match [regex]::Escape($gid)) {
                    $links.Add([PSCustomObject]@{
                        SourceGPOName = $gpos[$gid].Gpo.DisplayName; SourceGPOGuid = $gid; LinkType = 'Site'
                        SourceOU = $s.DistinguishedName; SourceOUCanonical = $s.Name; LinkOrder = ''
                        LinkEnabled = ''; Enforced = ''; BlockInheritance = ''
                    })
                }
            }
        }
    } catch { Write-Log "Could not read site links: $($_.Exception.Message)" 'WARN' }

    $links | Sort-Object SourceGPOName, SourceOU | Export-Csv $LinksFile -NoTypeInformation -Encoding UTF8
    Write-Log "Links found: $($links.Count) -> $LinksFile"

    # ---- Back up each GPO
    $manifest = @()
    foreach ($gid in $gpos.Keys) {
        $g = $gpos[$gid].Gpo
        $targetName = Get-TargetName $g.DisplayName $gpos[$gid].Override
        $myLinks = @($links | Where-Object { $_.SourceGPOGuid -eq $gid })

        # Security filtering (who the GPO applies to)
        # Deleted accounts have no name, so their SID is recorded instead; deny entries get "DENY:".
        $apply = @()
        try {
            $apply = @(Get-GPPermission -Guid $g.Id -All -Domain $SourceDomain -ErrorAction Stop |
                       Where-Object { $_.Permission -eq 'GpoApply' } |
                       ForEach-Object {
                           $t = $_.Trustee
                           $n = if ($t.Name) { if ($t.Domain) { "$($t.Domain)\$($t.Name)" } else { $t.Name } } else { [string]$t.Sid }
                           if ($_.Denied) { "DENY:$n" } else { $n }
                       })
        } catch { Write-Log "Could not read permissions for '$($g.DisplayName)': $($_.Exception.Message)" 'WARN' }
        $defaultFilter = -not (Test-NarrowFiltering ($apply -join '; '))
        $wmi = if ($g.WmiFilter) { $g.WmiFilter.Name } else { '' }

        $entry = [ordered]@{
            SourceName          = $g.DisplayName
            SourceGuid          = $gid
            TargetName          = $targetName
            BackupId            = ''
            GpoStatus           = [string]$g.GpoStatus
            WmiFilter           = $wmi
            SecurityFiltering   = ($apply -join '; ')
            NonDefaultFiltering = (-not $defaultFilter)
            OULinks             = @($myLinks | Where-Object { $_.LinkType -ne 'Site' }).Count
            SiteLinks           = @($myLinks | Where-Object { $_.LinkType -eq 'Site' }).Count
            Status              = ''
            Error               = ''
        }

        if ($DryRun) {
            $entry.Status = 'DryRun'
            Write-Log "WOULD back up '$($g.DisplayName)' -> '$targetName' ($($entry.OULinks) link(s))"
        } else {
            try {
                $b = Backup-GPO -Guid $g.Id -Path $BackupPath -Domain $SourceDomain `
                        -Comment "ENGR migration export $(Get-Date -Format 'yyyy-MM-dd')" -ErrorAction Stop
                $entry.BackupId = $b.Id.ToString()
                Get-GPOReport -Guid $g.Id -ReportType Html -Domain $SourceDomain `
                    -Path (Join-Path $ReportPath ((Get-SafeFileName $g.DisplayName) + '.html')) -ErrorAction SilentlyContinue
                $entry.Status = 'BackedUp'
                Write-Log "Backed up '$($g.DisplayName)' -> '$targetName' ($($entry.OULinks) link(s))" 'OK'
            } catch {
                $entry.Status = 'FAILED'
                $entry.Error = $_.Exception.Message
                Write-Log "FAILED to back up '$($g.DisplayName)': $($_.Exception.Message)" 'ERROR'
            }
        }
        if ($entry.NonDefaultFiltering) { Write-Log "  security filtering: $($entry.SecurityFiltering)" 'WARN' }
        if ($wmi) { Write-Log "  WMI filter: $wmi" 'WARN' }
        $manifest += [PSCustomObject]$entry
    }
    $manifest | Sort-Object SourceName | Export-Csv $ManifestFile -NoTypeInformation -Encoding UTF8

    # ---- Migration table template (principals + UNC paths referenced by the backups)
    if (-not $DryRun -and ($manifest | Where-Object Status -eq 'BackedUp')) {
        try {
            $gpm = New-Object -ComObject GPMgmt.GPM
            $k = $gpm.GetConstants()
            $mt = $gpm.CreateMigrationTable()
            $dir = $gpm.GetBackupDir($BackupPath)
            $sc = $gpm.CreateSearchCriteria()
            $sc.Add($k.SearchPropertyBackupMostRecent, $k.SearchOpEquals, $true)
            foreach ($b in $dir.SearchBackups($sc)) {
                if ($manifest.BackupId -contains $b.ID.Trim('{}')) { [void]$mt.Add($k.ProcessSecurity, $b) }
            }
            $mt.Save($MigTableTemplate)
            Write-Log "Migration table template: $MigTableTemplate" 'OK'
        } catch {
            Write-Log "Could not build migration table template (build it in GPMC instead): $($_.Exception.Message)" 'WARN'
        }
    }

    $ok = @($manifest | Where-Object Status -in 'BackedUp', 'DryRun').Count
    Write-Log "===== EXPORT DONE: $ok GPO(s), $($links.Count) link(s), $($notFound.Count) not found ====="
    Write-Log "Review: $ManifestFile, $LinksFile$(if ($notFound) {", $NotFoundFile"})"
    if (-not $DryRun) {
        Write-Log "Next: edit $MigTableTemplate in GPMC Migration Table Editor, save it as"
        Write-Log "      $MigTable, then copy $WorkRoot to the target side."
    }
}

# ==============================================================
#  IMPORT  (target domain)
# ==============================================================
function Invoke-Import {
    Write-Log "===== IMPORT ($TargetDomain) $(if ($DryRun) {'- DRY RUN'}) ====="
    foreach ($f in $ManifestFile, $LinksFile, $OUMapCsv) {
        if (-not (Test-Path $f)) { throw "Required file not found: $f" }
    }
    $manifest = @(Import-Csv $ManifestFile | Where-Object { $_.Status -eq 'BackedUp' })
    $links    = @(Import-Csv $LinksFile)
    Write-Log "GPOs to import: $($manifest.Count); source links: $($links.Count)"
    if ($manifest.Count -eq 0) { Write-Log "No backed-up GPOs in the manifest (was Export run without -DryRun?)" 'ERROR'; return }

    # OU map: source DN -> target DN
    $map = @{}
    $mapRows = @(Import-Csv $OUMapCsv)
    foreach ($c in $Columns.SourceOU, $Columns.TargetOU) {
        if ($mapRows.Count -and ($mapRows[0].PSObject.Properties.Name -notcontains $c)) { throw "OU map has no '$c' column" }
    }
    # A source DN on more than one row, or with conflicting targets, means the map was built
    # by OU name instead of full path (e.g. every "PI" OU resolved to the same DN). Those
    # entries are refused rather than risk linking to the wrong OU.
    $rowsPerSource = @{}
    foreach ($r in $mapRows) {
        $s = ([string]$r.($Columns.SourceOU)).Trim().ToLower()
        if ($s) { $rowsPerSource[$s] = 1 + [int]$rowsPerSource[$s] }
    }
    $ambiguous = @{}
    foreach ($r in $mapRows) {
        $s = ([string]$r.($Columns.SourceOU)).Trim()
        $t = ([string]$r.($Columns.TargetOU)).Trim()
        if (-not $s -or $t -notmatch '^(OU|DC)=') { continue }
        $k = $s.ToLower()
        if ($rowsPerSource[$k] -gt 1 -or ($map.ContainsKey($k) -and $map[$k] -ne $t)) { $ambiguous[$k] = $true; continue }
        $map[$k] = $t
    }
    foreach ($k in $ambiguous.Keys) { $map.Remove($k) }
    Write-Log "OU map: $($map.Count) usable source OU(s)"
    if ($ambiguous.Count) {
        Write-Log "OU map: $($ambiguous.Count) source OU(s) refused - listed on more than one row (map looks built by OU name, not full path)" 'WARN'
    }

    $useMigTable = Test-Path $MigTable
    if ($useMigTable) { Write-Log "Using migration table: $MigTable" }
    else { Write-Log "No $MigTable - principals and UNC paths import unchanged" 'WARN' }

    # GPOs with at least one link the OU map can place
    $mappedLinks = @{}
    foreach ($l in $links) {
        if ($l.LinkType -ne 'Site' -and $map.ContainsKey(([string]$l.SourceOU).ToLower())) {
            $mappedLinks[$l.SourceGPOGuid] = 1 + [int]$mappedLinks[$l.SourceGPOGuid]
        }
    }
    if ($OnlyLinked) {
        Write-Log "-OnlyLinked: $(@($manifest | Where-Object { $mappedLinks.ContainsKey($_.SourceGuid) }).Count) of $($manifest.Count) GPO(s) have a mapped link; the rest are skipped this run"
    }

    # ---- GPOs
    $results = @()
    $targetGuid = @{}     # source guid -> target guid (string), or 'DRYRUN'
    $holdDisabled = @{}   # source guid -> reason links are created disabled
    foreach ($e in $manifest) {
        $res = [ordered]@{ SourceName = $e.SourceName; TargetName = $e.TargetName; TargetGuid = ''; Action = ''; Note = ''; Error = '' }
        $reasons = @()
        if (Test-NarrowFiltering $e.SecurityFiltering) { $reasons += "security filtering ($($e.SecurityFiltering))" }
        if ($e.WmiFilter) { $reasons += "WMI filter ($($e.WmiFilter))" }
        if ($reasons) { $holdDisabled[$e.SourceGuid] = ($reasons -join ', ') }

        try {
            $existing = Get-GPO -Name $e.TargetName -Domain $TargetDomain -ErrorAction SilentlyContinue
            $ours = $existing -and ([string]$existing.Description).Contains("$MarkerPrefix {$($e.SourceGuid)}")

            if ($OnlyLinked -and -not $mappedLinks.ContainsKey($e.SourceGuid) -and -not $existing) {
                $res.Action = 'SKIPPED-NoMappedLinks'
                $res.Note = 'No link maps to a target OU yet - re-run after the OU map is filled in'
            } elseif ($existing -and -not $ours) {
                $res.Action = 'SKIPPED-NameInUse'
                $res.Note = "A GPO named '$($e.TargetName)' already exists in $TargetDomain and wasn't created by this process"
                Write-Log "SKIPPED '$($e.TargetName)': name already used by a GPO this process didn't create" 'ERROR'
            } elseif ($existing -and -not $OverwriteExisting) {
                $res.Action = 'Exists-NotChanged'
                $res.TargetGuid = $existing.Id.ToString()
                $targetGuid[$e.SourceGuid] = $res.TargetGuid
                Write-Log "Exists, settings left as-is: '$($e.TargetName)' (use -OverwriteExisting to re-import)"
            } elseif ($DryRun) {
                $res.Action = if ($existing) { 'WOULD-Overwrite' } else { 'WOULD-Create' }
                $targetGuid[$e.SourceGuid] = if ($existing) { $existing.Id.ToString() } else { 'DRYRUN' }
                Write-Log "$($res.Action) '$($e.TargetName)' from '$($e.SourceName)'"
            } else {
                $p = @{ BackupId = $e.BackupId; Path = $BackupPath; Domain = $TargetDomain; ErrorAction = 'Stop' }
                if ($useMigTable) { $p.MigrationTable = $MigTable }
                if ($existing) { $p.TargetGuid = $existing.Id } else { $p.TargetName = $e.TargetName; $p.CreateIfNeeded = $true }
                $g = Import-GPO @p
                $marker = "$MarkerPrefix {$($e.SourceGuid)}]"
                if (-not ([string]$g.Description).Contains($marker)) {
                    $g.Description = (([string]$g.Description).Trim() + " $marker").Trim()
                }
                if ($e.GpoStatus -and [string]$g.GpoStatus -ne $e.GpoStatus) {
                    try { $g.GpoStatus = $e.GpoStatus } catch { $res.Note = "Could not set GpoStatus to $($e.GpoStatus)" }
                }
                $res.Action = if ($existing) { 'Overwritten' } else { 'Created' }
                $res.TargetGuid = $g.Id.ToString()
                $targetGuid[$e.SourceGuid] = $res.TargetGuid
                Write-Log "$($res.Action) '$($e.TargetName)'" 'OK'
            }
        } catch {
            $res.Action = 'FAILED'
            $res.Error = $_.Exception.Message
            Write-Log "FAILED '$($e.TargetName)': $($_.Exception.Message)" 'ERROR'
        }
        if ($holdDisabled.ContainsKey($e.SourceGuid)) {
            $res.Note = (($res.Note, "links created disabled: $($holdDisabled[$e.SourceGuid])") | Where-Object { $_ }) -join '; '
        }
        $results += [PSCustomObject]$res
    }
    $results | Export-Csv $ImportResults -NoTypeInformation -Encoding UTF8

    # ---- Links
    $nameBySource = @{}; foreach ($e in $manifest) { $nameBySource[$e.SourceGuid] = $e.TargetName }
    $ouExists = @{}
    $linkOut = @()
    $planned = @{}        # "targetOU|guid" already handled this run

    # Per target OU, in source link order, so migrated links keep their relative order
    $work = foreach ($l in $links) {
        if (-not $nameBySource.ContainsKey($l.SourceGPOGuid)) { continue }
        $tou = if ($l.LinkType -eq 'Site') { $null } else { $map[([string]$l.SourceOU).ToLower()] }
        [PSCustomObject]@{ Link = $l; TargetOU = $tou; Order = [int]('0' + $l.LinkOrder) }
    }
    foreach ($w in ($work | Sort-Object TargetOU, Order)) {
        $l = $w.Link
        $o = [ordered]@{
            TargetGPOName = $nameBySource[$l.SourceGPOGuid]; SourceOU = $l.SourceOU; TargetOU = $w.TargetOU
            SourceOrder = $l.LinkOrder; LinkEnabled = ''; Enforced = ''; Action = ''; Note = ''
        }
        $tg = $targetGuid[$l.SourceGPOGuid]

        if ($l.LinkType -eq 'Site') { $o.Action = 'SKIPPED-SiteLink'; $o.Note = 'Site links are not migrated' }
        elseif (-not $w.TargetOU -and $ambiguous.ContainsKey(([string]$l.SourceOU).ToLower())) {
            $o.Action = 'SKIPPED-AmbiguousMap'; $o.Note = 'Source OU is on more than one row of the OU map - fix the map by full path'
        }
        elseif (-not $w.TargetOU) { $o.Action = 'SKIPPED-Unmapped'; $o.Note = 'No target OU in the OU map for this source OU' }
        elseif (-not $tg) { $o.Action = 'SKIPPED-GPONotImported' }
        else {
            $key = "$($w.TargetOU.ToLower())|$($l.SourceGPOGuid)"
            if ($planned.ContainsKey($key)) {
                $o.Action = 'SKIPPED-Duplicate'; $o.Note = 'Another source OU maps to the same target OU'
            } else {
                $planned[$key] = $true
                if (-not $ouExists.ContainsKey($w.TargetOU)) {
                    try { [void](Get-ADOrganizationalUnit -Identity $w.TargetOU -Server $TargetDomain -ErrorAction Stop); $ouExists[$w.TargetOU] = $true }
                    catch { $ouExists[$w.TargetOU] = $false }
                }
                $enabled = ($l.LinkEnabled -eq 'True')
                if ($LinkState -eq 'Disabled') { $enabled = $false }
                if ($holdDisabled.ContainsKey($l.SourceGPOGuid)) { $enabled = $false; $o.Note = "disabled: $($holdDisabled[$l.SourceGPOGuid])" }
                $o.LinkEnabled = if ($enabled) { 'Yes' } else { 'No' }
                $o.Enforced    = if ($l.Enforced -eq 'True') { 'Yes' } else { 'No' }
                if ($l.BlockInheritance -eq 'True') {
                    $o.Note = (($o.Note, 'source OU blocks inheritance - check the target OU') | Where-Object { $_ }) -join '; '
                }

                if (-not $ouExists[$w.TargetOU]) {
                    $o.Action = 'SKIPPED-TargetOUNotFound'
                } elseif ($tg -eq 'DRYRUN') {
                    $o.Action = 'WOULD-Link'
                } else {
                    try {
                        $inh = Get-GPInheritance -Target $w.TargetOU -Domain $TargetDomain -ErrorAction Stop
                        if (@($inh.GpoLinks) | Where-Object { $_ -and $_.GpoId -and $_.GpoId.ToString() -eq $tg }) {
                            $o.Action = 'AlreadyLinked-NotChanged'
                        } elseif ($DryRun) {
                            $o.Action = 'WOULD-Link'
                        } else {
                            [void](New-GPLink -Guid $tg -Target $w.TargetOU -Domain $TargetDomain `
                                    -LinkEnabled $o.LinkEnabled -Enforced $o.Enforced -ErrorAction Stop)
                            $o.Action = 'Linked'
                        }
                    } catch {
                        $o.Action = 'FAILED'; $o.Note = $_.Exception.Message
                    }
                }
            }
        }
        $linkOut += [PSCustomObject]$o
    }
    $linkOut | Export-Csv $LinkResults -NoTypeInformation -Encoding UTF8

    # ---- Summary
    Write-Log "===== IMPORT DONE ====="
    $results | Group-Object Action | Sort-Object Name | ForEach-Object { Write-Log ("  GPOs  {0,-26} {1}" -f $_.Name, $_.Count) }
    $linkOut | Group-Object Action | Sort-Object Name | ForEach-Object { Write-Log ("  Links {0,-26} {1}" -f $_.Name, $_.Count) }
    $held = @($linkOut | Where-Object { $_.LinkEnabled -eq 'No' -and $_.Action -in 'Linked', 'WOULD-Link' }).Count
    if ($held) { Write-Log "  $held link(s) created disabled - see Note in $LinkResults" 'WARN' }
    Write-Log "Results: $ImportResults, $LinkResults"
}

# ==============================================================
#  MAIN
# ==============================================================
Ensure-Directory $WorkRoot
try {
    Import-RequiredModules
    if ($Mode -eq 'Export') { Invoke-Export } else { Invoke-Import }
} catch {
    Write-Log $_.Exception.Message 'ERROR'
    exit 1
}

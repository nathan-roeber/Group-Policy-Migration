<#
.SYNOPSIS
    Inventories all GPOs in a source domain and exports a CSV for Excel-based migration scoping.

.DESCRIPTION
    For each GPO in $SourceDomain, pulls identity metadata, link information, a light
    settings-presence scan, and translation-risk flags (NETBIOS refs, UNC paths).

    Output CSV has auto-populated technical columns plus blank staff columns:
        InScope | TargetGPOName | Priority | ReviewNotes | Reviewer

    GPOs are sorted: linked first (descending link count), then alphabetically.
    Unlinked GPOs sort to the bottom -- likely out-of-scope candidates.

.PARAMETER SourceDomain
    FQDN of the source domain. Default: ad.engr.wisc.edu

.PARAMETER GPOPrefix
    Prefix prepended to GPO DisplayName in the SuggestedTargetName column.
    Default: ENGR-ME_

.PARAMETER OutputPath
    Full path for the output CSV. Defaults to the script directory with a timestamp.

.EXAMPLE
    .\GPO-Inventory.ps1
    .\GPO-Inventory.ps1 -SourceDomain ad.engr.wisc.edu -GPOPrefix "ENGR-ME_"
    .\GPO-Inventory.ps1 -Verbose   # shows per-GPO progress in the console

.NOTES
    Requires: GroupPolicy PowerShell module (RSAT)
    Run from a machine with domain connectivity to $SourceDomain.
    ASCII-only source -- no smart quotes or em dashes.
#>

#Requires -Modules GroupPolicy

[CmdletBinding()]
param(
    [string]$SourceDomain = 'ad.engr.wisc.edu',
    [string]$GPOPrefix    = 'ENGR-ME_',
    [string]$OutputPath   = "$PSScriptRoot\GPO-Inventory_$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
)

$ErrorActionPreference = 'Stop'

# --------------------------------------------------------------------------
# Helper: returns 'Y' if $Pattern matches anywhere in $Xml, else empty string.
# Single-quoted PS strings have no escape sequences, so \\\\ = 4 backslashes
# as a regex pattern = matches two consecutive backslashes (UNC paths, etc.).
# --------------------------------------------------------------------------
function Get-XmlFlag {
    param([string]$Xml, [string]$Pattern)
    if ([string]::IsNullOrEmpty($Xml)) { return '' }
    if ($Xml -match $Pattern) { 'Y' } else { '' }
}

# --------------------------------------------------------------------------
# Retrieve GPO list
# --------------------------------------------------------------------------
Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Connecting to $SourceDomain ..." -ForegroundColor Cyan

try {
    $allGPOs = Get-GPO -All -Domain $SourceDomain
} catch {
    Write-Error "Failed to retrieve GPOs from ${SourceDomain}.`n$_"
    exit 1
}

Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Found $($allGPOs.Count) GPOs. Building inventory..." -ForegroundColor Cyan

$total   = $allGPOs.Count
$counter = 0

$results = foreach ($gpo in ($allGPOs | Sort-Object DisplayName)) {

    $counter++
    Write-Progress -Activity "GPO Inventory" `
                   -Status  "$counter / $total : $($gpo.DisplayName)" `
                   -PercentComplete ([math]::Round(($counter / $total) * 100))

    Write-Verbose "  Processing: $($gpo.DisplayName)"

    # ------------------------------------------------------------------
    # Pull full XML report for this GPO
    # ------------------------------------------------------------------
    $rawXml = ''
    try {
        $rawXml = Get-GPOReport -Guid $gpo.Id -Domain $SourceDomain -ReportType Xml
    } catch {
        Write-Warning "  XML report failed for '$($gpo.DisplayName)': $_"
    }

    # ------------------------------------------------------------------
    # Links
    # <SOMPath> elements in the report XML list every OU/domain the GPO
    # is linked to.  Using regex avoids namespace manager complexity.
    # ------------------------------------------------------------------
    $linkPaths = @()
    if ($rawXml) {
        $linkMatches = [regex]::Matches($rawXml, '<SOMPath>([^<]+)</SOMPath>')
        $linkPaths   = @($linkMatches | ForEach-Object { $_.Groups[1].Value })
    }

    # ------------------------------------------------------------------
    # Settings presence flags  (Y = detected, '' = not detected)
    # These are rough indicators based on XML element presence -- not a
    # full parse.  Good enough for scoping; verify with GPMC for edge cases.
    # ------------------------------------------------------------------
    $hasScripts        = Get-XmlFlag $rawXml '<Scripts>'
    $hasSoftwareInst   = Get-XmlFlag $rawXml 'SoftwareInstallation|<SoftwarePkg'
    $hasDriveMaps      = Get-XmlFlag $rawXml 'DriveMapSettings|<Drive '
    $hasFolderRedir    = Get-XmlFlag $rawXml 'FolderRedirectionSettings|<Folder '
    $hasRegistryPrefs  = Get-XmlFlag $rawXml '<Registry '
    $hasFirewall       = Get-XmlFlag $rawXml 'WindowsFirewall|FirewallSetting'
    $hasAuditPolicy    = Get-XmlFlag $rawXml 'AuditSetting|<AuditPolicy'
    $hasInternetZones  = Get-XmlFlag $rawXml 'InternetSettings|IESettings'
    $hasEnvVars        = Get-XmlFlag $rawXml '<EnvironmentVariable'
    $hasFileSystem     = Get-XmlFlag $rawXml '<FileSystem|<FileSecurity'

    # ------------------------------------------------------------------
    # Translation-risk flags
    # HasNetbiosRefs : ENGR\ references that need AD\ substitution
    # HasUncRefs     : \\ paths that may reference source-domain servers
    #                  Pattern \\\\ (regex) matches \\ in text content
    # ------------------------------------------------------------------
    $hasNetbiosRefs = Get-XmlFlag $rawXml 'ENGR\\'
    $hasUncRefs     = Get-XmlFlag $rawXml '\\\\'

    # ------------------------------------------------------------------
    # WMI filter
    # ------------------------------------------------------------------
    $wmiFilter = if ($null -ne $gpo.WmiFilter) { $gpo.WmiFilter.Name } else { '' }

    # ------------------------------------------------------------------
    # Complexity estimate -- count how many setting categories are flagged
    # ------------------------------------------------------------------
    $flagCount = @(
        $hasScripts, $hasSoftwareInst, $hasDriveMaps, $hasFolderRedir,
        $hasRegistryPrefs, $hasFirewall, $hasAuditPolicy, $hasInternetZones,
        $hasEnvVars, $hasFileSystem
    ) | Where-Object { $_ -eq 'Y' } | Measure-Object | Select-Object -ExpandProperty Count

    $complexity = if ([string]::IsNullOrEmpty($rawXml)) {
        'Unknown'
    } elseif ($flagCount -eq 0) {
        'Low'
    } elseif ($flagCount -le 2) {
        'Low'
    } elseif ($flagCount -le 5) {
        'Medium'
    } else {
        'High'
    }

    # ------------------------------------------------------------------
    # Description: flatten newlines so CSV stays clean
    # ------------------------------------------------------------------
    $desc = if ($gpo.Description) {
        ($gpo.Description -replace "`r`n|`r|`n", ' ').Trim()
    } else { '' }

    # ------------------------------------------------------------------
    # Output object
    # ------------------------------------------------------------------
    [PSCustomObject]@{

        # -- Identity --
        GPOName             = $gpo.DisplayName
        GUID                = $gpo.Id.ToString()
        GPOStatus           = $gpo.GpoStatus          # AllSettingsEnabled | ComputerSettingsDisabled | UserSettingsDisabled | AllSettingsDisabled
        Description         = $desc
        CreatedDate         = $gpo.CreationTime.ToString('yyyy-MM-dd')
        ModifiedDate        = $gpo.ModificationTime.ToString('yyyy-MM-dd')

        # -- Links --
        LinkCount           = $linkPaths.Count
        LinkedOUs           = ($linkPaths -join ' | ')

        # -- Settings presence --
        HasScripts          = $hasScripts
        HasSoftwareInst     = $hasSoftwareInst
        HasDriveMaps        = $hasDriveMaps
        HasFolderRedir      = $hasFolderRedir
        HasRegistryPrefs    = $hasRegistryPrefs
        HasFirewall         = $hasFirewall
        HasAuditPolicy      = $hasAuditPolicy
        HasInternetZones    = $hasInternetZones
        HasEnvVars          = $hasEnvVars
        HasFileSystem       = $hasFileSystem
        WmiFilter           = $wmiFilter

        # -- Translation risk --
        HasNetbiosRefs      = $hasNetbiosRefs   # ENGR\ -> needs AD\ swap
        HasUncRefs          = $hasUncRefs        # \\ paths -> verify server names

        # -- Migration guidance --
        Complexity          = $complexity
        SuggestedTargetName = "$GPOPrefix$($gpo.DisplayName)"

        # ---- Staff columns: fill in Excel ----
        InScope             = ''    # Y / N / TBD
        TargetGPOName       = ''    # leave blank to use SuggestedTargetName
        Priority            = ''    # High / Medium / Low
        ReviewNotes         = ''
        Reviewer            = ''
    }
}

Write-Progress -Activity "GPO Inventory" -Completed

# --------------------------------------------------------------------------
# Sort: linked GPOs first (most likely in-scope), unlinked at the bottom
# --------------------------------------------------------------------------
$sorted = $results | Sort-Object @{ Expression = { $_.LinkCount }; Descending = $true }, GPOName

# --------------------------------------------------------------------------
# Export
# --------------------------------------------------------------------------
$sorted | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

# --------------------------------------------------------------------------
# Summary
# --------------------------------------------------------------------------
$linked      = ($sorted | Where-Object { $_.LinkCount -gt 0 }).Count
$unlinked    = ($sorted | Where-Object { $_.LinkCount -eq 0 }).Count
$riskCount   = ($sorted | Where-Object { $_.HasNetbiosRefs -eq 'Y' -or $_.HasUncRefs -eq 'Y' -or $_.WmiFilter }).Count
$wmiCount    = ($sorted | Where-Object { $_.WmiFilter }).Count
$unknownXml  = ($sorted | Where-Object { $_.Complexity -eq 'Unknown' }).Count

Write-Host ""
Write-Host "Inventory complete." -ForegroundColor Green
Write-Host "  Total GPOs        : $($sorted.Count)"
Write-Host "  Linked            : $linked"
Write-Host "  Unlinked          : $unlinked  (bottom of CSV -- likely out-of-scope)"
Write-Host "  Translation risk  : $riskCount  (NetBIOS refs, UNC refs, or WMI filter)"
Write-Host "  WMI filters       : $wmiCount  (manual review required -- won't migrate automatically)"
if ($unknownXml -gt 0) {
    Write-Host "  XML report errors : $unknownXml  (check warnings above)" -ForegroundColor Yellow
}
Write-Host ""
Write-Host "  Output: $OutputPath" -ForegroundColor Cyan

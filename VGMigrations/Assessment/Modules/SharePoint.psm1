#Requires -Version 7.0

Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -DisableNameChecking -Force -Global

<#
.SYNOPSIS
    Module for SharePoint/OneDrive collection during Discovery, plus the standalone cache
    builder used by Update-SharePointSitesCache.ps1.
.DESCRIPTION
    Get-SPOSite -Limit All against the whole tenant (and again with -IncludePersonalSite for
    OneDrive) is the slow part of SharePoint collection, even though the tenant's site list
    barely changes day to day. That walk now lives only in Update-SharePointSitesCacheFile,
    run standalone via Update-SharePointSitesCache.ps1 (or on demand from the app's Discovery
    Cache screen) - Invoke-SharePointCollection (the Discovery-time path) only ever reads the
    cached SharePointSites.json and filters it locally, so it no longer needs an active
    SharePoint Online connection or even the SPO module installed. Run-Assessment.ps1 checks
    the cache is present and fresh (Test-DiscoveryCachePrereqs) before this ever runs.
#>

# -----------------------------------------------------------------------
# Private functions
# -----------------------------------------------------------------------

<#
.SYNOPSIS
    Builds a site output record with null guards on every property due to -UseWindowsPowerShell deserialization.
#>
function Build-SPOSiteRecord {
    param([object]$Site)

    $groupId = $null
    if ($Site.GroupId) {
        $gStr = $Site.GroupId.ToString()
        if ($gStr -ne '00000000-0000-0000-0000-000000000000') { $groupId = $gStr }
    }

    $siteId  = if ($Site.Id)                  { $Site.Id.ToString() }                                                    else { '' }
    $usedGB  = if ($Site.StorageUsageCurrent)  { Convert-SizeToGB -Value $Site.StorageUsageCurrent -FromMB }              else { [double]0 }
    $quotaGB = if ($Site.StorageQuota)         { Convert-SizeToGB -Value $Site.StorageQuota -FromMB }                     else { [double]0 }
    $usedPct = if ($Site.StorageQuota -gt 0)   { [math]::Round($Site.StorageUsageCurrent / $Site.StorageQuota * 100, 1) } else { [double]0 }
    $owner   = if ($Site.Owner)                { $Site.Owner }                                                            else { '' }
    $lastMod = if ($Site.LastContentModifiedDate) { $Site.LastContentModifiedDate }                                       else { $null }
    $status  = if ($Site.Status)               { "" + $Site.Status }                                                      else { '' }

    $isChannelSite = if ($Site.IsTeamsChannelConnected) { [bool]$Site.IsTeamsChannelConnected } else { $false }

    # Channel sites carry the parent team's group ID in RelatedGroupId, not GroupId
    $relatedGroupId = $null
    if ($Site.RelatedGroupId) {
        $rStr = $Site.RelatedGroupId.ToString()
        if ($rStr -ne '00000000-0000-0000-0000-000000000000') { $relatedGroupId = $rStr }
    }

    [PSCustomObject]@{
        Title                   = $Site.Title
        Url                     = $Site.Url
        Template                = $Site.Template
        SiteId                  = $siteId
        GroupId                 = $groupId
        RelatedGroupId          = $relatedGroupId
        IsTeamsChannelConnected = $isChannelSite
        StorageUsedGB           = $usedGB
        StorageQuotaGB          = $quotaGB
        StorageUsedPercent      = $usedPct
        SiteOwner               = $owner
        LastModified            = $lastMod
        Status                  = $status
        IsTeamsConnected        = $null
        AssociatedObject        = $null
        MigrationObjectType     = $null
        # Initialised here so the property exists before Workbook.psm1 assigns it
        IsOrphaned              = $false
    }
}

<#
.SYNOPSIS
    Filters the cached tenant-wide site list to the VBU's non-personal sites.
#>
function Get-SPOSiteData {
    param([PSCustomObject]$Context, [Parameter(Mandatory)][string]$CacheFolder)

    $searchTerm = $Context.VBUSearchTerm
    $domain     = $Context.VBUDomain
    $allSites   = Import-AssessmentJson -FileName 'SharePointSites.json' -RawPath $CacheFolder
    $sites      = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($s in ($allSites | Where-Object { $_.Url -notlike '*/personal/*' })) {
        if (($s.Url   -like "*$searchTerm*") -or
            ($s.Url   -like "*$domain*")     -or
            ($s.Title -like "*$searchTerm*")) {
            $sites.Add($s)
        }
    }

    return $sites.ToArray()
}

<#
.SYNOPSIS
    Filters the cached tenant-wide site list to personal (OneDrive) sites owned by a VBU user.
#>
function Get-OneDriveData {
    param([PSCustomObject]$Context, [Parameter(Mandatory)][string]$CacheFolder)

    $adUsersPath = Join-Path $Context.RawPath 'ADUsers.json'
    if (-not (Test-Path $adUsersPath)) {
        Write-Host ($PREFIX_WARN + 'ADUsers.json not found - OneDrive cross-reference skipped') -ForegroundColor Yellow
        return @()
    }

    $adUsers   = Get-Content $adUsersPath -Raw | ConvertFrom-Json
    $upnToUser = @{}
    foreach ($u in $adUsers) {
        if ($u.UserPrincipalName) { $upnToUser[$u.UserPrincipalName.ToLower()] = $u }
    }

    $allSites = Import-AssessmentJson -FileName 'SharePointSites.json' -RawPath $CacheFolder
    $records  = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($od in ($allSites | Where-Object { $_.Url -like '*/personal/*' })) {
        # Strip claims provider prefix: i:0#.f|membership|user@domain.com -> user@domain.com
        $ownerLogin = $od.SiteOwner
        if ($ownerLogin -like '*|*') { $ownerLogin = $ownerLogin.Split('|')[-1] }

        if (-not $ownerLogin -or -not $upnToUser.ContainsKey($ownerLogin.ToLower())) { continue }

        $adUser   = $upnToUser[$ownerLogin.ToLower()]
        $remainGB = [math]::Round($od.StorageQuotaGB - $od.StorageUsedGB, 2)

        $records.Add([PSCustomObject]@{
            OwnerUPN           = $ownerLogin
            OwnerDisplayName   = $adUser.DisplayName
            OneDriveUrl        = $od.Url
            StorageUsedGB      = $od.StorageUsedGB
            StorageQuotaGB     = $od.StorageQuotaGB
            StorageRemainingGB = $remainGB
            StorageUsedPercent = $od.StorageUsedPercent
            LastModified       = $od.LastModified
            IsProvisioned      = ($od.Status -eq 'Active')
        })
    }

    return $records.ToArray()
}

# -----------------------------------------------------------------------
# Public functions
# -----------------------------------------------------------------------

<#
.SYNOPSIS
    Walks every site in the connected tenant (incl. OneDrive) and writes the cache file.
.DESCRIPTION
    Called by Update-SharePointSitesCache.ps1, which owns connecting to SharePoint Online first
    (Invoke-SharePointCollection - the Discovery-time path - no longer connects to SPO at all).
    A single -IncludePersonalSite $true -Limit All call covers both regular sites and OneDrives,
    so both Get-SPOSiteData and Get-OneDriveData can be served from the one cache file.
#>
function Update-SharePointSitesCacheFile {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$CacheFolder)

    Write-Host ($PREFIX_INFO + 'Enumerating all SharePoint sites (incl. OneDrive) tenant-wide...') -ForegroundColor DarkGray
    $allSites = @(Get-SPOSite -IncludePersonalSite $true -Limit All -ErrorAction Stop)
    $records  = @($allSites | ForEach-Object { Build-SPOSiteRecord -Site $_ })
    Write-ProgressLine -Label 'SharePoint + OneDrive sites (tenant-wide)' -Count $records.Count

    New-Item -ItemType Directory -Path $CacheFolder -Force | Out-Null
    $cachePath = Join-Path $CacheFolder 'SharePointSites.json'
    $records | ConvertTo-Json -Depth 10 | Set-Content -Path $cachePath -Encoding UTF8
    Write-Host ($PREFIX_OK + "SharePointSites.json cache written: $cachePath") -ForegroundColor Green

    return $records.Count
}

<#
.SYNOPSIS
    Orchestrates SharePoint site and OneDrive collection for a Discovery run, from cache only.
.DESCRIPTION
    Reads and locally filters the cached SharePointSites.json (see Update-SharePointSitesCacheFile)
    - no SharePoint Online connection is made here. Failures are non-critical and are returned as
    a failed collector result.
#>
function Invoke-SharePointCollection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][PSCustomObject]$Context,
        [Parameter(Mandatory)][string]$CacheFolder
    )

    $start = Get-Date
    Write-SectionHeader 'SharePoint Online'

    try {
        Write-Host ($PREFIX_INFO + 'Reading SharePoint sites from cache...') -ForegroundColor DarkGray
        $sites = Get-SPOSiteData -Context $Context -CacheFolder $CacheFolder
        Write-ProgressLine -Label 'SharePoint Sites' -Count $sites.Count

        Write-Host ($PREFIX_INFO + 'Reading OneDrive accounts from cache...') -ForegroundColor DarkGray
        $oneDrives = Get-OneDriveData -Context $Context -CacheFolder $CacheFolder
        Write-ProgressLine -Label 'OneDrive Accounts' -Count $oneDrives.Count

        Write-JsonOutput -FileName 'SharePointSites.json' -Data $sites     -RawPath $Context.RawPath
        Write-JsonOutput -FileName 'OneDrives.json'       -Data $oneDrives -RawPath $Context.RawPath

        $counts = @{
            SPOSiteCount  = $sites.Count
            OneDriveCount = $oneDrives.Count
        }
        $msg = "Sites: $($counts.SPOSiteCount), OneDrives: $($counts.OneDriveCount)"
        Update-CollectorStatus -CollectorName 'SharePoint Data' -Status 'Complete' `
            -RawPath $Context.RawPath -StartTime $start -Message $msg

        return New-CollectorResult -Success $true -Counts $counts
    }
    catch {
        Write-Host ($PREFIX_FAIL + 'SharePoint collection failed: ' + $_.Exception.Message) -ForegroundColor Red
        Update-CollectorStatus -CollectorName 'SharePoint Data' -Status 'Failed' `
            -RawPath $Context.RawPath -StartTime $start -Message $_.Exception.Message
        return New-CollectorResult -Success $false -ErrorMessage $_.Exception.Message
    }
}

# -----------------------------------------------------------------------
# Exports
# -----------------------------------------------------------------------

Export-ModuleMember -Function 'Invoke-SharePointCollection', 'Update-SharePointSitesCacheFile'

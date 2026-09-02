#Requires -Version 7.0
<#
.SYNOPSIS
    Update-DistributionGroupsCache.ps1 — Refreshes the tenant-wide distribution-group (+
    mail-enabled security group) cache the Assessment engine's Discovery run reads from.

.DESCRIPTION
    Get-DistributionGroup plus a per-group Get-DistributionGroupMember call used to run live,
    VBU-filtered, on every single VBU's Discovery run - correct, but slower than it needs to be
    on a tenant with many groups, and repeats the same tenant-wide-ish work once per VBU when a
    domain gets split into several. This script does the walk once, unfiltered (every group
    tenant-wide, not scoped to any one VBU), and writes the result to a cache file that every
    VBU's Discovery run then reads and filters locally instead - see
    Assessment\Modules\Exchange.psm1's Get-DistributionGroupData.

    Unlike the SharePoint Sites and Teams Channels caches, this one is a soft/optional speed-up,
    not a hard prerequisite: Run-Assessment.ps1 does NOT refuse to start without it. If this
    cache is missing, or older than 14 days, Get-DistributionGroupData falls straight back to the
    original live, VBU-filtered query - so Discovery keeps working exactly as before for anyone
    who hasn't refreshed this cache yet.

.PARAMETER SharePointAdminUrl
    Not used for sign-in here (this script connects to Exchange Online, not SharePoint Online) -
    only used as the cache key, so it must match what Run-Assessment.ps1 is given (or its own
    default) so the two agree on which cache folder to use. Same convention already used by
    Update-TeamsChannelsCache.ps1 for the identical reason.

.PARAMETER CacheRoot
    Base folder for cache files. Defaults to Assessment\Cache next to this script - same default
    Run-Assessment.ps1 uses.

.EXAMPLE
    .\Update-DistributionGroupsCache.ps1 -SharePointAdminUrl https://ourvolaris-admin.sharepoint.com
#>

param(
    [Parameter(Mandatory)]
    [string]$SharePointAdminUrl,

    [string]$CacheRoot
)

$moduleRoot = Join-Path $PSScriptRoot 'Assessment\Modules'
Import-Module (Join-Path $moduleRoot 'Common.psm1')   -Force -DisableNameChecking -Global
Import-Module (Join-Path $moduleRoot 'Exchange.psm1') -Force -DisableNameChecking -Global

$cacheRootPath = if ($CacheRoot) { $CacheRoot } else { Join-Path $PSScriptRoot 'Assessment\Cache' }
$cacheFolder   = Get-DiscoveryCacheFolder -SharePointAdminUrl $SharePointAdminUrl -CacheRoot $cacheRootPath
New-Item -ItemType Directory -Path $cacheFolder -Force | Out-Null

Write-Host ''
Write-Host 'Distribution Groups & Members Cache Refresh' -ForegroundColor Cyan
Write-Host ('=' * 40) -ForegroundColor Cyan
Write-Host ''

try { Get-ConnectionInformation -ErrorAction Stop | Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
Write-Host ($PREFIX_INFO + 'Connecting to Exchange Online...') -ForegroundColor DarkGray
Write-Host ($PREFIX_WARN + 'Your default browser will open - sign in as an Exchange administrator') -ForegroundColor Yellow
try {
    # -DisableWAM: same fix used throughout this toolkit - WAM needs a window handle this
    # headlessly-spawned process doesn't have, and device-code (its usual fallback) is blocked
    # tenant-wide by Conditional Access on baselined tenants anyway.
    Connect-ExchangeOnline -ShowBanner:$false -DisableWAM -ErrorAction Stop
    Write-Host ($PREFIX_OK + 'Connected') -ForegroundColor Green
}
catch {
    Write-Host ($PREFIX_FAIL + 'Exchange Online connection failed: ' + $_.Exception.Message) -ForegroundColor Red
    return
}

try {
    Update-DistributionGroupsCacheFile -CacheFolder $cacheFolder | Out-Null
    Write-Host ''
    Write-Host ($PREFIX_OK + "Cache refreshed: $cacheFolder") -ForegroundColor Green
}
catch {
    Write-Host ($PREFIX_FAIL + 'Cache refresh failed: ' + $_.Exception.Message) -ForegroundColor Red
}
finally {
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
}

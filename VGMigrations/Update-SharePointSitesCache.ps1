#Requires -Version 7.0
<#
.SYNOPSIS
    Update-SharePointSitesCache.ps1 — Refreshes the tenant-wide SharePoint sites cache the
    Assessment engine's Discovery run reads from.

.DESCRIPTION
    Get-SPOSite -Limit All against the whole source tenant (and again with -IncludePersonalSite
    for OneDrive) used to run live on every single VBU's Discovery run, even though the tenant's
    site list barely changes day to day. This script does that walk once and writes the result to
    a cache file that every VBU's Discovery run then reads and filters locally instead - see
    Assessment\Modules\SharePoint.psm1's Get-SPOSiteData / Get-OneDriveData.

    Run-Assessment.ps1 refuses to start if this cache is missing or older than 7 days for the
    given -SharePointAdminUrl - run this script first, or whenever the tenant's SharePoint sites
    may have changed materially.

.PARAMETER SharePointAdminUrl
    SPO admin URL for the source tenant, e.g. https://ourvolaris-admin.sharepoint.com. Also the
    cache key - must match what Run-Assessment.ps1 is given (or its own default) so the two agree
    on which cache folder to use.

.PARAMETER CacheRoot
    Base folder for cache files. Defaults to Assessment\Cache next to this script - same default
    Run-Assessment.ps1 uses.

.EXAMPLE
    .\Update-SharePointSitesCache.ps1 -SharePointAdminUrl https://ourvolaris-admin.sharepoint.com
#>

param(
    [Parameter(Mandatory)]
    [string]$SharePointAdminUrl,

    [string]$CacheRoot
)

$moduleRoot = Join-Path $PSScriptRoot 'Assessment\Modules'
Import-Module (Join-Path $moduleRoot 'Common.psm1')     -Force -DisableNameChecking -Global
Import-Module (Join-Path $moduleRoot 'SharePoint.psm1') -Force -DisableNameChecking -Global

$cacheRootPath = if ($CacheRoot) { $CacheRoot } else { Join-Path $PSScriptRoot 'Assessment\Cache' }
$cacheFolder   = Get-DiscoveryCacheFolder -SharePointAdminUrl $SharePointAdminUrl -CacheRoot $cacheRootPath
New-Item -ItemType Directory -Path $cacheFolder -Force | Out-Null

Write-Host ''
Write-Host 'SharePoint Sites Cache Refresh' -ForegroundColor Cyan
Write-Host ('=' * 40) -ForegroundColor Cyan
Write-Host ''

try {
    $prev = $WarningPreference
    $WarningPreference = 'SilentlyContinue'
    Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell -DisableNameChecking -ErrorAction Stop
    $WarningPreference = $prev
}
catch {
    $WarningPreference = $prev
    Write-Host ($PREFIX_FAIL + 'Microsoft.Online.SharePoint.PowerShell not available: ' + $_.Exception.Message) -ForegroundColor Red
    return
}

try { Disconnect-SPOService -ErrorAction SilentlyContinue } catch {}
Write-Host ($PREFIX_INFO + "Connecting to $SharePointAdminUrl...") -ForegroundColor DarkGray
Write-Host ($PREFIX_WARN + 'Your default browser will open - sign in as SharePoint Administrator') -ForegroundColor Yellow
try {
    # -UseSystemBrowser is required when spawned headlessly from Electron (as this always is) -
    # a plain Connect-SPOService popup has no window to open against and fails with "No valid
    # OAuth 2.0 authentication session exists" (confirmed live against Provision-OneDrives.ps1,
    # 2026-09-01, which shared this same connect pattern before this fix).
    Connect-SPOService -Url $SharePointAdminUrl -UseSystemBrowser:$true -ErrorAction Stop
    Write-Host ($PREFIX_OK + 'Connected') -ForegroundColor Green
}
catch {
    Write-Host ($PREFIX_FAIL + 'SPO connection failed: ' + $_.Exception.Message) -ForegroundColor Red
    return
}

try {
    Update-SharePointSitesCacheFile -CacheFolder $cacheFolder | Out-Null
    Write-Host ''
    Write-Host ($PREFIX_OK + "Cache refreshed: $cacheFolder") -ForegroundColor Green
}
catch {
    Write-Host ($PREFIX_FAIL + 'Cache refresh failed: ' + $_.Exception.Message) -ForegroundColor Red
}
finally {
    try { Disconnect-SPOService -ErrorAction SilentlyContinue } catch {}
}

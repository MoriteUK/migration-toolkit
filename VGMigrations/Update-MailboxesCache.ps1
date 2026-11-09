#Requires -Version 7.0
<#
.SYNOPSIS
    Update-MailboxesCache.ps1 — Refreshes the tenant-wide mailbox cache including statistics and permissions.

.DESCRIPTION
    Caches all mailboxes (User, Shared, Room, Equipment) with their statistics and permissions (FullAccess, SendAs).
    This allows Run-Assessment.ps1 to skip the Exchange connection entirely when the cache is fresh (< 7 days),
    making assessments much faster and avoiding Exchange API throttling.

    The cache includes:
    - All mailbox objects with properties (UPN, addresses, CustomAttributes, etc.)
    - Mailbox statistics for User and Shared mailboxes (item count, size, last logon)
    - FullAccess permissions (who can open the mailbox)
    - SendAs permissions (who can send as the mailbox)
    - SendOnBehalf is already in the mailbox object (GrantSendOnBehalfTo)

.PARAMETER SharePointAdminUrl
    Used as the cache key to match what Run-Assessment.ps1 uses, so both agree on which cache folder.

.PARAMETER CacheRoot
    Base folder for cache files. Defaults to Assessment\Cache next to this script.

.EXAMPLE
    .\Update-MailboxesCache.ps1 -SharePointAdminUrl https://ourvolaris-admin.sharepoint.com
#>

param(
    [Parameter(Mandatory)]
    [string]$SharePointAdminUrl,

    [string]$CacheRoot
)

$ErrorActionPreference = 'Stop'

$moduleRoot = Join-Path $PSScriptRoot 'Assessment\Modules'
Import-Module (Join-Path $moduleRoot 'Common.psm1')   -Force -DisableNameChecking -Global
Import-Module (Join-Path $moduleRoot 'Exchange.psm1') -Force -DisableNameChecking -Global

$cacheRootPath = if ($CacheRoot) { $CacheRoot } else { Join-Path $PSScriptRoot 'Assessment\Cache' }
$cacheFolder   = Get-DiscoveryCacheFolder -SharePointAdminUrl $SharePointAdminUrl -CacheRoot $cacheRootPath
New-Item -ItemType Directory -Path $cacheFolder -Force | Out-Null

Write-Host ''
Write-Host 'Exchange Mailboxes, Statistics & Permissions Cache Refresh' -ForegroundColor Cyan
Write-Host ('=' * 60) -ForegroundColor Cyan
Write-Host ''

# Disconnect any existing sessions
try { Get-ConnectionInformation -ErrorAction Stop | Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}

Write-Host ($PREFIX_INFO + 'Connecting to Exchange Online...') -ForegroundColor DarkGray
Write-Host ($PREFIX_WARN + 'Your default browser will open - sign in as an Exchange administrator') -ForegroundColor Yellow
try {
    Connect-ExchangeOnline -ShowBanner:$false -DisableWAM -ErrorAction Stop
    Write-Host ($PREFIX_OK + 'Connected') -ForegroundColor Green
}
catch {
    Write-Host ($PREFIX_FAIL + 'Exchange Online connection failed: ' + $_.Exception.Message) -ForegroundColor Red
    return
}

try {
    Update-MailboxesCacheFile -CacheFolder $cacheFolder | Out-Null
    Update-CacheCompletionStatus -CacheFolder $cacheFolder -CacheType 'Mailboxes'
    Write-Host ''
    Write-Host ($PREFIX_OK + "Cache refreshed: $cacheFolder") -ForegroundColor Green
    Write-Host ''
}
catch {
    Write-Host ''
    Write-Host ($PREFIX_FAIL + 'Cache refresh failed: ' + $_.Exception.Message) -ForegroundColor Red
    Write-Host ''
}
finally {
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
}

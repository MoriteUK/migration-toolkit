#Requires -Version 7.0
<#
.SYNOPSIS
    Update-TeamsChannelsCache.ps1 — Refreshes the tenant-wide Teams/channels/members cache the
    Assessment engine's Discovery run reads from.

.DESCRIPTION
    Walking every Team in the source tenant, every one of its channels (including private/shared
    channel membership), and every member used to run live on every single VBU's Discovery run -
    the slowest collector on a large tenant, which is why it used to be opt-out via
    -SkipTeamMemberships. This script does that walk once, unfiltered (every Team tenant-wide,
    not scoped to any one VBU), and writes the result to a cache file that every VBU's Discovery
    run then reads and matches against its own domain's users locally instead - see
    Assessment\Modules\TeamMemberships.psm1's Invoke-TeamMembershipCollection.

    Run-Assessment.ps1 refuses to start if this cache is missing or older than 7 days for the
    given -SharePointAdminUrl - run this script first, or whenever the tenant's Teams/channels may
    have changed materially.

.PARAMETER SharePointAdminUrl
    Not used for sign-in here (this script connects to Microsoft Graph, not SharePoint Online) -
    only used as the cache key, so it must match what Run-Assessment.ps1 is given (or its own
    default) so the two agree on which cache folder to use.

.PARAMETER CacheRoot
    Base folder for cache files. Defaults to Assessment\Cache next to this script - same default
    Run-Assessment.ps1 uses.

.EXAMPLE
    .\Update-TeamsChannelsCache.ps1 -SharePointAdminUrl https://ourvolaris-admin.sharepoint.com
#>

param(
    [Parameter(Mandatory)]
    [string]$SharePointAdminUrl,

    [string]$CacheRoot
)

$moduleRoot = Join-Path $PSScriptRoot 'Assessment\Modules'
Import-Module (Join-Path $moduleRoot 'Common.psm1')          -Force -DisableNameChecking -Global
Import-Module (Join-Path $moduleRoot 'TeamMemberships.psm1') -Force -DisableNameChecking -Global

$cacheRootPath = if ($CacheRoot) { $CacheRoot } else { Join-Path $PSScriptRoot 'Assessment\Cache' }
$cacheFolder   = Get-DiscoveryCacheFolder -SharePointAdminUrl $SharePointAdminUrl -CacheRoot $cacheRootPath
New-Item -ItemType Directory -Path $cacheFolder -Force | Out-Null

Write-Host ''
Write-Host 'Teams Channels & Members Cache Refresh' -ForegroundColor Cyan
Write-Host ('=' * 40) -ForegroundColor Cyan
Write-Host ''

# Pins Microsoft.Graph.* to 2.33.0 - the last version before WAM became mandatory for interactive
# sign-in (see Ensure-GraphModules.ps1's own header for the full history).
. (Join-Path $PSScriptRoot 'Ensure-GraphModules.ps1') -GraphModules @('Microsoft.Graph.Groups', 'Microsoft.Graph.Teams')
Import-Module Microsoft.Graph.Groups -DisableNameChecking -ErrorAction Stop
Import-Module Microsoft.Graph.Teams  -DisableNameChecking -ErrorAction Stop

try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
Write-Host ($PREFIX_INFO + 'Connecting to Microsoft Graph...') -ForegroundColor DarkGray
try {
    Connect-MgGraph -Scopes @(
        'Team.ReadBasic.All', 'Channel.ReadBasic.All', 'ChannelMember.Read.All',
        'TeamMember.Read.All', 'Group.Read.All'
    ) -NoWelcome -ErrorAction Stop
    Write-Host ($PREFIX_OK + 'Connected') -ForegroundColor Green
}
catch {
    Write-Host ($PREFIX_FAIL + 'Graph connection failed: ' + $_.Exception.Message) -ForegroundColor Red
    return
}

try {
    Update-TeamsChannelsCacheFile -CacheFolder $cacheFolder | Out-Null
    Write-Host ''
    Write-Host ($PREFIX_OK + "Cache refreshed: $cacheFolder") -ForegroundColor Green
}
catch {
    Write-Host ($PREFIX_FAIL + 'Cache refresh failed: ' + $_.Exception.Message) -ForegroundColor Red
}
finally {
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
}

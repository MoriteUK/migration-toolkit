#Requires -Version 7.0
<#
.SYNOPSIS
    Update-DomainTeamMembershipsCache.ps1 — Refreshes the cached domain Team memberships data.

.DESCRIPTION
    Scans ALL Teams in the tenant and captures which users from the specified domain belong to
    each Team and private/shared channel. The results are cached for 7 days to avoid re-scanning
    the entire tenant every time.

    The cache is stored in Assessment\Cache\<tenant-hash>\DomainTeamMemberships-<domain>.csv

    Get-DomainTeamMemberships.ps1 will automatically use this cache if it's less than 7 days old,
    and will only refresh if the cache is stale or missing.

.PARAMETER Domain
    The email domain to capture memberships for (e.g., contoso.com)

.PARAMETER SharePointAdminUrl
    Used as the cache key to identify the tenant. Must match what's used in other scripts.
    Optional - will prompt for connection if not provided.

.PARAMETER OutputCsv
    Path for the CSV output. If not specified, uses the cache folder.

.PARAMETER IncludeStandardChannels
    Also emit a row per standard channel for each matched user (normally redundant).

.PARAMETER CacheRoot
    Base folder for cache files. Defaults to Assessment\Cache next to this script.

.PARAMETER Force
    Force refresh even if cache is fresh.

.EXAMPLE
    .\Update-DomainTeamMembershipsCache.ps1 -Domain contoso.com -SharePointAdminUrl https://contoso-admin.sharepoint.com

.EXAMPLE
    .\Update-DomainTeamMembershipsCache.ps1 -Domain contoso.com -Force
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$Domain,

    [string]$SharePointAdminUrl = '',
    [string]$OutputCsv = '',
    [switch]$IncludeStandardChannels,
    [string]$CacheRoot = '',
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$Domain = $Domain.Trim().TrimStart('@').ToLowerInvariant()

# ── Determine cache location ─────────────────────────────────────────────
$cacheRootPath = if ($CacheRoot) { $CacheRoot } else { Join-Path $PSScriptRoot 'Assessment\Cache' }

# Use SharePointAdminUrl to create tenant-specific cache folder (same pattern as other caches)
if ($SharePointAdminUrl) {
    $moduleRoot = Join-Path $PSScriptRoot 'Assessment\Modules'
    if (Test-Path (Join-Path $moduleRoot 'Common.psm1')) {
        Import-Module (Join-Path $moduleRoot 'Common.psm1') -Force -DisableNameChecking -Global
        $cacheFolder = Get-DiscoveryCacheFolder -SharePointAdminUrl $SharePointAdminUrl -CacheRoot $cacheRootPath
    } else {
        # Fallback if Common.psm1 not available
        $hash = [System.BitConverter]::ToString([System.Security.Cryptography.SHA256]::Create().ComputeHash([System.Text.Encoding]::UTF8.GetBytes($SharePointAdminUrl.ToLowerInvariant()))).Replace('-','').Substring(0,16)
        $cacheFolder = Join-Path $cacheRootPath $hash
    }
} else {
    # No SharePointAdminUrl provided - use a generic cache folder
    $cacheFolder = Join-Path $cacheRootPath 'TeamMemberships'
}

New-Item -ItemType Directory -Path $cacheFolder -Force | Out-Null

$cacheFile = Join-Path $cacheFolder "DomainTeamMemberships-$Domain.csv"

# Check if cache is fresh (less than 7 days old)
if (-not $Force -and (Test-Path $cacheFile)) {
    $cacheAge = (Get-Date) - (Get-Item $cacheFile).LastWriteTime
    if ($cacheAge.TotalDays -lt 7) {
        Write-Host ""
        Write-Host "=== Cache is Fresh ===" -ForegroundColor Green
        Write-Host "Cache file: $cacheFile" -ForegroundColor Cyan
        Write-Host "Age: $([math]::Round($cacheAge.TotalDays, 1)) days old (refreshes every 7 days)" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Use -Force to refresh anyway, or wait until cache is older than 7 days." -ForegroundColor Yellow
        Write-Host ""
        return
    } else {
        Write-Host ""
        Write-Host "Cache is stale ($([math]::Round($cacheAge.TotalDays, 1)) days old) - refreshing..." -ForegroundColor Yellow
    }
}

# Set output location to cache file if not specified
if (-not $OutputCsv) {
    $OutputCsv = $cacheFile
}

Write-Host ""
Write-Host "=== Domain Team Memberships Cache Update ===" -ForegroundColor Cyan
Write-Host "Domain:      $Domain" -ForegroundColor Cyan
Write-Host "Cache file:  $cacheFile" -ForegroundColor Cyan
Write-Host "Cache root:  $cacheRootPath" -ForegroundColor DarkGray
Write-Host ""

# Call the main Get-DomainTeamMemberships.ps1 script with -OutputCsv pointing to cache
$scriptPath = Join-Path $PSScriptRoot 'Get-DomainTeamMemberships.ps1'
$args = @('-Domain', $Domain, '-OutputCsv', $OutputCsv)
if ($IncludeStandardChannels) { $args += '-IncludeStandardChannels' }

& $scriptPath @args

if (Test-Path $cacheFile) {
    $cacheAge = (Get-Date) - (Get-Item $cacheFile).LastWriteTime
    Write-Host ""
    Write-Host "=== Cache Updated ===" -ForegroundColor Green
    Write-Host "Location: $cacheFile" -ForegroundColor Cyan
    Write-Host "Valid for: 7 days (next refresh due: $((Get-Date).AddDays(7).ToString('yyyy-MM-dd HH:mm')))" -ForegroundColor Cyan
    Write-Host ""
}

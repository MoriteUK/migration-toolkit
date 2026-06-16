#Requires -Version 5.1
<#
.SYNOPSIS
    Permanently purges destination SharePoint sites (from a mapping CSV) that are
    sitting in the SPO recycle bin. Run this before recreating sites.
#>
param(
    [Parameter(Mandatory=$true)]
    [string]$MappingFile
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $MappingFile)) {
    Write-Error "Mapping file not found: $MappingFile"
    exit 1
}

Write-Host "Reading: $MappingFile" -ForegroundColor Cyan

$rawBytes = [System.IO.File]::ReadAllBytes($MappingFile)
$enc = if ($rawBytes.Length -ge 3 -and $rawBytes[0] -eq 0xEF -and $rawBytes[1] -eq 0xBB -and $rawBytes[2] -eq 0xBF) { 'UTF8' }
      elseif ($rawBytes.Length -ge 2 -and $rawBytes[0] -eq 0xFF -and $rawBytes[1] -eq 0xFE) { 'Unicode' }
      else { 'Default' }

$mappings = Import-Csv $MappingFile -Encoding $enc
Write-Host "Loaded $($mappings.Count) row(s)" -ForegroundColor Gray

$headers    = ($mappings | Select-Object -First 1).PSObject.Properties.Name
$destUrlCol = $headers | Where-Object { $_ -imatch '^destination' -and $_ -imatch 'url|site' } | Select-Object -First 1
if (-not $destUrlCol) { $destUrlCol = $headers | Where-Object { $_ -imatch '^destination' } | Select-Object -First 1 }
if (-not $destUrlCol) { Write-Error "Cannot find a Destination URL column."; exit 1 }

$destUrls = $mappings |
    Where-Object { $_.$destUrlCol -match 'https://' } |
    Select-Object -ExpandProperty $destUrlCol |
    ForEach-Object {
        $url = $_.Trim()
        if ($url -match '(https://[^/]+/(?:sites|teams)/[^/?#]+)') { $Matches[1] }
        elseif ($url -match 'https://') { $url }
    } | Sort-Object -Unique | Where-Object { $_ }

if (-not $destUrls) { Write-Error "No destination URLs found in mapping file."; exit 1 }

Write-Host ""
Write-Host "$($destUrls.Count) destination site(s) from mapping:" -ForegroundColor Cyan
$destUrls | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }

$uri          = [System.Uri]$destUrls[0]
$destAdminUrl = "$($uri.Scheme)://$($uri.Host -replace '\.sharepoint\.com$', '-admin.sharepoint.com')"
Write-Host ""
Write-Host "Destination admin URL: $destAdminUrl" -ForegroundColor Cyan

if (-not (Get-Module -Name Microsoft.Online.SharePoint.PowerShell -ListAvailable)) {
    Write-Error "Microsoft.Online.SharePoint.PowerShell is not installed."
    exit 1
}
Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction Stop -WarningAction SilentlyContinue

Write-Host ""
Write-Host "Connecting..." -ForegroundColor Cyan
Write-Host "(Browser sign-in window will open — authenticate as a SharePoint Administrator of the destination tenant)" -ForegroundColor Yellow
try {
    Connect-SPOService -Url $destAdminUrl -ErrorAction Stop
    Write-Host "Connected" -ForegroundColor Green
} catch {
    Write-Error "Connection failed: $_"
    exit 1
}

$purged  = 0
$skipped = 0
$failed  = 0

foreach ($url in $destUrls) {
    Write-Host ""
    Write-Host "Checking: $url" -ForegroundColor White

    # Check active sites first
    $site = $null
    try { $site = Get-SPOSite -Identity $url -ErrorAction Stop } catch { }

    if ($site -and $site.Status -ne 'Recycled') {
        Write-Host "  Site is active (Status=$($site.Status)) — skipping" -ForegroundColor Gray
        $skipped++
        continue
    }

    # Look for it in the deleted sites recycle bin
    $deleted = $null
    try { $deleted = Get-SPODeletedSite -Identity $url -ErrorAction Stop } catch { }

    if (-not $deleted) {
        Write-Host "  Not found in recycle bin — nothing to purge" -ForegroundColor Gray
        $skipped++
        continue
    }

    Write-Host "  Found in recycle bin — purging permanently..." -ForegroundColor Yellow
    try {
        Remove-SPODeletedSite -Identity $url -Confirm:$false -ErrorAction Stop
        Write-Host "  Purged" -ForegroundColor Green
        $purged++
    } catch {
        Write-Warning "  Failed to purge: $_"
        $failed++
    }
}

Write-Host ""
$colour = if ($failed -gt 0) { 'Yellow' } else { 'Green' }
Write-Host "=== Done — $purged purged, $skipped skipped, $failed failed ===" -ForegroundColor $colour

try { Disconnect-SPOService -ErrorAction SilentlyContinue } catch {}

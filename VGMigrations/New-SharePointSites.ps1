#Requires -Version 5.1
param(
    [Parameter(Mandatory=$true)]
    [string]$MappingFile,

    [Parameter(Mandatory=$true)]
    [string]$SiteOwner   # UPN of a SharePoint Admin in the destination tenant
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $MappingFile)) {
    Write-Error "Mapping file not found: $MappingFile"
    exit 1
}

Write-Host "Reading: $MappingFile" -ForegroundColor Cyan

# Detect encoding from BOM
$rawBytes = [System.IO.File]::ReadAllBytes($MappingFile)
$enc = if ($rawBytes.Length -ge 3 -and $rawBytes[0] -eq 0xEF -and $rawBytes[1] -eq 0xBB -and $rawBytes[2] -eq 0xBF) {
    'UTF8'
} elseif ($rawBytes.Length -ge 2 -and $rawBytes[0] -eq 0xFF -and $rawBytes[1] -eq 0xFE) {
    'Unicode'
} elseif ($rawBytes.Length -ge 2 -and $rawBytes[0] -eq 0xFE -and $rawBytes[1] -eq 0xFF) {
    'BigEndianUnicode'
} else {
    'Default'
}

$mappings = Import-Csv $MappingFile -Encoding $enc
Write-Host "Loaded $($mappings.Count) row(s)" -ForegroundColor Gray

# Find the destination URL column
$headers = ($mappings | Select-Object -First 1).PSObject.Properties.Name
$destUrlCol = $headers | Where-Object { $_ -imatch '^destination' -and $_ -imatch 'url|site' } | Select-Object -First 1
if (-not $destUrlCol) {
    $destUrlCol = $headers | Where-Object { $_ -imatch '^destination' } | Select-Object -First 1
}
if (-not $destUrlCol) {
    Write-Error "Cannot find a Destination URL column. Columns: $($headers -join ', ')"
    exit 1
}
Write-Host "Using column: $destUrlCol" -ForegroundColor Gray

# Extract unique destination site collection URLs
$destUrls = $mappings |
    Where-Object { $_.$destUrlCol -and $_.$destUrlCol -match 'https://' } |
    Select-Object -ExpandProperty $destUrlCol |
    ForEach-Object {
        $url = $_.Trim()
        # Normalise to site collection root — strip library/folder sub-paths
        if ($url -match '(https://[^/]+/(?:sites|teams)/[^/?#]+)') {
            $Matches[1]
        } elseif ($url -match 'https://') {
            $url
        }
    } | Sort-Object -Unique | Where-Object { $_ }

if (-not $destUrls) {
    Write-Error "No destination SharePoint URLs found in mapping file."
    exit 1
}

Write-Host ""
Write-Host "$($destUrls.Count) unique destination site(s) to check:" -ForegroundColor Cyan
$destUrls | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }

# Derive destination admin URL
$uri = [System.Uri]$destUrls[0]
$destAdminUrl = "$($uri.Scheme)://$($uri.Host -replace '\.sharepoint\.com$', '-admin.sharepoint.com')"
Write-Host ""
Write-Host "Destination admin URL: $destAdminUrl" -ForegroundColor Cyan

# Ensure the SPO module is available
if (-not (Get-Module -Name Microsoft.Online.SharePoint.PowerShell -ListAvailable)) {
    Write-Error "Microsoft.Online.SharePoint.PowerShell is not installed.`nInstall it with: Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser"
    exit 1
}
Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction Stop -WarningAction SilentlyContinue

Write-Host ""
Write-Host "Connecting to destination SharePoint Online..." -ForegroundColor Cyan

# Prefer PnP PowerShell (device code — works from non-interactive process); fall back to SPO module
$pnpName = @('PnP.PowerShell','SharePointPnPPowerShellOnline') |
    Where-Object { Get-Module -Name $_ -ListAvailable -ErrorAction SilentlyContinue } |
    Select-Object -First 1
$usePnp = $false

if ($pnpName) {
    Import-Module $pnpName -ErrorAction Stop -WarningAction SilentlyContinue
    Write-Host ">>> Visit https://microsoft.com/devicelogin and enter the code shown below <<<" -ForegroundColor Yellow
    try {
        Connect-PnPOnline -Url $destAdminUrl -DeviceLogin -ErrorAction Stop
        Write-Host "Connected via $pnpName" -ForegroundColor Green
        $usePnp = $true
    } catch {
        Write-Warning "PnP connection failed: $_"
    }
}

if (-not $usePnp) {
    if (-not (Get-Module -Name Microsoft.Online.SharePoint.PowerShell -ListAvailable)) {
        Write-Error "Neither PnP.PowerShell nor Microsoft.Online.SharePoint.PowerShell is installed."
        exit 1
    }
    Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction Stop -WarningAction SilentlyContinue
    Write-Host "(A sign-in window will open — authenticate as SharePoint Administrator)" -ForegroundColor Yellow
    try {
        Connect-SPOService -Url $destAdminUrl -ErrorAction Stop
        Write-Host "Connected via SPO module" -ForegroundColor Green
    } catch {
        Write-Error "Connection failed: $_"
        exit 1
    }
}

# Enable subsite creation
Write-Host ""
Write-Host "Enabling subsite creation on destination tenant..." -ForegroundColor Cyan
$subSiteEnabled = $false

if ($usePnp) {
    try {
        Set-PnPTenant -DisableSubSiteCreation $false -ErrorAction Stop
        Write-Host "  Enabled (PnP)" -ForegroundColor Green
        $subSiteEnabled = $true
    } catch {
        Write-Warning "  Could not enable subsite creation: $_"
    }
} else {
    try {
        Set-SPOTenant -DisableSubSiteCreation $false -ErrorAction Stop
        Write-Host "  Enabled (SPO module)" -ForegroundColor Green
        $subSiteEnabled = $true
    } catch {
        Write-Warning "  SPO module error: $_"
    }
}

if (-not $subSiteEnabled) {
    $settingsUrl = "$destAdminUrl/_layouts/15/online/AdminHome.aspx#/settings/subSiteCreation"
    Write-Host ""
    Write-Host "  Could not enable subsite creation automatically." -ForegroundColor Red
    Write-Host "  Opening SharePoint Admin Center in your browser now..." -ForegroundColor Yellow
    Start-Process $settingsUrl
    Write-Host ""
    Write-Host "  In the browser that just opened:" -ForegroundColor Cyan
    Write-Host "    1. Sign in as $SiteOwner if prompted" -ForegroundColor White
    Write-Host "    2. Under 'Subsite creation', select 'Let users create subsites'" -ForegroundColor White
    Write-Host "    3. Click Save" -ForegroundColor White
    Write-Host "    4. Return here and re-run the Fly migration" -ForegroundColor White
    Write-Host ""
}

$created = 0
$existed = 0
$failed  = 0

foreach ($url in $destUrls) {
    Write-Host ""
    Write-Host "Checking: $url" -ForegroundColor White

    $site = $null
    try {
        if ($usePnp) { $site = Get-PnPTenantSite -Url $url -ErrorAction Stop }
        else         { $site = Get-SPOSite -Identity $url -ErrorAction Stop }
    } catch { }

    # A deleted site sits in the recycle bin — status = Recycled
    if ($site -and $site.Status -ne 'Recycled') {
        Write-Host "  Already exists ($($site.Status)) — skipping" -ForegroundColor Gray
        $existed++
        continue
    }

    if ($site -and $site.Status -eq 'Recycled') {
        Write-Host "  Found in recycle bin — removing before recreating..." -ForegroundColor Yellow
        try {
            if ($usePnp) { Remove-PnPTenantDeletedSite -Url $url -Force -ErrorAction Stop }
            else         { Remove-SPODeletedSite -Identity $url -Confirm:$false -ErrorAction Stop }
            Write-Host "  Removed from recycle bin" -ForegroundColor Gray
        } catch {
            Write-Warning "  Could not remove from recycle bin: $_"
        }
    }

    $title = ($url -split '/')[-1] -replace '[_-]+', ' '
    Write-Host "  Creating '$title'..." -ForegroundColor Yellow
    try {
        if ($usePnp) {
            New-PnPSite -Type CommunicationSite -Url $url -Owner $SiteOwner -Title $title -ErrorAction Stop
        } else {
            New-SPOSite -Url $url -Owner $SiteOwner -StorageQuota 1024 -Title $title -Template 'STS#3' -ErrorAction Stop
        }
        Write-Host "  Created OK" -ForegroundColor Green
        $created++
    } catch {
        Write-Warning "  Failed: $_"
        $failed++
    }
}

Write-Host ""
$colour = if ($failed -gt 0) { 'Yellow' } else { 'Green' }
Write-Host "=== Done — $created created, $existed already existed, $failed failed ===" -ForegroundColor $colour

try { if ($usePnp) { Disconnect-PnPOnline -ErrorAction SilentlyContinue } else { Disconnect-SPOService -ErrorAction SilentlyContinue } } catch {}

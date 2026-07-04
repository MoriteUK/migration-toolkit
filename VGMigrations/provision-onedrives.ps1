#Requires -Version 7.0
<#
.SYNOPSIS
    Pre-provisions OneDrive sites from a mapping CSV (headless, streaming output).
.PARAMETER MappingFile
    Path to CSV containing destination UPNs.
.PARAMETER AdminUrl
    SharePoint Online admin URL (e.g. https://tenant-admin.sharepoint.com).
.PARAMETER Column
    Column name to read UPNs from. Auto-detected if omitted.
.PARAMETER WhatIf
    Preview only — list UPNs that would be submitted without making changes.
#>
param(
    [Parameter(Mandatory=$true)]  [string]$MappingFile,
    [Parameter(Mandatory=$true)]  [string]$AdminUrl,
    [Parameter(Mandatory=$false)] [string]$Column = "",
    [Parameter(Mandatory=$false)] [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'

Write-Host "=== Provision OneDrives ===" -ForegroundColor Cyan
Write-Host "File:      $MappingFile"
Write-Host "Admin URL: $AdminUrl"
if ($WhatIf) { Write-Host "Mode:      WhatIf (no changes will be made)" -ForegroundColor Yellow }

# Read mapping file
if (-not (Test-Path $MappingFile)) { Write-Error "Mapping file not found: $MappingFile"; exit 1 }
$ext = [System.IO.Path]::GetExtension($MappingFile).ToLowerInvariant()
if ($ext -ne '.csv') { Write-Error "Only CSV files are supported. Got: $ext"; exit 1 }

$rows = @(Import-Csv -Path $MappingFile -Encoding UTF8 -ErrorAction Stop)
if ($rows.Count -eq 0) { Write-Error "No rows found in CSV."; exit 1 }
Write-Host "$($rows.Count) row(s) loaded." -ForegroundColor Green

$cols = @($rows[0].PSObject.Properties.Name)
Write-Host "Columns: $($cols -join ', ')"

# Auto-detect UPN column
$candidates = @('Destination user','Destination','DestinationUPN','DestinationUserUPN',
                'DestinationUserPrincipalName','TargetUPN','Target','UserPrincipalName','UPN')
$col = $null
if ($Column) {
    if ($cols -contains $Column) { $col = $Column }
    else { Write-Error "Column '$Column' not found. Available: $($cols -join ', ')"; exit 1 }
} else {
    foreach ($c in $candidates) { if ($cols -contains $c) { $col = $c; break } }
    if (-not $col) {
        Write-Error "Could not auto-detect UPN column. Tried: $($candidates -join ', ').`nUse the Column override field."
        exit 1
    }
}
Write-Host "Using column: '$col'" -ForegroundColor Green

# Extract and validate UPNs
$upns = @(
    $rows |
    ForEach-Object { if ($_.$col) { $_.$col.ToString().Trim() } } |
    Where-Object   { $_ -match '^[^@]+@[^@]+\.[^@]+$' } |
    Select-Object  -Unique
)
$dropped = $rows.Count - $upns.Count
if ($dropped -gt 0) { Write-Warning "$dropped row(s) skipped — blank or invalid UPNs." }
if ($upns.Count -eq 0) { Write-Error "No valid UPNs found."; exit 1 }
Write-Host "$($upns.Count) unique UPN(s) ready." -ForegroundColor Green

if ($WhatIf) {
    Write-Host "`nWhatIf — would submit $($upns.Count) UPN(s) to: $AdminUrl"
    $upns | ForEach-Object { Write-Host "  $_" }
    Write-Host "`n=== WhatIf complete — no changes made ===" -ForegroundColor Yellow
    exit 0
}

# Load SPO module
Write-Host "`nLoading SharePoint Online module..." -ForegroundColor Cyan
$mod = Get-Module -ListAvailable -Name 'Microsoft.Online.SharePoint.PowerShell' -ErrorAction SilentlyContinue
if (-not $mod) {
    Write-Error "Microsoft.Online.SharePoint.PowerShell is not installed.`nRun: Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser"
    exit 1
}
Import-Module 'Microsoft.Online.SharePoint.PowerShell' -DisableNameChecking -ErrorAction Stop
Write-Host "SPO module $($mod.Version) loaded." -ForegroundColor Green

# Connect — use -UseWebLogin so the OS default browser handles auth
# (avoids the embedded popup that fails when spawned from a non-interactive process)
Write-Host "Connecting to $AdminUrl..." -ForegroundColor Cyan
Write-Host "(A browser sign-in tab will open — authenticate and then return here)" -ForegroundColor Yellow
Connect-SPOService -Url $AdminUrl.TrimEnd('/') -UseWebLogin -ErrorAction Stop
Write-Host "Connected to SharePoint Online." -ForegroundColor Green

# Submit in batches of 200
$batchSize  = 200
$submitted  = 0
$failed     = 0
$totalBatch = [math]::Ceiling($upns.Count / $batchSize)

for ($i = 0; $i -lt $upns.Count; $i += $batchSize) {
    $end   = [math]::Min($i + $batchSize, $upns.Count)
    $chunk = $upns[$i..($end - 1)]
    $batch = [math]::Floor($i / $batchSize) + 1
    Write-Host "Batch $batch/${totalBatch}: submitting $($chunk.Count) UPN(s)..." -ForegroundColor Cyan
    try {
        Request-SPOPersonalSite -UserEmails $chunk -ErrorAction Stop
        $submitted += $chunk.Count
        Write-Host "  Batch $batch accepted." -ForegroundColor Green
    } catch {
        $failed += $chunk.Count
        Write-Warning "  Batch $batch failed: $($_.Exception.Message.Split([Environment]::NewLine)[0])"
    }
}

Write-Host "`n=== Provisioning complete ===" -ForegroundColor Green
Write-Host "Submitted: $submitted   Failed: $failed"
if ($failed -gt 0) {
    Write-Warning "$failed UPN(s) were in failed batches."
    exit 1
}

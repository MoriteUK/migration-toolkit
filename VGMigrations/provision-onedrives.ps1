#Requires -Version 7.0
<#
.SYNOPSIS
    Pre-provisions OneDrive sites from a mapping CSV (headless, streaming output).
.DESCRIPTION
    Uses SharePoint Online's Request-SPOPersonalSite - the Microsoft-documented way to bulk
    queue OneDrive provisioning for a list of users. An earlier version of this script tried to
    provision OneDrive by calling Microsoft Graph's Get-MgUserDrive (GET /users/{id}/drive) on
    the theory that reading a user's drive would trigger provisioning if it didn't exist yet -
    confirmed live (2026-09-01, 0/61 succeeded) that this is not the case: Get-MgUserDrive is a
    plain read and returns 'ResourceNotFound: User's mysite not found' for every not-yet-provisioned
    user, regardless of how many times it's called. Request-SPOPersonalSite is the only reliable
    bulk pre-provisioning mechanism Microsoft documents for this.

    Provisioning via Request-SPOPersonalSite is asynchronous - this script only confirms the
    batch was accepted, not that each individual OneDrive now exists. Wait a few minutes, then
    use Check-OneDriveStatus.ps1 against the same mapping file to confirm.
.PARAMETER MappingFile
    Path to CSV containing destination UPNs.
.PARAMETER AdminUrl
    SharePoint Online admin URL for the destination tenant, e.g.
    https://tenant-admin.sharepoint.com. Required - this is what Request-SPOPersonalSite
    connects to.
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

# MSAL/SPO exceptions frequently put the actual detail on the second (or later) line, with a
# generic message on the first — taking only the first line silently discards the real reason.
function Get-CleanErrorMessage($ErrorRecord) {
    $lines = @($ErrorRecord.Exception.Message -split "`r?`n" | Where-Object { $_.Trim() })
    if ($lines.Count -eq 0) { return $ErrorRecord.Exception.GetType().Name }
    return ($lines | Select-Object -First 3) -join ' | '
}

Write-Host "=== Provision OneDrives ===" -ForegroundColor Cyan
Write-Host "File: $MappingFile"
if ($WhatIf) { Write-Host "Mode: WhatIf (no changes will be made)" -ForegroundColor Yellow }

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
    Write-Host "`nWhatIf — would submit $($upns.Count) UPN(s) to Request-SPOPersonalSite:"
    $upns | ForEach-Object { Write-Host "  $_" }
    Write-Host "`n=== WhatIf complete — no changes made ===" -ForegroundColor Yellow
    exit 0
}

# Load SharePoint Online module
Write-Host "`nLoading Microsoft.Online.SharePoint.PowerShell..." -ForegroundColor Cyan
try {
    $prev = $WarningPreference
    $WarningPreference = 'SilentlyContinue'
    Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell -DisableNameChecking -ErrorAction Stop
    $WarningPreference = $prev
}
catch {
    $WarningPreference = $prev
    Write-Error "Microsoft.Online.SharePoint.PowerShell not available: $($_.Exception.Message)"
    exit 1
}
Write-Host "Module loaded." -ForegroundColor Green

Write-Host "`nConnecting to SharePoint Online ($AdminUrl)..." -ForegroundColor Cyan
# Disconnect first — a stale session left over from an earlier script run (a different tenant)
# can otherwise be silently reused instead of prompting fresh sign-in.
try { Disconnect-SPOService -ErrorAction SilentlyContinue } catch {}
try {
    Connect-SPOService -Url $AdminUrl -ErrorAction Stop
    Write-Host "Connected to SharePoint Online." -ForegroundColor Green
}
catch {
    Write-Error "SPO connection failed: $($_.Exception.Message)"
    exit 1
}

# Request-SPOPersonalSite accepts up to 200 email addresses per call - batch accordingly.
# -NoWait queues provisioning asynchronously rather than blocking here until every site exists,
# which for dozens/hundreds of users can take well beyond a reasonable script run time.
Write-Host "`nQueuing OneDrive provisioning for $($upns.Count) user(s)..." -ForegroundColor Cyan
$batchSize = 200
$queued    = 0
$failed    = 0

for ($i = 0; $i -lt $upns.Count; $i += $batchSize) {
    $batch = $upns[$i..([math]::Min($i + $batchSize - 1, $upns.Count - 1))]
    try {
        Request-SPOPersonalSite -UserEmails $batch -NoWait -ErrorAction Stop
        foreach ($u in $batch) { Write-Host "  QUEUED $u" -ForegroundColor Green }
        $queued += $batch.Count
    }
    catch {
        $msg = Get-CleanErrorMessage $_
        foreach ($u in $batch) { Write-Host "  FAIL $u — $msg" -ForegroundColor Red }
        $failed += $batch.Count
    }
}

try { Disconnect-SPOService -ErrorAction SilentlyContinue } catch {}

Write-Host "`n=== Provisioning request complete ===" -ForegroundColor Green
Write-Host "Queued: $queued   Failed to queue: $failed"
Write-Host "Provisioning is asynchronous - it can take anywhere from a few minutes to a few hours." -ForegroundColor DarkGray
Write-Host "Re-run Check-OneDriveStatus.ps1 against the same mapping file later to confirm which OneDrives now exist." -ForegroundColor DarkGray
if ($failed -gt 0) {
    Write-Warning "$failed UPN(s) could not be queued for provisioning."
    exit 1
}

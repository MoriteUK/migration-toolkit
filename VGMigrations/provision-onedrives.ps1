#Requires -Version 7.0
<#
.SYNOPSIS
    Pre-provisions OneDrive sites from a mapping CSV (headless, streaming output).
.PARAMETER MappingFile
    Path to CSV containing destination UPNs.
.PARAMETER AdminUrl
    SharePoint Online admin URL — kept for UI compatibility, not used for auth.
.PARAMETER Column
    Column name to read UPNs from. Auto-detected if omitted.
.PARAMETER WhatIf
    Preview only — list UPNs that would be submitted without making changes.
#>
param(
    [Parameter(Mandatory=$true)]  [string]$MappingFile,
    [Parameter(Mandatory=$false)] [string]$AdminUrl = "",
    [Parameter(Mandatory=$false)] [string]$Column = "",
    [Parameter(Mandatory=$false)] [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'

# MSAL/Graph exceptions frequently put the actual detail on the second (or later) line, with a
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
    Write-Host "`nWhatIf — would provision OneDrive for $($upns.Count) UPN(s):"
    $upns | ForEach-Object { Write-Host "  $_" }
    Write-Host "`n=== WhatIf complete — no changes made ===" -ForegroundColor Yellow
    exit 0
}

# Load Microsoft Graph modules
Write-Host "`nLoading Microsoft Graph modules..." -ForegroundColor Cyan
. (Join-Path $PSScriptRoot 'Ensure-GraphModules.ps1') -GraphModules @('Microsoft.Graph.Files')
Write-Host "Graph modules loaded." -ForegroundColor Green

Write-Host "`nConnecting to Microsoft Graph — sign in with the browser window that opens..." -ForegroundColor Cyan
# Disconnect first — a stale session left over from an earlier script run (narrower scopes, or
# a different tenant) can otherwise be silently reused instead of prompting fresh sign-in.
try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
Connect-MgGraph -Scopes 'Sites.ReadWrite.All','User.Read.All' -ContextScope Process -NoWelcome -ErrorAction Stop
$ctx = Get-MgContext
Write-Host "Connected to Microsoft Graph as $($ctx.Account) (tenant $($ctx.TenantId))." -ForegroundColor Green
Write-Host "Granted scopes: $($ctx.Scopes -join ', ')" -ForegroundColor DarkGray
if ($ctx.Scopes -notcontains 'Sites.ReadWrite.All' -and $ctx.Scopes -notcontains 'Sites.FullControl.All') {
    Write-Host "WARNING: 'Sites.ReadWrite.All' was not granted — every Get-MgUserDrive call below will fail. Re-run and approve that permission when prompted, or ask an admin to grant consent." -ForegroundColor Yellow
}

# Accessing each user's drive triggers OneDrive provisioning if not yet provisioned
Write-Host "`nProvisioning $($upns.Count) OneDrive(s)..." -ForegroundColor Cyan
$ok   = 0
$fail = 0

foreach ($upn in $upns) {
    try {
        $drive = Get-MgUserDrive -UserId $upn -ErrorAction Stop
        Write-Host "  OK  $upn  ($($drive.DriveType))" -ForegroundColor Green
        $ok++
    } catch {
        $msg = Get-CleanErrorMessage $_
        Write-Host "  FAIL $upn — $msg" -ForegroundColor Red
        $fail++
    }
}

try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}

Write-Host "`n=== Provisioning complete ===" -ForegroundColor Green
Write-Host "Provisioned: $ok   Failed: $fail"
if ($fail -gt 0) {
    Write-Warning "$fail UPN(s) could not be provisioned."
    exit 1
}

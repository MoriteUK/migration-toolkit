#Requires -Version 7.0
<#
.SYNOPSIS
    Checks which users in a mapping CSV have OneDrive provisioned and which don't.
.PARAMETER MappingFile
    Path to CSV containing destination UPNs.
.PARAMETER AdminUrl
    SharePoint Online admin URL — kept for UI compatibility, not used for auth.
.PARAMETER Column
    Column name to read UPNs from. Auto-detected if omitted.
.PARAMETER ExportCleanCsv
    Optional path to write a filtered CSV containing only rows where OneDrive exists.
#>
param(
    [Parameter(Mandatory=$true)]  [string]$MappingFile,
    [Parameter(Mandatory=$false)] [string]$AdminUrl = "",
    [Parameter(Mandatory=$false)] [string]$Column = "",
    [Parameter(Mandatory=$false)] [string]$ExportCleanCsv = ""
)

$ErrorActionPreference = 'Stop'

Write-Host "=== Check OneDrive Status ===" -ForegroundColor Cyan
Write-Host "File: $MappingFile"
if ($ExportCleanCsv) { Write-Host "Export to: $ExportCleanCsv" }

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

# Extract unique UPNs (preserving row order for export)
$seen = @{}
$orderedUpns = [System.Collections.Generic.List[string]]::new()
foreach ($row in $rows) {
    $val = if ($row.$col) { $row.$col.ToString().Trim() } else { '' }
    if ($val -match '^[^@]+@[^@]+\.[^@]+$' -and -not $seen.ContainsKey($val.ToLower())) {
        $seen[$val.ToLower()] = $true
        $orderedUpns.Add($val)
    }
}
$dropped = $rows.Count - $orderedUpns.Count
if ($dropped -gt 0) { Write-Warning "$dropped row(s) skipped — blank, invalid, or duplicate UPNs." }
if ($orderedUpns.Count -eq 0) { Write-Error "No valid UPNs found."; exit 1 }
Write-Host "$($orderedUpns.Count) unique UPN(s) to check." -ForegroundColor Green

# Load Microsoft Graph modules
Write-Host "`nLoading Microsoft Graph modules..." -ForegroundColor Cyan
. (Join-Path $PSScriptRoot 'Ensure-GraphModules.ps1') -GraphModules @('Microsoft.Graph.Files')
Write-Host "Graph modules loaded." -ForegroundColor Green

Write-Host "`nConnecting to Microsoft Graph — sign in with the browser window that opens..." -ForegroundColor Cyan
Connect-MgGraph -Scopes 'Sites.ReadWrite.All','User.Read.All' -NoWelcome -ErrorAction Stop
Write-Host "Connected to Microsoft Graph." -ForegroundColor Green

# Check each UPN for OneDrive
Write-Host "`nChecking $($orderedUpns.Count) UPN(s)..." -ForegroundColor Cyan
$withOdb    = [System.Collections.Generic.List[string]]::new()
$withoutOdb = [System.Collections.Generic.List[string]]::new()

foreach ($upn in $orderedUpns) {
    try {
        $drive = Get-MgUserDrive -UserId $upn -ErrorAction Stop
        if ($drive -and $drive.Id) {
            $withOdb.Add($upn)
        } else {
            $withoutOdb.Add($upn)
        }
    } catch {
        $withoutOdb.Add($upn)
    }
}

try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}

# Output results
Write-Host "`n--- Has OneDrive ($($withOdb.Count)) ---" -ForegroundColor Green
foreach ($u in $withOdb) { Write-Host "  v $u" -ForegroundColor Green }

Write-Host "`n--- No OneDrive ($($withoutOdb.Count)) ---" -ForegroundColor Yellow
foreach ($u in $withoutOdb) { Write-Host "  x $u" -ForegroundColor Yellow }

Write-Host "`n--- Summary ---" -ForegroundColor Cyan
Write-Host "  Checked:       $($orderedUpns.Count)"
Write-Host ("  Has OneDrive:  " + $withOdb.Count) -ForegroundColor Green
Write-Host ("  No OneDrive:   " + $withoutOdb.Count) -ForegroundColor Yellow

# Export clean CSV
if ($ExportCleanCsv) {
    $keepSet = @{}
    foreach ($u in $withOdb) { $keepSet[$u.ToLower()] = $true }
    $cleanRows = $rows | Where-Object {
        $v = if ($_.$col) { $_.$col.ToString().Trim().ToLower() } else { '' }
        $keepSet.ContainsKey($v)
    }
    $cleanRows | Export-Csv -Path $ExportCleanCsv -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
    Write-Host "`nClean CSV exported: $ExportCleanCsv ($($cleanRows.Count) rows)" -ForegroundColor Green
}

Write-Host "`n=== Check complete ===" -ForegroundColor Cyan

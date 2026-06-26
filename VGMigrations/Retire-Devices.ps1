#Requires -Version 7.0
<#
.SYNOPSIS
    Retire-Devices.ps1 — Disable Entra ID devices from a discovery CSV (headless).

.DESCRIPTION
    Reads 11_Devices.csv from a discovery folder (or a specific CSV file) and disables
    each device in Entra ID by setting AccountEnabled = $false via Microsoft Graph.
    The device record is preserved — it is not deleted.

    Use this during domain removal to prevent old-domain devices from authenticating
    without permanently removing the device objects.

.PARAMETER DiscoveryFolder
    Path to a discovery folder containing 11_Devices.csv.
    If the folder is not named 'Discovery', the script looks for a Discovery subfolder.

.PARAMETER CsvFile
    Direct path to a devices CSV. Takes precedence over DiscoveryFolder.
    CSV must have a DeviceId or DeviceObjectId column.

.PARAMETER WhatIf
    List devices that would be retired without making any changes.
#>

param(
    [string]$DiscoveryFolder = '',
    [string]$CsvFile         = '',
    [switch]$WhatIf
)

# Resolve CSV path
$csvPath = ''
if ($CsvFile) {
    $csvPath = $CsvFile.Trim().Trim('"')
} elseif ($DiscoveryFolder) {
    $folder    = $DiscoveryFolder.Trim().Trim('"')
    $candidate = Join-Path $folder 'Discovery'
    if ((Split-Path $folder -Leaf) -ne 'Discovery' -and (Test-Path $candidate)) { $folder = $candidate }
    $csvPath = Join-Path $folder '11_Devices.csv'
}

if (-not $csvPath -or -not (Test-Path $csvPath)) {
    Write-Host "ERROR: Device CSV not found: $csvPath"
    Write-Host 'Usage: Retire-Devices.ps1 -DiscoveryFolder <path>  OR  -CsvFile <path>'
    exit 1
}

$devices = @(Import-Csv -Path $csvPath -Encoding UTF8)
Write-Host "=== Retire Devices$(if ($WhatIf) { ' [WhatIf]' }) ==="
Write-Host "CSV          : $csvPath"
Write-Host "Device count : $($devices.Count)"

if ($devices.Count -eq 0) { Write-Host 'No devices in CSV.'; exit 0 }

# Detect ID column
$cols = @($devices[0].PSObject.Properties.Name)
$idCol = if ($cols -contains 'DeviceObjectId') { 'DeviceObjectId' }
         elseif ($cols -contains 'DeviceId')   { 'DeviceId' }
         else { Write-Host "ERROR: CSV must have a 'DeviceObjectId' or 'DeviceId' column. Found: $($cols -join ', ')"; exit 1 }
Write-Host "ID column    : $idCol"
Write-Host ''

# Load Graph modules
$graphMods = @('Microsoft.Graph.Authentication','Microsoft.Graph.Identity.DirectoryManagement')
foreach ($m in $graphMods) {
    if (-not (Get-Module -ListAvailable -Name $m -ErrorAction SilentlyContinue)) {
        Write-Host "ERROR: Required module not installed: $m"
        Write-Host "Install it with:  Install-Module Microsoft.Graph -Scope CurrentUser -Force"
        exit 1
    }
    Import-Module $m -ErrorAction Stop
}

Write-Host 'Connecting to Microsoft Graph...'
try {
    Connect-MgGraph -Scopes 'Device.ReadWrite.All','Directory.ReadWrite.All' `
        -NoWelcome -ErrorAction Stop
    Write-Host 'Connected.'
} catch {
    Write-Host "ERROR: Could not connect to Microsoft Graph: $($_.Exception.Message.Split([Environment]::NewLine)[0])"
    exit 1
}

$ok = 0; $fail = 0; $skip = 0; $i = 0
foreach ($d in $devices) {
    $i++
    $id    = $d.$idCol
    $name  = $d.DeviceName
    $owner = if ($d.PSObject.Properties['OwnerUPN']) { $d.OwnerUPN } else { '' }

    if (-not $id) { Write-Host "  [$i/$($devices.Count)] SKIPPED — no ID: $name"; $skip++; continue }

    if ($WhatIf) {
        Write-Host "  [$i/$($devices.Count)] WhatIf : $name  [$id]$(if ($owner) { "  (owner: $owner)" })"
        $ok++
    } else {
        try {
            Update-MgDevice -DeviceId $id -AccountEnabled $false -ErrorAction Stop
            Write-Host "  [$i/$($devices.Count)] Retired : $name  [$id]$(if ($owner) { "  (owner: $owner)" })"
            $ok++
        } catch {
            Write-Host "  [$i/$($devices.Count)] FAILED  : $name — $($_.Exception.Message.Split([Environment]::NewLine)[0])"
            $fail++
        }
    }
}

Write-Host ''
Write-Host "=== Complete: $ok retired  |  $fail failed  |  $skip skipped ==="
Write-Host 'Note: Retired devices remain in Entra ID but cannot authenticate. Delete them later if required.'

try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}

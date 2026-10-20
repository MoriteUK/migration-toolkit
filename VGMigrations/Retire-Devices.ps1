#Requires -Version 7.0
<#
.SYNOPSIS
    Retire-Devices.ps1 — Disable Entra ID devices from a discovery CSV (headless).

.DESCRIPTION
    Reads 12_Devices.csv from a discovery folder (or a specific CSV file) and disables
    each device in Entra ID by setting AccountEnabled = $false via Microsoft Graph.
    The device record is preserved — it is not deleted.

    Use this during domain removal to prevent old-domain devices from authenticating
    without permanently removing the device objects.

.PARAMETER DiscoveryFolder
    Path to a discovery folder containing 12_Devices.csv.
    If the folder is not named 'Discovery', the script looks for a Discovery subfolder.

.PARAMETER CsvFile
    Direct path to a devices CSV. Takes precedence over DiscoveryFolder.
    CSV must have a DeviceId or DeviceObjectId column.

.PARAMETER WhatIf
    List devices that would be retired without making any changes.

.PARAMETER Domain
    Customer tenant domain or tenant ID to sign in to (passed to Connect-MgGraph -TenantId).
    Without this, sign-in lands in the signed-in account's home tenant, which will not contain
    the device object IDs from the discovery CSV if that account isn't native to the customer
    tenant — every Update-MgDevice call then fails with Request_ResourceNotFound.

.PARAMETER UserListFile
    Optional text file containing usernames (one per line) to filter devices. Only devices
    owned by users in this list will be retired. Use this to selectively retire devices for
    users being migrated while preserving devices for users with old domain aliases who aren't
    being migrated. The script matches against the OwnerUPN column in the devices CSV.
#>

param(
    [string]$DiscoveryFolder = '',
    [string]$CsvFile         = '',
    [string]$Domain          = '',
    [string]$UserListFile    = '',
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
    $csvPath = Join-Path $folder '12_Devices.csv'
}

if (-not $csvPath -or -not (Test-Path $csvPath)) {
    Write-Host "ERROR: Device CSV not found: $csvPath"
    Write-Host 'Usage: Retire-Devices.ps1 -DiscoveryFolder <path>  OR  -CsvFile <path>'
    exit 1
}

$allDevices = @(Import-Csv -Path $csvPath -Encoding UTF8)
Write-Host "=== Retire Devices$(if ($WhatIf) { ' [WhatIf]' }) ==="
Write-Host "CSV          : $csvPath"
Write-Host "Total devices: $($allDevices.Count)"

if ($allDevices.Count -eq 0) { Write-Host 'No devices in CSV.'; exit 0 }

# Filter by user list if provided
$devices = $allDevices
if ($UserListFile) {
    if (-not (Test-Path $UserListFile)) {
        Write-Host "ERROR: User list file not found: $UserListFile" -ForegroundColor Red
        exit 1
    }

    $userList = @(Get-Content -Path $UserListFile -Encoding UTF8 |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ -and -not $_.StartsWith('#') })

    Write-Host "User list    : $UserListFile"
    Write-Host "Users loaded : $($userList.Count)"

    # Check if devices have OwnerUPN column
    $cols = @($allDevices[0].PSObject.Properties.Name)
    if ($cols -notcontains 'OwnerUPN') {
        Write-Host "ERROR: Devices CSV must have an 'OwnerUPN' column when filtering by user list. Found: $($cols -join ', ')" -ForegroundColor Red
        exit 1
    }

    # Filter devices to only those owned by users in the list (case-insensitive)
    $userSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($u in $userList) { $userSet.Add($u) | Out-Null }

    $devices = @($allDevices | Where-Object {
        $owner = $_.OwnerUPN
        $owner -and $userSet.Contains($owner)
    })

    Write-Host "Filtered to  : $($devices.Count) device(s) owned by users in the list" -ForegroundColor Cyan

    if ($devices.Count -eq 0) {
        Write-Host 'No devices match the user list.' -ForegroundColor Yellow
        exit 0
    }
}

Write-Host "Retiring     : $($devices.Count) device(s)"

if ($devices.Count -eq 0) { Write-Host 'No devices to retire.'; exit 0 }

# Detect ID column
$cols = @($devices[0].PSObject.Properties.Name)
$idCol = if ($cols -contains 'DeviceObjectId') { 'DeviceObjectId' }
         elseif ($cols -contains 'DeviceId')   { 'DeviceId' }
         else { Write-Host "ERROR: CSV must have a 'DeviceObjectId' or 'DeviceId' column. Found: $($cols -join ', ')"; exit 1 }
Write-Host "ID column    : $idCol"
Write-Host ''

. (Join-Path $PSScriptRoot 'Ensure-GraphModules.ps1') -GraphModules @('Microsoft.Graph.Identity.DirectoryManagement')

if ($Domain) {
    Write-Host "Connecting to Microsoft Graph (tenant: $Domain)..."
} else {
    Write-Host 'WARNING: No -Domain specified — signing in without -TenantId will connect to the' -ForegroundColor Yellow
    Write-Host 'signed-in account''s home tenant, not necessarily the customer tenant the devices' -ForegroundColor Yellow
    Write-Host 'live in. If every device then fails with Request_ResourceNotFound, re-run with -Domain <customer tenant domain or ID>.' -ForegroundColor Yellow
    Write-Host 'Connecting to Microsoft Graph...'
}
try {
    $connectParams = @{
        Scopes      = 'Device.ReadWrite.All','Directory.ReadWrite.All'
        NoWelcome   = $true
        ErrorAction = 'Stop'
    }
    if ($Domain) { $connectParams['TenantId'] = $Domain }
    Connect-MgGraph @connectParams
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
            Update-MgDevice -DeviceId $id -AccountEnabled:$false -ErrorAction Stop
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

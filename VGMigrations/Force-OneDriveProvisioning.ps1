#Requires -Version 7.0
<#
.SYNOPSIS
    Force OneDrive provisioning for users that haven't provisioned after running provision-onedrives.ps1

.DESCRIPTION
    This script helps diagnose and force OneDrive provisioning when the original provision request
    didn't complete. It:

    1. Checks which OneDrives actually exist now
    2. Checks user license status (users need OneDrive for Business license)
    3. Force re-provisions any that are still missing using Request-SPOPersonalSite WITHOUT -NoWait

    Common reasons OneDrive provisioning fails:
    - User doesn't have a license (needs at least OneDrive for Business)
    - User is not enabled/active in Azure AD
    - User's mailbox doesn't exist (Exchange license required)
    - Tenant storage quota exceeded
    - User's OneDrive was previously deleted and is in recycle bin

.PARAMETER MappingFile
    Path to the same CSV file used with provision-onedrives.ps1

.PARAMETER AdminUrl
    SharePoint Online admin URL, e.g. https://tenant-admin.sharepoint.com

.PARAMETER Column
    Column name containing UPNs (auto-detected if omitted)

.PARAMETER CheckOnly
    Only check status without forcing provisioning

.EXAMPLE
    .\Force-OneDriveProvisioning.ps1 -MappingFile .\mapping.csv -AdminUrl https://contoso-admin.sharepoint.com

.EXAMPLE
    .\Force-OneDriveProvisioning.ps1 -MappingFile .\mapping.csv -AdminUrl https://contoso-admin.sharepoint.com -CheckOnly
#>
param(
    [Parameter(Mandatory=$true)]  [string]$MappingFile,
    [Parameter(Mandatory=$true)]  [string]$AdminUrl,
    [Parameter(Mandatory=$false)] [string]$Column = "",
    [Parameter(Mandatory=$false)] [switch]$CheckOnly
)

$ErrorActionPreference = 'Stop'

Write-Host "=== Force OneDrive Provisioning ===" -ForegroundColor Cyan
Write-Host "File: $MappingFile"
if ($CheckOnly) { Write-Host "Mode: Check Only (no provisioning)" -ForegroundColor Yellow }

# Read mapping file
if (-not (Test-Path $MappingFile)) { Write-Error "Mapping file not found: $MappingFile"; exit 1 }
$rows = @(Import-Csv -Path $MappingFile -Encoding UTF8 -ErrorAction Stop)
if ($rows.Count -eq 0) { Write-Error "No rows in CSV"; exit 1 }
Write-Host "$($rows.Count) row(s) loaded." -ForegroundColor Green

# Auto-detect UPN column
$cols = @($rows[0].PSObject.Properties.Name)
$candidates = @('Destination user','Destination','DestinationUPN','UserPrincipalName','UPN')
$col = $null
if ($Column) {
    if ($cols -contains $Column) { $col = $Column }
    else { Write-Error "Column '$Column' not found"; exit 1 }
} else {
    foreach ($c in $candidates) { if ($cols -contains $c) { $col = $c; break } }
    if (-not $col) { Write-Error "Could not auto-detect UPN column. Use -Column"; exit 1 }
}
Write-Host "Using column: '$col'" -ForegroundColor Green

# Extract UPNs
$upns = @(
    $rows | ForEach-Object { if ($_.$col) { $_.$col.ToString().Trim() } } |
    Where-Object { $_ -match '^[^@]+@[^@]+\.[^@]+$' } |
    Select-Object -Unique
)
if ($upns.Count -eq 0) { Write-Error "No valid UPNs found"; exit 1 }
Write-Host "$($upns.Count) unique UPN(s)" -ForegroundColor Green

# ========================================================================
# Step 1: Check OneDrive status via Microsoft Graph
# ========================================================================
Write-Host "`n=== Step 1: Checking OneDrive Status ===" -ForegroundColor Cyan
. (Join-Path $PSScriptRoot 'Ensure-GraphModules.ps1') -GraphModules @('Microsoft.Graph.Users','Microsoft.Graph.Files')

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
Connect-MgGraph -Scopes 'User.Read.All','Sites.ReadWrite.All' -ContextScope Process -NoWelcome -ErrorAction Stop
Write-Host "Connected to Microsoft Graph" -ForegroundColor Green

$provisionedUpns = @()
$notProvisionedUpns = @()
$licenseIssues = @()

foreach ($upn in $upns) {
    Write-Host "  Checking $upn..." -NoNewline

    # Check OneDrive
    $hasOneDrive = $false
    try {
        $drive = Get-MgUserDrive -UserId $upn -ErrorAction Stop
        if ($drive -and $drive.Id) {
            $hasOneDrive = $true
            $provisionedUpns += $upn
            Write-Host " ✓ OneDrive EXISTS" -ForegroundColor Green
        }
    } catch {
        # OneDrive doesn't exist - check user status
        try {
            $user = Get-MgUser -UserId $upn -Property 'accountEnabled,assignedLicenses' -ErrorAction Stop

            if (-not $user.AccountEnabled) {
                Write-Host " ✗ User DISABLED" -ForegroundColor Red
                $licenseIssues += [PSCustomObject]@{ UPN = $upn; Issue = 'User account disabled' }
            } elseif ($user.AssignedLicenses.Count -eq 0) {
                Write-Host " ✗ NO LICENSES" -ForegroundColor Red
                $licenseIssues += [PSCustomObject]@{ UPN = $upn; Issue = 'No licenses assigned' }
            } else {
                Write-Host " ✗ NOT PROVISIONED (user active, has licenses)" -ForegroundColor Yellow
                $notProvisionedUpns += $upn
            }
        } catch {
            Write-Host " ✗ USER NOT FOUND" -ForegroundColor Red
            $licenseIssues += [PSCustomObject]@{ UPN = $upn; Issue = 'User not found in tenant' }
        }
    }
}

try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}

# Summary
Write-Host "`n--- Status Summary ---" -ForegroundColor Cyan
Write-Host "  ✓ OneDrive EXISTS:       $($provisionedUpns.Count)" -ForegroundColor Green
Write-Host "  ✗ Not provisioned:       $($notProvisionedUpns.Count)" -ForegroundColor Yellow
Write-Host "  ⚠ License/User issues:   $($licenseIssues.Count)" -ForegroundColor Red

if ($licenseIssues.Count -gt 0) {
    Write-Host "`n--- License/User Issues ---" -ForegroundColor Red
    $licenseIssues | Format-Table -AutoSize
    Write-Host "Fix these issues before OneDrive can be provisioned." -ForegroundColor Yellow
}

if ($notProvisionedUpns.Count -eq 0) {
    if ($provisionedUpns.Count -gt 0) {
        Write-Host "`n✓ All OneDrives are provisioned!" -ForegroundColor Green
    }
    exit 0
}

if ($CheckOnly) {
    Write-Host "`nCheck-only mode — not forcing provisioning." -ForegroundColor Yellow
    Write-Host "Run without -CheckOnly to force provisioning for $($notProvisionedUpns.Count) user(s)." -ForegroundColor Cyan
    exit 0
}

# ========================================================================
# Step 2: Force provisioning via SharePoint Online
# ========================================================================
Write-Host "`n=== Step 2: Force OneDrive Provisioning ===" -ForegroundColor Cyan
Write-Host "Will force provision $($notProvisionedUpns.Count) OneDrive(s)..." -ForegroundColor Yellow

# Load SharePoint Online module
Write-Host "`nLoading SharePoint Online module..." -ForegroundColor Cyan
try {
    $prev = $WarningPreference
    $WarningPreference = 'SilentlyContinue'
    Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell -DisableNameChecking -ErrorAction Stop
    $WarningPreference = $prev
    Write-Host "Module loaded" -ForegroundColor Green
} catch {
    $WarningPreference = $prev
    Write-Error "SharePoint module not available: $($_.Exception.Message)"
    exit 1
}

Write-Host "`nConnecting to SharePoint Online ($AdminUrl)..." -ForegroundColor Cyan
Write-Host ">>> Browser will open — sign in as SharePoint Admin <<<" -ForegroundColor Yellow
try { Disconnect-SPOService -ErrorAction SilentlyContinue } catch {}
Connect-SPOService -Url $AdminUrl -UseSystemBrowser:$true -ErrorAction Stop
Write-Host "Connected" -ForegroundColor Green

# Force provisioning WITHOUT -NoWait to ensure they actually get created
Write-Host "`nForce provisioning (this may take several minutes)..." -ForegroundColor Cyan
Write-Host "NOTE: Using synchronous mode (no -NoWait) to verify completion" -ForegroundColor Yellow

$batchSize = 10  # Smaller batches for synchronous mode
$success = 0
$failed = 0

for ($i = 0; $i -lt $notProvisionedUpns.Count; $i += $batchSize) {
    $batch = $notProvisionedUpns[$i..([math]::Min($i + $batchSize - 1, $notProvisionedUpns.Count - 1))]
    Write-Host "`nProcessing batch $([math]::Floor($i / $batchSize) + 1) ($($batch.Count) users)..."

    try {
        # WITHOUT -NoWait = synchronous, waits for actual provisioning
        Request-SPOPersonalSite -UserEmails $batch -ErrorAction Stop
        foreach ($u in $batch) {
            Write-Host "  ✓ $u" -ForegroundColor Green
            $success++
        }
    } catch {
        foreach ($u in $batch) {
            Write-Host "  ✗ $u — $($_.Exception.Message)" -ForegroundColor Red
            $failed++
        }
    }
}

try { Disconnect-SPOService -ErrorAction SilentlyContinue } catch {}

Write-Host "`n=== Complete ===" -ForegroundColor Green
Write-Host "Success: $success   Failed: $failed"

if ($success -gt 0) {
    Write-Host "`n✓ $success OneDrive(s) force-provisioned successfully!" -ForegroundColor Green
    Write-Host "Users can now access their OneDrive at:" -ForegroundColor Cyan
    Write-Host "  https://<tenant>-my.sharepoint.com/personal/<username>_<domain>_<tld>" -ForegroundColor Gray
}

if ($failed -gt 0) {
    Write-Host "`n⚠ $failed OneDrive(s) failed to provision" -ForegroundColor Red
    Write-Host "Common reasons:" -ForegroundColor Yellow
    Write-Host "  - User already has OneDrive (deleted recently, still in recycle bin)" -ForegroundColor Gray
    Write-Host "  - Tenant storage quota exceeded" -ForegroundColor Gray
    Write-Host "  - User needs time to sync (try again in 10 minutes)" -ForegroundColor Gray
}

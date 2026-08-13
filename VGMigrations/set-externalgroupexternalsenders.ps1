#Requires -Version 7.0
<#
.SYNOPSIS
    Audits all distribution groups in a tenant for whether they accept messages
    from external (outside-org) senders, and optionally fixes any that don't.

.DESCRIPTION
    Distribution group "accept external senders" is controlled by the Exchange
    Online property RequireSenderAuthenticationEnabled:
        $true  = only authenticated (internal) senders allowed -> external BLOCKED
        $false = external senders allowed
    This is an Exchange Online mailbox/recipient setting, not a Graph/Entra
    property, so this script uses the ExchangeOnlineManagement module rather
    than Invoke-MgGraphRequest (Graph has no equivalent for classic DLs).

    Run with no switches first to see a report of current state (dry run).
    Add -Commit to actually change any group found to be blocking external mail.

.PARAMETER Commit
    Without this switch, the script only reports findings (dry run).
    With this switch, any group with RequireSenderAuthenticationEnabled = $true
    is updated to $false so it accepts external senders.

.PARAMETER CsvPath
    Optional output folder for the timestamped CSV report. Defaults to current
    directory.

.EXAMPLE
    ./Set-DistributionGroupExternalSenders.ps1
    Dry run - reports all DLs and whether they currently allow external senders.

.EXAMPLE
    ./Set-DistributionGroupExternalSenders.ps1 -Commit
    Reports, then updates every DL that is blocking external senders so it allows them.
#>

[CmdletBinding()]
param(
    [switch]$Commit,
    [string]$CsvPath = (Get-Location).Path
)

$ErrorActionPreference = 'Stop'

function Ensure-EXOConnection {
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Host "ExchangeOnlineManagement module not found - installing for current user..." -ForegroundColor Yellow
        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    $connected = $false
    try {
        Get-ConnectionInformation -ErrorAction Stop | Where-Object { $_.State -eq 'Connected' } | Out-Null
        $connected = $true
    } catch { $connected = $false }

    if (-not $connected) {
        Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
        try {
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        }
        catch {
            Write-Host "Interactive/WAM sign-in failed, falling back to device code..." -ForegroundColor Yellow
            Connect-ExchangeOnline -ShowBanner:$false -Device -ErrorAction Stop
        }
    }
}

Ensure-EXOConnection

Write-Host "Retrieving all distribution groups (this may take a moment for large tenants)..." -ForegroundColor Cyan
$allGroups = Get-DistributionGroup -ResultSize Unlimited

Write-Host "Found $($allGroups.Count) distribution groups. Checking external sender settings..." -ForegroundColor Cyan

$results = foreach ($group in $allGroups) {

    $blocksExternal = $group.RequireSenderAuthenticationEnabled  # $true = internal senders only
    $status         = if ($blocksExternal) { 'Blocked' } else { 'Allowed' }
    $action         = 'None'

    if ($blocksExternal) {
        if ($Commit) {
            try {
                Set-DistributionGroup -Identity $group.Identity -RequireSenderAuthenticationEnabled $false -ErrorAction Stop
                $action = 'Fixed'
                $status = 'Allowed (was Blocked)'
                Write-Host "  [FIXED] $($group.PrimarySmtpAddress)" -ForegroundColor Green
            }
            catch {
                $action = "FAILED: $($_.Exception.Message)"
                Write-Host "  [FAILED] $($group.PrimarySmtpAddress) - $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        else {
            $action = 'Would fix (dry run)'
            Write-Host "  [BLOCKS EXTERNAL] $($group.PrimarySmtpAddress)" -ForegroundColor Yellow
        }
    }

    [PSCustomObject]@{
        DisplayName        = $group.DisplayName
        PrimarySmtpAddress = $group.PrimarySmtpAddress
        Identity           = $group.Identity
        GroupType          = $group.RecipientTypeDetails
        ExternalSenders    = $status
        Action             = $action
    }
}

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$outFile   = Join-Path $CsvPath "DistributionGroup-ExternalSenders_$timestamp.csv"
$results | Export-Csv -Path $outFile -NoTypeInformation -Encoding UTF8

$blockedCount = ($results | Where-Object { $_.ExternalSenders -like 'Blocked*' -or $_.Action -eq 'Would fix (dry run)' -or $_.Action -eq 'Fixed' }).Count

Write-Host ""
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "Total groups checked : $($results.Count)"
Write-Host "Blocking external    : $blockedCount"
if (-not $Commit -and $blockedCount -gt 0) {
    Write-Host "Dry run only - re-run with -Commit to apply fixes." -ForegroundColor Yellow
}
Write-Host "CSV report saved to  : $outFile" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan

Disconnect-ExchangeOnline -Confirm:$false | Out-Null
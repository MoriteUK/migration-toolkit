#Requires -Version 7.0
<#
.SYNOPSIS
    Finds every recipient (mailboxes, mail users, mail contacts, distribution
    groups, etc.) with a SIP/IM proxy address on a given domain, and removes
    just those SIP entries - the classic reason a domain won't let itself be
    removed from a tenant even after everything else has been moved off it.

.DESCRIPTION
    Domain removal in M365 checks ALL proxy addresses, not just the primary
    SMTP address. A leftover "SIP:user@olddomain.com" entry (stamped when a
    Teams/SfB license was first assigned, see prior discussion) is enough to
    block removal even if the mailbox's primary address and UPN have long
    since moved to a different domain.

    This script only touches SIP: type addresses on the target domain. It
    does NOT touch SMTP/smtp entries, UPNs, or anything else - if those are
    also still on the domain, Set-User/Set-Mailbox etc. need to be handled
    separately (this is intentionally narrow, matching the specific SIP
    blocker).

    Uses Set-Recipient -EmailAddresses, which works across recipient types
    (mailboxes, mail users, mail contacts, distribution groups) since it's
    the generic recipient cmdlet rather than a type-specific one.

.PARAMETER Domain
    The domain to search for and clean SIP addresses from, e.g. olddomain.com

.PARAMETER Commit
    Without this switch, the script only reports what it found (dry run).
    With this switch, it actually removes the matching SIP: entries.

.PARAMETER CsvPath
    Optional output folder for the timestamped CSV report. Defaults to current
    directory.

.EXAMPLE
    ./Remove-SipAddressesForDomain.ps1 -Domain olddomain.com
    Dry run - lists every recipient with a SIP: address on olddomain.com.

.EXAMPLE
    ./Remove-SipAddressesForDomain.ps1 -Domain olddomain.com -Commit
    Removes those SIP: entries so the domain can be removed.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$Domain,

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

Write-Host "Searching all recipients for SIP: addresses on '$Domain'..." -ForegroundColor Cyan
$allRecipients = Get-Recipient -ResultSize Unlimited -Properties EmailAddresses

$sipSuffix = "@$Domain"
$matches = $allRecipients | Where-Object {
    $_.EmailAddresses | Where-Object { $_ -like "SIP:*$sipSuffix" }
}

Write-Host "Found $($matches.Count) recipient(s) with a SIP address on $Domain." -ForegroundColor Cyan

$results = foreach ($recipient in $matches) {

    $sipEntries = $recipient.EmailAddresses | Where-Object { $_ -like "SIP:*$sipSuffix" }
    $remaining  = $recipient.EmailAddresses | Where-Object { $_ -notin $sipEntries }
    $action     = 'None'

    if ($Commit) {
        try {
            Set-Recipient -Identity $recipient.Identity -EmailAddresses $remaining -ErrorAction Stop
            $action = 'Removed'
            Write-Host "  [REMOVED] $($recipient.DisplayName) - $($sipEntries -join ', ')" -ForegroundColor Green
        }
        catch {
            $action = "FAILED: $($_.Exception.Message)"
            Write-Host "  [FAILED] $($recipient.DisplayName) - $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    else {
        $action = 'Would remove (dry run)'
        Write-Host "  [FOUND] $($recipient.DisplayName) - $($sipEntries -join ', ')" -ForegroundColor Yellow
    }

    [PSCustomObject]@{
        DisplayName       = $recipient.DisplayName
        RecipientType     = $recipient.RecipientTypeDetails
        Identity          = $recipient.Identity
        SipAddressesFound = ($sipEntries -join '; ')
        Action            = $action
    }
}

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$outFile   = Join-Path $CsvPath "SipAddressCleanup_$($Domain)_$timestamp.csv"
$results | Export-Csv -Path $outFile -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "Recipients with SIP on $Domain : $($results.Count)"
if (-not $Commit -and $results.Count -gt 0) {
    Write-Host "Dry run only - re-run with -Commit to remove these entries." -ForegroundColor Yellow
}
Write-Host "CSV report saved to : $outFile" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan

Disconnect-ExchangeOnline -Confirm:$false | Out-Null
#Requires -Version 7.0
<#
.SYNOPSIS
    Finds every mailbox on a given domain with an active in-place archive
    and disables that archive.

.DESCRIPTION
    Search-Mailbox -DeleteContent (the old way of bulk-purging mailbox/
    archive content) has been deprecated and removed from Exchange Online.
    Its replacement - Security & Compliance eDiscovery purge - is built for
    targeted content removal, doesn't cleanly scope to "archive only", needs
    the Search And Purge role, and processes in slow batches. It's the wrong
    tool for "empty this whole archive."

    The clean, fully-supported way to achieve the same end result is
    Disable-Mailbox -Archive: it immediately soft-deletes the archive and
    all its content (it disappears from Outlook/OWA right away), and Exchange
    permanently purges it after the standard 30-day retention window. If
    someone realises they need it back within that window, it's recoverable
    via Enable-Mailbox -Archive -ArchiveGuid <guid>; after 30 days it's gone
    for good. This script uses that approach rather than attempting a
    per-item purge.

    NOTE: if a mailbox has an active Litigation Hold or retention hold,
    Disable-Mailbox -Archive may not immediately/fully purge content bound
    by that hold - the hold still applies to Recoverable Items even after
    the archive is disabled. The script flags any mailbox on hold in its
    report so you can deal with those separately rather than assuming
    they're fully cleared.

.PARAMETER Domain
    The domain to check, e.g. olddomain.com. Matches against the mailbox's
    primary SMTP address domain.

.PARAMETER Commit
    Without this switch, the script only reports which mailboxes have an
    active archive on this domain (dry run) - nothing is disabled.
    With this switch, it actually disables the archive for each one.

.PARAMETER Mailbox
    Optional. Scopes the run to a single mailbox on the domain (matched
    against PrimarySmtpAddress) instead of every mailbox on the domain.
    Use this to test the script - including -Commit - against one account
    before running it against the whole domain.

.PARAMETER CsvPath
    Optional output folder for the timestamped CSV report. Defaults to
    current directory.

.EXAMPLE
    ./Disable-DomainMailboxArchives.ps1 -Domain olddomain.com
    Dry run - lists every mailbox on olddomain.com with an active archive.

.EXAMPLE
    ./Disable-DomainMailboxArchives.ps1 -Domain olddomain.com -Mailbox test.user@olddomain.com -Commit
    Test run - disables the archive for just that one mailbox.

.EXAMPLE
    ./Disable-DomainMailboxArchives.ps1 -Domain olddomain.com -Commit
    Soft-deletes (disables) the archive for every one of those mailboxes.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$Domain,

    [switch]$Commit,

    [string]$Mailbox,

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

Write-Host "Retrieving mailboxes on '$Domain' with an active archive..." -ForegroundColor Cyan

$mailboxes = Get-EXOMailbox -ResultSize Unlimited -Archive -PropertySets Archive, Hold |
    Where-Object {
        $_.PrimarySmtpAddress -like "*@$Domain" -and $_.ArchiveStatus -eq 'Active'
    }

if ($Mailbox) {
    $mailboxes = $mailboxes | Where-Object { $_.PrimarySmtpAddress -eq $Mailbox }
    if ($mailboxes.Count -eq 0) {
        Write-Host "No mailbox matching '$Mailbox' with an active archive was found on $Domain." -ForegroundColor Red
        Disconnect-ExchangeOnline -Confirm:$false | Out-Null
        exit 0
    }
    Write-Host "Scoped to single mailbox: $Mailbox" -ForegroundColor Cyan
}

Write-Host "Found $($mailboxes.Count) mailbox(es) with an active archive on $Domain." -ForegroundColor Cyan

if ($Commit -and $mailboxes.Count -gt 0) {
    Write-Host ""
    Write-Host "You are about to disable (soft-delete) the archive for $($mailboxes.Count) mailbox(es)." -ForegroundColor Yellow
    Write-Host "Content is recoverable for 30 days, then permanently purged." -ForegroundColor Yellow
    $confirm = Read-Host "Type YES to proceed"
    if ($confirm -ne 'YES') {
        Write-Host "Aborted - no changes made." -ForegroundColor Red
        Disconnect-ExchangeOnline -Confirm:$false | Out-Null
        exit 0
    }
}

$results = foreach ($mbx in $mailboxes) {

    $onHold = [bool]($mbx.LitigationHoldEnabled -or $mbx.InPlaceHolds)
    $action = 'None'

    if ($Commit) {
        try {
            Disable-Mailbox -Archive -Identity $mbx.Identity -Confirm:$false -ErrorAction Stop
            $action = 'Archive disabled'
            Write-Host "  [DISABLED] $($mbx.PrimarySmtpAddress)$(if ($onHold) {' (ON HOLD - verify purge separately)'})" -ForegroundColor $(if ($onHold) {'Yellow'} else {'Green'})
        }
        catch {
            $action = "FAILED: $($_.Exception.Message)"
            Write-Host "  [FAILED] $($mbx.PrimarySmtpAddress) - $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    else {
        $action = 'Would disable (dry run)'
        Write-Host "  [FOUND] $($mbx.PrimarySmtpAddress)$(if ($onHold) {' (ON HOLD)'})" -ForegroundColor Yellow
    }

    [PSCustomObject]@{
        DisplayName        = $mbx.DisplayName
        PrimarySmtpAddress = $mbx.PrimarySmtpAddress
        OnHold             = $onHold
        Action             = $action
    }
}

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$outFile   = Join-Path $CsvPath "DisableArchive_$($Domain)_$timestamp.csv"
$results | Export-Csv -Path $outFile -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "Mailboxes with active archive on $Domain : $($results.Count)"
$holdCount = ($results | Where-Object { $_.OnHold }).Count
if ($holdCount -gt 0) {
    Write-Host "On hold (verify purge separately)         : $holdCount" -ForegroundColor Yellow
}
if (-not $Commit -and $results.Count -gt 0) {
    Write-Host "Dry run only - re-run with -Commit to disable these archives." -ForegroundColor Yellow
}
Write-Host "CSV report saved to : $outFile" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan

Disconnect-ExchangeOnline -Confirm:$false | Out-Null

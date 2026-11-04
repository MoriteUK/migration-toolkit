<#
.SYNOPSIS
    Exports the LegacyExchangeDN and every existing X500 proxy address for all recipients in the
    SOURCE Exchange Online tenant to CSV, ready to stamp onto the matching objects in the target.

.DESCRIPTION
    For a tenant-to-tenant migration, the source LegacyExchangeDN has to be added to the target
    object as an X500 address (along with any X500s the source already carries) so cached
    Outlook autocomplete entries, old meeting invites and replies keep resolving.

    One row per recipient. The X500Addresses column is a semicolon-separated list that already
    includes the LegacyExchangeDN converted to "X500:<LegacyExchangeDN>", de-duplicated.

    The TargetIdentity column is left blank. Fill it in if the target object can't be found by
    PrimarySmtpAddress / UPN (for example when the domain changes and you aren't using
    -SourceDomain/-TargetDomain on the import script).

.PARAMETER RecipientTypes
    RecipientTypeDetails to export. Defaults to user, shared, room and equipment mailboxes.

.PARAMETER IncludeGroups
    Also export distribution groups, mail-enabled security groups and Microsoft 365 groups.

.PARAMETER IncludeMailUsers
    Also export MailUser objects (e.g. users already migrated / hybrid remote mailboxes).

.PARAMETER OutputPath
    CSV path. Default: .\SourceX500_<domain>.csv (when filtering) or .\SourceX500.csv

.PARAMETER DomainFilter
    Only export recipients with PrimarySmtpAddress or UserPrincipalName matching this domain suffix (e.g., 'celcat.com').
    Matches are case-insensitive.

.PARAMETER SkipConnect
    Don't call Connect-ExchangeOnline (use an existing session).

.EXAMPLE
    .\Export-SourceX500.ps1
    # Exports to SourceX500.csv

.EXAMPLE
    .\Export-SourceX500.ps1 -DomainFilter celcat.com
    # Exports to SourceX500_celcat.com.csv

.EXAMPLE
    .\Export-SourceX500.ps1 -DomainFilter celcat.com -IncludeGroups -IncludeMailUsers
    # Exports celcat.com mailboxes, groups, and mail users to SourceX500_celcat.com.csv

.EXAMPLE
    .\Export-SourceX500.ps1 -IncludeGroups -IncludeMailUsers -OutputPath C:\Temp\src_x500.csv
    # Custom output path
#>
[CmdletBinding()]
param(
    [string[]]$RecipientTypes = @('UserMailbox', 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox'),
    [switch]$IncludeGroups,
    [switch]$IncludeMailUsers,
    [string]$DomainFilter,
    [string]$OutputPath,
    [switch]$SkipConnect
)

Import-Module ExchangeOnlineManagement -ErrorAction Stop

# Set default output path - include domain in filename if filtering
if (-not $OutputPath) {
    if ($DomainFilter) {
        $cleanDomain = $DomainFilter.TrimStart('@')
        $OutputPath = ".\SourceX500_$cleanDomain.csv"
    } else {
        $OutputPath = ".\SourceX500.csv"
    }
}

if (-not $SkipConnect) {
    Write-Host 'Connecting to SOURCE Exchange Online...' -ForegroundColor Cyan
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
}

# Filter function - checks if recipient matches domain filter
function Test-DomainMatch {
    param($Recipient)

    if (-not $DomainFilter) { return $true }

    $domain = $DomainFilter.TrimStart('@').ToLowerInvariant()

    # Check PrimarySmtpAddress
    if ($Recipient.PrimarySmtpAddress -and $Recipient.PrimarySmtpAddress.ToString().ToLowerInvariant().EndsWith("@$domain")) {
        return $true
    }

    # Check UserPrincipalName
    if ($Recipient.UserPrincipalName -and $Recipient.UserPrincipalName.ToLowerInvariant().EndsWith("@$domain")) {
        return $true
    }

    return $false
}

# Builds a row with the LegacyExchangeDN + any existing X500s, de-duplicated (case-insensitive)
function New-X500Row {
    param($Recipient, [string]$ObjectType)

    $x500 = New-Object System.Collections.Generic.List[string]
    $seen = New-Object System.Collections.Generic.HashSet[string]([StringComparer]::OrdinalIgnoreCase)

    if ($Recipient.LegacyExchangeDN) {
        $val = "X500:$($Recipient.LegacyExchangeDN)"
        if ($seen.Add($val)) { $x500.Add($val) }
    }
    foreach ($addr in @($Recipient.EmailAddresses)) {
        if ($addr -like 'x500:*') {
            # Normalise the prefix so case differences don't create duplicates
            $val = 'X500:' + $addr.Substring(5)
            if ($seen.Add($val)) { $x500.Add($val) }
        }
    }

    [pscustomobject][ordered]@{
        DisplayName          = $Recipient.DisplayName
        ObjectType           = $ObjectType
        RecipientTypeDetails = $Recipient.RecipientTypeDetails
        PrimarySmtpAddress   = $Recipient.PrimarySmtpAddress
        UserPrincipalName    = $Recipient.UserPrincipalName
        Alias                = $Recipient.Alias
        LegacyExchangeDN     = $Recipient.LegacyExchangeDN
        X500Count            = $x500.Count
        X500Addresses        = $x500 -join ';'
        TargetIdentity       = ''
    }
}

$results = New-Object System.Collections.Generic.List[object]

if ($DomainFilter) {
    Write-Host "Domain filter: @$($DomainFilter.TrimStart('@'))" -ForegroundColor Yellow
}

Write-Host "Reading mailboxes ($($RecipientTypes -join ', '))..." -ForegroundColor Cyan
$mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails $RecipientTypes `
        -Properties LegacyExchangeDN, EmailAddresses, DisplayName, Alias, UserPrincipalName, PrimarySmtpAddress)
Write-Host "  Retrieved $($mailboxes.Count) mailbox(es)."

# Apply domain filter
$filteredMailboxes = @($mailboxes | Where-Object { Test-DomainMatch -Recipient $_ })
if ($DomainFilter) {
    Write-Host "  After domain filter: $($filteredMailboxes.Count) mailbox(es)." -ForegroundColor Yellow
}

foreach ($m in $filteredMailboxes) { $results.Add((New-X500Row -Recipient $m -ObjectType 'Mailbox')) }

if ($IncludeMailUsers) {
    Write-Host 'Reading mail users...' -ForegroundColor Cyan
    $mailUsers = @(Get-MailUser -ResultSize Unlimited)
    Write-Host "  Retrieved $($mailUsers.Count) mail user(s)."

    # Apply domain filter
    $filteredMailUsers = @($mailUsers | Where-Object { Test-DomainMatch -Recipient $_ })
    if ($DomainFilter) {
        Write-Host "  After domain filter: $($filteredMailUsers.Count) mail user(s)." -ForegroundColor Yellow
    }

    foreach ($mu in $filteredMailUsers) { $results.Add((New-X500Row -Recipient $mu -ObjectType 'MailUser')) }
}

if ($IncludeGroups) {
    Write-Host 'Reading distribution / mail-enabled security groups...' -ForegroundColor Cyan
    $dgs = @(Get-DistributionGroup -ResultSize Unlimited)
    Write-Host "  Retrieved $($dgs.Count) group(s)."

    # Apply domain filter
    $filteredDgs = @($dgs | Where-Object { Test-DomainMatch -Recipient $_ })
    if ($DomainFilter) {
        Write-Host "  After domain filter: $($filteredDgs.Count) group(s)." -ForegroundColor Yellow
    }

    foreach ($g in $filteredDgs) { $results.Add((New-X500Row -Recipient $g -ObjectType 'DistributionGroup')) }

    Write-Host 'Reading Microsoft 365 groups...' -ForegroundColor Cyan
    $ugs = @(Get-UnifiedGroup -ResultSize Unlimited)
    Write-Host "  Retrieved $($ugs.Count) group(s)."

    # Apply domain filter
    $filteredUgs = @($ugs | Where-Object { Test-DomainMatch -Recipient $_ })
    if ($DomainFilter) {
        Write-Host "  After domain filter: $($filteredUgs.Count) group(s)." -ForegroundColor Yellow
    }

    foreach ($g in $filteredUgs) { $results.Add((New-X500Row -Recipient $g -ObjectType 'UnifiedGroup')) }
}

if ($results.Count -eq 0) {
    Write-Warning 'No recipients found. Nothing exported.'
    return
}

$results | Sort-Object ObjectType, DisplayName |
    Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

$noLegDN = @($results | Where-Object { -not $_.LegacyExchangeDN }).Count
Write-Host "Exported $($results.Count) row(s) to $((Resolve-Path $OutputPath).Path)" -ForegroundColor Green
if ($noLegDN) { Write-Warning "$noLegDN row(s) had no LegacyExchangeDN." }

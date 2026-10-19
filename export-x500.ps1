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
    CSV path. Default: .\SourceX500_<timestamp>.csv

.PARAMETER SkipConnect
    Don't call Connect-ExchangeOnline (use an existing session).

.EXAMPLE
    .\Export-SourceX500.ps1
.EXAMPLE
    .\Export-SourceX500.ps1 -IncludeGroups -IncludeMailUsers -OutputPath C:\Temp\src_x500.csv
#>
[CmdletBinding()]
param(
    [string[]]$RecipientTypes = @('UserMailbox', 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox'),
    [switch]$IncludeGroups,
    [switch]$IncludeMailUsers,
    [string]$OutputPath = ".\SourceX500_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",
    [switch]$SkipConnect
)

Import-Module ExchangeOnlineManagement -ErrorAction Stop

if (-not $SkipConnect) {
    Write-Host 'Connecting to SOURCE Exchange Online...' -ForegroundColor Cyan
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
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

Write-Host "Reading mailboxes ($($RecipientTypes -join ', '))..." -ForegroundColor Cyan
$mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails $RecipientTypes `
        -Properties LegacyExchangeDN, EmailAddresses, DisplayName, Alias, UserPrincipalName, PrimarySmtpAddress)
Write-Host "  $($mailboxes.Count) mailbox(es)."
foreach ($m in $mailboxes) { $results.Add((New-X500Row -Recipient $m -ObjectType 'Mailbox')) }

if ($IncludeMailUsers) {
    Write-Host 'Reading mail users...' -ForegroundColor Cyan
    $mailUsers = @(Get-MailUser -ResultSize Unlimited)
    Write-Host "  $($mailUsers.Count) mail user(s)."
    foreach ($mu in $mailUsers) { $results.Add((New-X500Row -Recipient $mu -ObjectType 'MailUser')) }
}

if ($IncludeGroups) {
    Write-Host 'Reading distribution / mail-enabled security groups...' -ForegroundColor Cyan
    $dgs = @(Get-DistributionGroup -ResultSize Unlimited)
    Write-Host "  $($dgs.Count) group(s)."
    foreach ($g in $dgs) { $results.Add((New-X500Row -Recipient $g -ObjectType 'DistributionGroup')) }

    Write-Host 'Reading Microsoft 365 groups...' -ForegroundColor Cyan
    $ugs = @(Get-UnifiedGroup -ResultSize Unlimited)
    Write-Host "  $($ugs.Count) group(s)."
    foreach ($g in $ugs) { $results.Add((New-X500Row -Recipient $g -ObjectType 'UnifiedGroup')) }
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

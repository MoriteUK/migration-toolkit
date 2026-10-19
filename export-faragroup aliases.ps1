<#
.SYNOPSIS
    Exports all secondary SMTP aliases on the Fara domains for groups to a CSV, one alias per line.

.DESCRIPTION
    Connects to Exchange Online, finds every group that has an address on one of the
    Fara domains, and writes one row per secondary SMTP alias:

        PrimarySmtpAddress,Alias,GroupType
        finance@fara.co.uk,accounts@fara.co.uk,MailUniversalDistributionGroup
        projects@fara.co.uk,projects@faragroup.com,GroupMailbox

    GroupType tells you which cmdlet to use when adding the alias back:
        MailUniversalDistributionGroup / MailUniversalSecurityGroup / RoomList -> Set-DistributionGroup
        DynamicDistributionGroup                                               -> Set-DynamicDistributionGroup
        GroupMailbox (Microsoft 365 group)                                     -> Set-UnifiedGroup

    The primary address itself is not repeated as an alias. X500, SPO and
    onmicrosoft.com addresses are excluded.

.PARAMETER Domains
    Exact domain names to export. If omitted, every accepted domain containing "fara"
    is used (the script shows which ones it picked).

.PARAMETER OutputPath
    CSV path. Defaults to .\FaraGroupAliases_<timestamp>.csv

.EXAMPLE
    .\Export-FaraGroupAliases.ps1

.EXAMPLE
    .\Export-FaraGroupAliases.ps1 -Domains fara.co.uk,faragroup.com
#>
[CmdletBinding()]
param(
    [string[]]$Domains,
    [string]$OutputPath = ".\FaraGroupAliases_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

$ErrorActionPreference = 'Stop'

$groupTypes = @(
    'MailUniversalDistributionGroup',
    'MailUniversalSecurityGroup',
    'MailNonUniversalGroup',
    'RoomList',
    'DynamicDistributionGroup',
    'GroupMailbox'
)

if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    throw "ExchangeOnlineManagement module not found. Install-Module ExchangeOnlineManagement -Scope CurrentUser"
}

if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
    Connect-ExchangeOnline -ShowBanner:$false
}

# Resolve the Fara domains
if (-not $Domains) {
    $Domains = Get-AcceptedDomain | Where-Object { $_.DomainName -like '*fara*' } |
        Select-Object -ExpandProperty DomainName
    if (-not $Domains) { throw "No accepted domains matching '*fara*' found. Use -Domains to specify them." }
}
$Domains = $Domains | ForEach-Object { $_.ToLower().Trim() }
Write-Host "Domains: $($Domains -join ', ')" -ForegroundColor Cyan

Write-Host "Retrieving groups..." -ForegroundColor Cyan
$groups = Get-EXORecipient -ResultSize Unlimited -RecipientTypeDetails $groupTypes `
    -Properties EmailAddresses, PrimarySmtpAddress, RecipientTypeDetails
Write-Host "  $($groups.Count) groups found"

$results = foreach ($g in $groups) {
    $primary = $g.PrimarySmtpAddress.ToString().ToLower()

    foreach ($addr in $g.EmailAddresses) {
        # Secondary SMTP addresses only - lowercase "smtp:" prefix (primary is "SMTP:")
        if ($addr -cnotlike 'smtp:*') { continue }

        $alias = $addr.Substring(5).ToLower()
        if ($alias -eq $primary) { continue }

        $aliasDomain = $alias.Split('@')[-1]
        if ($Domains -notcontains $aliasDomain) { continue }

        [PSCustomObject]@{
            PrimarySmtpAddress = $primary
            Alias              = $alias
            GroupType          = $g.RecipientTypeDetails.ToString()
        }
    }
}

$results = $results | Sort-Object PrimarySmtpAddress, Alias -Unique

$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

$groupCount = ($results | Select-Object -ExpandProperty PrimarySmtpAddress -Unique).Count
Write-Host "Exported $($results.Count) aliases for $groupCount groups to $OutputPath" -ForegroundColor Green
$results | Group-Object GroupType | ForEach-Object { Write-Host "  $($_.Name): $($_.Count)" }

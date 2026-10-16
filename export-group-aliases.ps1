<#
.SYNOPSIS
    Exports all secondary SMTP aliases for groups on specified domains to a CSV, one alias per line.

.DESCRIPTION
    Connects to Exchange Online, finds every group that has an address on the specified
    domains (or search pattern), and writes one row per secondary SMTP alias:

        PrimarySmtpAddress,Alias,GroupType
        finance@example.com,accounts@example.com,MailUniversalDistributionGroup
        projects@example.com,projects@examplegroup.com,GroupMailbox

    GroupType tells you which cmdlet to use when adding the alias back:
        MailUniversalDistributionGroup / MailUniversalSecurityGroup / RoomList -> Set-DistributionGroup
        DynamicDistributionGroup                                               -> Set-DynamicDistributionGroup
        GroupMailbox (Microsoft 365 group)                                     -> Set-UnifiedGroup

    The primary address itself is not repeated as an alias. X500, SPO and
    onmicrosoft.com addresses are excluded.

.PARAMETER Domains
    Exact domain names to export (comma-separated). If omitted and -SearchPattern is not provided,
    you'll be prompted to enter domain names.

.PARAMETER SearchPattern
    Wildcard pattern to auto-detect domains from accepted domains list (e.g., "*example*").
    Used only if -Domains is not specified.

.PARAMETER OutputPath
    CSV path. Defaults to .\GroupAliases_<timestamp>.csv

.EXAMPLE
    .\Export-Group-Aliases.ps1 -Domains example.com,examplegroup.com

.EXAMPLE
    .\Export-Group-Aliases.ps1 -SearchPattern "*example*"

.EXAMPLE
    .\Export-Group-Aliases.ps1  (prompts for domain names)
#>
[CmdletBinding()]
param(
    [string[]]$Domains,
    [string]$SearchPattern,
    [string]$OutputPath = ".\GroupAliases_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
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

# Resolve domains
if (-not $Domains) {
    if ($SearchPattern) {
        $Domains = Get-AcceptedDomain | Where-Object { $_.DomainName -like $SearchPattern } |
            Select-Object -ExpandProperty DomainName
        if (-not $Domains) {
            throw "No accepted domains matching '$SearchPattern' found. Use -Domains to specify them explicitly."
        }
        Write-Host "Auto-detected domains matching '$SearchPattern':" -ForegroundColor Yellow
    } else {
        # Prompt for domains if neither -Domains nor -SearchPattern provided
        $domainInput = Read-Host "Enter domain names to export (comma-separated, e.g., example.com,example.co.uk)"
        if (-not $domainInput) { throw "No domains specified. Use -Domains or -SearchPattern parameters." }
        $Domains = $domainInput -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        if (-not $Domains) { throw "No valid domains specified." }
    }
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

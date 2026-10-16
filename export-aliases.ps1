<#
.SYNOPSIS
    Exports all secondary SMTP aliases for specified domains to a CSV, one alias per line.

.DESCRIPTION
    Connects to Exchange Online, finds every recipient that has an address on the specified
    domains (or search pattern), and writes one row per secondary SMTP alias:

        PrimarySmtpAddress,Alias
        jane.doe@example.com,jdoe@example.com
        jane.doe@example.com,jane@examplegroup.com

    The primary address itself is not repeated as an alias. X500, SIP, SPO and
    onmicrosoft.com addresses are excluded, so the output can be fed straight into
    Set-Mailbox -EmailAddresses @{Add=...} in the new tenant.

.PARAMETER Domains
    Exact domain names to export (comma-separated). If omitted and -SearchPattern is not provided,
    you'll be prompted to enter domain names.

.PARAMETER SearchPattern
    Wildcard pattern to auto-detect domains from accepted domains list (e.g., "*example*").
    Used only if -Domains is not specified.

.PARAMETER RecipientTypes
    Recipient types to include. Defaults to user and shared mailboxes.

.PARAMETER OutputPath
    CSV path. Defaults to .\EmailAliases_<timestamp>.csv

.EXAMPLE
    .\Export-Aliases.ps1 -Domains example.com,examplegroup.com

.EXAMPLE
    .\Export-Aliases.ps1 -SearchPattern "*example*"

.EXAMPLE
    .\Export-Aliases.ps1 -Domains example.com -RecipientTypes UserMailbox,SharedMailbox,MailUser
#>
[CmdletBinding()]
param(
    [string[]]$Domains,
    [string]$SearchPattern,
    [string[]]$RecipientTypes = @('UserMailbox', 'SharedMailbox'),
    [string]$OutputPath = ".\EmailAliases_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

$ErrorActionPreference = 'Stop'

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

Write-Host "Retrieving recipients ($($RecipientTypes -join ', '))..." -ForegroundColor Cyan
$recipients = Get-EXORecipient -ResultSize Unlimited -RecipientTypeDetails $RecipientTypes `
    -Properties EmailAddresses, PrimarySmtpAddress
Write-Host "  $($recipients.Count) recipients found"

$results = foreach ($r in $recipients) {
    $primary = $r.PrimarySmtpAddress.ToString().ToLower()

    foreach ($addr in $r.EmailAddresses) {
        # Secondary SMTP addresses only - lowercase "smtp:" prefix (primary is "SMTP:")
        if ($addr -cnotlike 'smtp:*') { continue }

        $alias = $addr.Substring(5).ToLower()
        if ($alias -eq $primary) { continue }

        $aliasDomain = $alias.Split('@')[-1]
        if ($Domains -notcontains $aliasDomain) { continue }

        [PSCustomObject]@{
            PrimarySmtpAddress = $primary
            Alias              = $alias
        }
    }
}

$results = $results | Sort-Object PrimarySmtpAddress, Alias -Unique

$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

$userCount = ($results | Select-Object -ExpandProperty PrimarySmtpAddress -Unique).Count
Write-Host "Exported $($results.Count) aliases for $userCount recipients to $OutputPath" -ForegroundColor Green

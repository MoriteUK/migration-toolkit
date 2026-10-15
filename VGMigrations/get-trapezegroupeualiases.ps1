<#
.SYNOPSIS
    Exports all recipients (users/mailboxes) that have an SMTP alias on @trapezegroup.eu.

.DESCRIPTION
    Connects to Exchange Online, retrieves all recipients, filters those with at least one
    SMTP proxy address ending in the target domain, and exports the results to CSV.

.EXAMPLE
    .\Get-TrapezeGroupEuAliases.ps1
    .\Get-TrapezeGroupEuAliases.ps1 -Domain "trapezegroup.eu" -OutputPath "C:\Temp\TrapezeEU.csv"
    .\Get-TrapezeGroupEuAliases.ps1 -IncludeAllRecipientTypes
#>
[CmdletBinding()]
param(
    [string]$Domain = "trapezegroup.eu",
    [string]$OutputPath = ".\TrapezeGroupEU_Aliases_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",
    # By default only user/shared/room/equipment mailboxes and mail users are returned.
    # Use this switch to include groups, contacts, public folders etc.
    [switch]$IncludeAllRecipientTypes
)

if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    throw "ExchangeOnlineManagement module not installed. Run: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
}

if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
    Connect-ExchangeOnline -ShowBanner:$false
}

$domainSuffix = "@$Domain"

$recipientParams = @{
    ResultSize = 'Unlimited'
    Properties = 'EmailAddresses', 'RecipientTypeDetails', 'PrimarySmtpAddress', 'DisplayName', 'Alias', 'ExternalDirectoryObjectId'
}
if (-not $IncludeAllRecipientTypes) {
    $recipientParams.RecipientTypeDetails = 'UserMailbox', 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox', 'MailUser'
}

Write-Host "Retrieving recipients from Exchange Online..." -ForegroundColor Cyan
$recipients = Get-EXORecipient @recipientParams

Write-Host "Checking $($recipients.Count) recipients for $domainSuffix addresses..." -ForegroundColor Cyan

$results = foreach ($r in $recipients) {
    $matchingSmtp = $r.EmailAddresses |
        Where-Object { $_ -like 'smtp:*' -and $_ -like "*$domainSuffix" } |
        ForEach-Object { $_ -replace '^smtp:', '' }

    if ($matchingSmtp) {
        $allSmtp = $r.EmailAddresses |
            Where-Object { $_ -like 'smtp:*' } |
            ForEach-Object { $_ -replace '^smtp:', '' }

        [PSCustomObject]@{
            DisplayName          = $r.DisplayName
            Alias                = $r.Alias
            RecipientTypeDetails = $r.RecipientTypeDetails
            PrimarySmtpAddress   = $r.PrimarySmtpAddress
            IsPrimaryOnDomain    = $r.PrimarySmtpAddress -like "*$domainSuffix"
            TrapezeEUAddresses   = ($matchingSmtp -join '; ')
            AllSmtpAddresses     = ($allSmtp -join '; ')
            ObjectId             = $r.ExternalDirectoryObjectId
        }
    }
}

$results = @($results | Sort-Object DisplayName)
$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

Write-Host "Found $($results.Count) recipients with a $domainSuffix address." -ForegroundColor Green
Write-Host "Exported to: $((Resolve-Path $OutputPath).Path)" -ForegroundColor Green
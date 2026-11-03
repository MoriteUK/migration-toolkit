<#
.SYNOPSIS
    Exports distribution lists for a specific domain and their members.

.DESCRIPTION
    Produces two CSVs:
      - DistributionLists_<timestamp>.csv  : one row per DL (incl. mail-enabled security groups
                                             and, optionally, dynamic DLs)
      - DistributionListMembers_<timestamp>.csv : one row per DL/member pair

.PARAMETER Domain
    The domain name to filter distribution lists (e.g., 'contoso.com'). Required.

.PARAMETER OutputPath
    Folder for the CSVs. Defaults to the script folder.

.PARAMETER IncludeDynamic
    Also export dynamic distribution groups and their calculated membership.

.EXAMPLE
    .\Export-DistributionLists.ps1 -Domain 'contoso.com'

.EXAMPLE
    .\Export-DistributionLists.ps1 -Domain 'contoso.com' -IncludeDynamic
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Domain,

    [string]$OutputPath = $PSScriptRoot,
    [switch]$IncludeDynamic
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    throw "ExchangeOnlineManagement module not found. Install with: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
}

if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
    Connect-ExchangeOnline -ShowBanner:$false
}

$timestamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
$dlFile      = Join-Path $OutputPath "DistributionLists_${Domain}_$timestamp.csv"
$memberFile  = Join-Path $OutputPath "DistributionListMembers_${Domain}_$timestamp.csv"

$dlReport     = [System.Collections.Generic.List[object]]::new()
$memberReport = [System.Collections.Generic.List[object]]::new()

Write-Host "Retrieving distribution groups for domain: $Domain..." -ForegroundColor Cyan
$groups = @(Get-DistributionGroup -ResultSize Unlimited | Where-Object { $_.PrimarySmtpAddress -like "*@$Domain" })
if ($IncludeDynamic) {
    Write-Host "Retrieving dynamic distribution groups for domain: $Domain..." -ForegroundColor Cyan
    $groups += @(Get-DynamicDistributionGroup -ResultSize Unlimited | Where-Object { $_.PrimarySmtpAddress -like "*@$Domain" })
}
Write-Host "Found $($groups.Count) groups for domain $Domain." -ForegroundColor Green

if ($groups.Count -eq 0) {
    Write-Warning "No distribution groups found for domain: $Domain"
    return
}

$i = 0
foreach ($group in $groups) {
    $i++
    Write-Progress -Activity "Processing distribution lists" -Status $group.DisplayName -PercentComplete (($i / $groups.Count) * 100)

    $isDynamic = $group.RecipientTypeDetails -eq 'DynamicDistributionGroup'
    try {
        $members = if ($isDynamic) {
            @(Get-DynamicDistributionGroupMember -Identity $group.Identity -ResultSize Unlimited)
        } else {
            @(Get-DistributionGroupMember -Identity $group.Identity -ResultSize Unlimited)
        }
    } catch {
        Write-Warning "Failed to get members for $($group.PrimarySmtpAddress): $($_.Exception.Message)"
        $members = @()
    }

    $dlReport.Add([pscustomobject]@{
        DisplayName                        = $group.DisplayName
        PrimarySmtpAddress                 = $group.PrimarySmtpAddress
        Alias                              = $group.Alias
        GroupType                          = $group.RecipientTypeDetails
        MemberCount                        = $members.Count
        ManagedBy                          = ($group.ManagedBy -join '; ')
        RequireSenderAuthenticationEnabled = $group.RequireSenderAuthenticationEnabled
        HiddenFromAddressListsEnabled      = $group.HiddenFromAddressListsEnabled
        IsDirSynced                        = $group.IsDirSynced
        RecipientFilter                    = if ($isDynamic) { $group.RecipientFilter } else { $null }
        EmailAddresses                     = (($group.EmailAddresses | Where-Object { $_ -like 'smtp:*' }) -replace '^smtp:', '' -join '; ')
        WhenCreated                        = $group.WhenCreated
    })

    foreach ($m in $members) {
        $memberReport.Add([pscustomobject]@{
            GroupDisplayName   = $group.DisplayName
            GroupEmail         = $group.PrimarySmtpAddress
            GroupType          = $group.RecipientTypeDetails
            MemberDisplayName  = $m.DisplayName
            MemberEmail        = $m.PrimarySmtpAddress
            MemberType         = $m.RecipientTypeDetails
        })
    }
}
Write-Progress -Activity "Processing distribution lists" -Completed

$dlReport     | Export-Csv -Path $dlFile     -NoTypeInformation -Encoding UTF8
$memberReport | Export-Csv -Path $memberFile -NoTypeInformation -Encoding UTF8

Write-Host "`nExported $($dlReport.Count) groups      -> $dlFile" -ForegroundColor Green
Write-Host "Exported $($memberReport.Count) memberships -> $memberFile" -ForegroundColor Green

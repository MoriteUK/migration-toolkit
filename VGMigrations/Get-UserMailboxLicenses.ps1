<#
.SYNOPSIS
    Reports the mailbox-relevant Microsoft 365 licenses for a set of users in a
    source tenant, and optionally generates an AvePoint Fly Exchange mapping CSV.

.PARAMETER CsvPath
    Path to the input CSV. Required column: SourceUPN (or UserPrincipalName / Email).
    Optional column: DestinationUPN — the UPN the user will have in the destination
    tenant. If omitted the script assumes the UPN is identical in both tenants
    (same-domain or same-address moves only).

.PARAMETER OutputPath
    Path for the license report CSV. Defaults to license-report.csv next to the input file.

.PARAMETER FlyMappingPath
    If supplied, also writes an AvePoint Fly Exchange mapping CSV to this path,
    ready to load directly into the runner / Import-FlyMappings.ps1.
    If omitted, the file is written as exchange-mappings.csv next to the report.
    Pass -NoFlyMapping to suppress it entirely.

.PARAMETER NoFlyMapping
    Suppress generation of the Fly Exchange mapping CSV.

.PARAMETER MailboxType
    The Fly mailbox type to stamp on every row of the mapping CSV.
    Default: 'User mailbox'. Other values: 'Shared mailbox', 'Resource mailbox', etc.

.PARAMETER TenantId
    Optional. Source tenant ID or domain (e.g. contoso.onmicrosoft.com).
    If omitted you will be prompted to sign in interactively.

.EXAMPLE
    .\Get-UserMailboxLicenses.ps1 -CsvPath C:\users.csv
    .\Get-UserMailboxLicenses.ps1 -CsvPath C:\users.csv -TenantId contoso.onmicrosoft.com
    .\Get-UserMailboxLicenses.ps1 -CsvPath C:\users.csv -NoFlyMapping
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$CsvPath,

    [string]$OutputPath,

    [string]$FlyMappingPath,

    [switch]$NoFlyMapping,

    [string]$MailboxType = 'User mailbox',

    [string]$TenantId
)

#region --- Mailbox SKU definitions ------------------------------------------

$MailboxSkus = @(
    'EXCHANGESTANDARD',          # Exchange Online Plan 1
    'EXCHANGEENTERPRISE',        # Exchange Online Plan 2
    'EXCHANGEARCHIVE_ADDON',
    'EXCHANGEARCHIVE',
    'O365_BUSINESS_ESSENTIALS',  # M365 Business Basic
    'O365_BUSINESS_PREMIUM',     # M365 Business Standard
    'SPB',                       # M365 Business Premium
    'STANDARDPACK',              # Office 365 E1
    'STANDARDWOFFPACK',
    'ENTERPRISEPACK',            # M365 E3 / Office 365 E3
    'ENTERPRISEPREMIUM',         # M365 E5
    'ENTERPRISEPREMIUM_NOPSTNCONF',
    'ENTERPRISEWITHSCAL',
    'DESKLESSPACK',
    'M365_F1',
    'SPE_E3',
    'SPE_E5'
)

$FriendlyNames = @{
    'EXCHANGESTANDARD'             = 'Exchange Online Plan 1'
    'EXCHANGEENTERPRISE'           = 'Exchange Online Plan 2'
    'EXCHANGEARCHIVE_ADDON'        = 'Exchange Online Archiving Add-on'
    'EXCHANGEARCHIVE'              = 'Exchange Online Archiving'
    'O365_BUSINESS_ESSENTIALS'     = 'Microsoft 365 Business Basic'
    'O365_BUSINESS_PREMIUM'        = 'Microsoft 365 Business Standard'
    'SPB'                          = 'Microsoft 365 Business Premium'
    'STANDARDPACK'                 = 'Office 365 E1'
    'STANDARDWOFFPACK'             = 'Office 365 E1 (no Teams)'
    'ENTERPRISEPACK'               = 'Microsoft 365 E3'
    'ENTERPRISEPREMIUM'            = 'Microsoft 365 E5'
    'ENTERPRISEPREMIUM_NOPSTNCONF' = 'Microsoft 365 E5 (no Audio Conf)'
    'ENTERPRISEWITHSCAL'           = 'Office 365 E4'
    'DESKLESSPACK'                 = 'Microsoft 365 F1'
    'M365_F1'                      = 'Microsoft 365 F1'
    'SPE_E3'                       = 'Microsoft 365 E3'
    'SPE_E5'                       = 'Microsoft 365 E5'
}

function Get-FriendlyName([string]$SkuPartNumber) {
    if ($FriendlyNames.ContainsKey($SkuPartNumber)) { return $FriendlyNames[$SkuPartNumber] }
    return $SkuPartNumber
}

#endregion

#region --- Prerequisites ----------------------------------------------------

if (-not (Get-Module Microsoft.Graph.Users -ListAvailable) -or
    -not (Get-Module Microsoft.Graph.Identity.DirectoryManagement -ListAvailable)) {
    Write-Host "Installing Microsoft.Graph modules..." -ForegroundColor Cyan
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
}

#endregion

#region --- Connect ----------------------------------------------------------

$connectParams = @{ Scopes = @('User.Read.All', 'Organization.Read.All') }
if ($TenantId) { $connectParams['TenantId'] = $TenantId }

Write-Host "Connecting to Microsoft Graph (source tenant)..." -ForegroundColor Cyan
Connect-MgGraph @connectParams -NoWelcome

Write-Host "Loading tenant SKU catalogue..." -ForegroundColor Cyan
$skuLookup = @{}
Get-MgSubscribedSku -All | ForEach-Object {
    $skuLookup[$_.SkuId] = $_.SkuPartNumber
}

#endregion

#region --- Read input -------------------------------------------------------

if (-not (Test-Path $CsvPath)) {
    Write-Error "CSV not found: $CsvPath"
    exit 1
}

$users = Import-Csv $CsvPath

# Accept several common column names for the source UPN
$srcCol = $users[0].PSObject.Properties.Name |
    Where-Object { $_ -iin @('SourceUPN', 'UserPrincipalName', 'Email', 'UPN') } |
    Select-Object -First 1

if (-not $srcCol) {
    Write-Error "CSV must have a 'SourceUPN', 'UserPrincipalName', or 'Email' column."
    exit 1
}

# DestinationUPN column is optional — fall back to source UPN when absent
$dstColPresent = $users[0].PSObject.Properties.Name -icontains 'DestinationUPN'

Write-Host "$($users.Count) users loaded. Source column: '$srcCol'$(if ($dstColPresent) { ", Destination column: 'DestinationUPN'" } else { " (no DestinationUPN column — will mirror source UPN)" })" -ForegroundColor Cyan

#endregion

#region --- Query Graph ------------------------------------------------------

$report  = [System.Collections.Generic.List[pscustomobject]]::new()
$i = 0

foreach ($row in $users) {
    $i++
    $srcUpn = $row.$srcCol.Trim()
    $dstUpn = if ($dstColPresent -and -not [string]::IsNullOrWhiteSpace($row.DestinationUPN)) {
                  $row.DestinationUPN.Trim()
              } else {
                  $srcUpn
              }

    Write-Progress -Activity "Querying licenses" -Status $srcUpn -PercentComplete (($i / $users.Count) * 100)

    try {
        $mgUser = Get-MgUser -UserId $srcUpn `
            -Property 'DisplayName,UserPrincipalName,AssignedLicenses,AccountEnabled' `
            -ErrorAction Stop
    }
    catch {
        $report.Add([pscustomobject]@{
            SourceUPN            = $srcUpn
            DestinationUPN       = $dstUpn
            DisplayName          = ''
            AccountEnabled       = ''
            AllAssignedSkus      = ''
            MailboxLicenses      = ''
            MailboxLicenseCodes  = ''
            HasMailbox           = 'ERROR'
            SuggestedTarget      = ''
            Error                = $_.Exception.Message
        })
        continue
    }

    $allSkus  = $mgUser.AssignedLicenses | ForEach-Object { $skuLookup[$_.SkuId] ?? $_.SkuId }
    $mboxSkus = $allSkus | Where-Object { $MailboxSkus -contains $_ }

    $suggested = if ($mboxSkus -match 'ENTERPRISEPREMIUM|SPE_E5|SPB') { 'BusinessPremium' }
                 elseif ($mboxSkus)                                    { 'BusinessBasic' }
                 else                                                   { 'None' }

    $report.Add([pscustomobject]@{
        SourceUPN            = $mgUser.UserPrincipalName
        DestinationUPN       = $dstUpn
        DisplayName          = $mgUser.DisplayName
        AccountEnabled       = $mgUser.AccountEnabled
        AllAssignedSkus      = ($allSkus -join ' | ')
        MailboxLicenses      = ($mboxSkus | ForEach-Object { Get-FriendlyName $_ }) -join ' | '
        MailboxLicenseCodes  = ($mboxSkus -join ' | ')
        HasMailbox           = ($mboxSkus.Count -gt 0).ToString()
        SuggestedTarget      = $suggested
        Error                = ''
    })
}

Write-Progress -Activity "Querying licenses" -Completed

#endregion

#region --- Write license report ---------------------------------------------

$baseDir = [System.IO.Path]::GetDirectoryName((Resolve-Path $CsvPath))

if (-not $OutputPath) {
    $OutputPath = [System.IO.Path]::Combine($baseDir, 'license-report.csv')
}

$report | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "`nLicense report: $OutputPath" -ForegroundColor Green

#endregion

#region --- Write Fly Exchange mapping CSV -----------------------------------

if (-not $NoFlyMapping) {
    if (-not $FlyMappingPath) {
        $FlyMappingPath = [System.IO.Path]::Combine($baseDir, 'exchange-mappings.csv')
    }

    $flyRows = $report |
        Where-Object { $_.HasMailbox -eq 'True' } |
        ForEach-Object {
            [pscustomobject]@{
                'Source'           = $_.SourceUPN
                'Source type'      = $MailboxType
                'Destination'      = $_.DestinationUPN
                'Destination type' = $MailboxType
            }
        }

    $flyRows | Export-Csv -Path $FlyMappingPath -NoTypeInformation -Encoding utf8BOM -UseQuotes AsNeeded
    Write-Host "Fly Exchange mapping CSV ($($flyRows.Count) rows): $FlyMappingPath" -ForegroundColor Green
    Write-Host "  -> Load this file into the Exchange workload row in the Migration Runner." -ForegroundColor Gray
}

#endregion

#region --- Summary ----------------------------------------------------------

Write-Host "`nSummary:" -ForegroundColor Cyan
$report | Group-Object SuggestedTarget | Sort-Object Name | Format-Table Name, Count -AutoSize

$errors = $report | Where-Object { $_.Error }
if ($errors) {
    Write-Host "Errors ($($errors.Count)):" -ForegroundColor Red
    $errors | Format-Table SourceUPN, Error -AutoSize
}

Disconnect-MgGraph | Out-Null

#endregion

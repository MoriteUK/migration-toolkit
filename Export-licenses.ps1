<#
.SYNOPSIS
    Exports every user in a given domain with their assigned licenses, using friendly license names (e.g. "Microsoft 365 E3").

.DESCRIPTION
    - Finds users whose UserPrincipalName ends with @<Domain> (optionally also matching the primary mail address).
    - Resolves each assigned SkuId to its SkuPartNumber via Get-MgSubscribedSku.
    - Converts the SkuPartNumber to a friendly name using Microsoft's published licensing CSV
      (downloaded at runtime), falling back to a built-in table of common SKUs if the download fails.
    - Adds a short name column (E3, E5, F3, Business Premium, ...) for quick filtering.

.PARAMETER Domain
    The domain to filter on, e.g. contoso.com

.PARAMETER OutputPath
    CSV file path. Defaults to .\UserLicenses_<domain>_<timestamp>.csv

.PARAMETER IncludeMail
    Also include users whose primary Mail address is in the domain (not just UPN).

.PARAMETER OneRowPerLicense
    Output one row per user/license combination instead of one row per user.

.EXAMPLE
    .\Export-UserLicensesByDomain.ps1 -Domain contoso.com

.EXAMPLE
    .\Export-UserLicensesByDomain.ps1 -Domain contoso.com -IncludeMail -OneRowPerLicense
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$Domain,

    [string]$OutputPath,

    [switch]$IncludeMail,

    [switch]$OneRowPerLicense
)

$ErrorActionPreference = 'Stop'
$Domain = $Domain.TrimStart('@').ToLower()

if (-not $OutputPath) {
    $OutputPath = Join-Path (Get-Location) ("UserLicenses_{0}_{1}.csv" -f $Domain, (Get-Date -Format 'yyyyMMdd_HHmmss'))
}

# --- Connect to Graph ---------------------------------------------------------
foreach ($mod in 'Microsoft.Graph.Authentication', 'Microsoft.Graph.Users', 'Microsoft.Graph.Identity.DirectoryManagement') {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        throw "Module $mod is not installed. Run: Install-Module Microsoft.Graph -Scope CurrentUser"
    }
}

if (-not (Get-MgContext)) {
    Connect-MgGraph -Scopes 'User.Read.All', 'Organization.Read.All' -NoWelcome
}

# --- Friendly name lookup -----------------------------------------------------
# Fallback for when Microsoft's CSV can't be downloaded
$fallbackNames = @{
    'O365_BUSINESS_ESSENTIALS'          = 'Microsoft 365 Business Basic'
    'O365_BUSINESS_PREMIUM'             = 'Microsoft 365 Business Standard'
    'O365_BUSINESS'                     = 'Microsoft 365 Apps for Business'
    'SPB'                               = 'Microsoft 365 Business Premium'
    'STANDARDPACK'                      = 'Office 365 E1'
    'ENTERPRISEPACK'                    = 'Office 365 E3'
    'ENTERPRISEPREMIUM'                 = 'Office 365 E5'
    'ENTERPRISEPREMIUM_NOPSTNCONF'      = 'Office 365 E5 (without Audio Conferencing)'
    'SPE_E3'                            = 'Microsoft 365 E3'
    'SPE_E5'                            = 'Microsoft 365 E5'
    'SPE_F1'                            = 'Microsoft 365 F3'
    'M365_F1'                           = 'Microsoft 365 F1'
    'DESKLESSPACK'                      = 'Office 365 F3'
    'OFFICESUBSCRIPTION'                = 'Microsoft 365 Apps for Enterprise'
    'EXCHANGESTANDARD'                  = 'Exchange Online (Plan 1)'
    'EXCHANGEENTERPRISE'                = 'Exchange Online (Plan 2)'
    'EXCHANGEDESKLESS'                  = 'Exchange Online Kiosk'
    'EXCHANGEARCHIVE_ADDON'             = 'Exchange Online Archiving for Exchange Online'
    'EMS'                               = 'Enterprise Mobility + Security E3'
    'EMSPREMIUM'                        = 'Enterprise Mobility + Security E5'
    'AAD_PREMIUM'                       = 'Microsoft Entra ID P1'
    'AAD_PREMIUM_P2'                    = 'Microsoft Entra ID P2'
    'INTUNE_A'                          = 'Intune Plan 1'
    'POWER_BI_PRO'                      = 'Power BI Pro'
    'POWER_BI_STANDARD'                 = 'Power BI (free)'
    'PBI_PREMIUM_PER_USER'              = 'Power BI Premium Per User'
    'PROJECTPREMIUM'                    = 'Project Plan 5'
    'PROJECTPROFESSIONAL'               = 'Project Plan 3'
    'VISIOCLIENT'                       = 'Visio Plan 2'
    'VISIOONLINE_PLAN1'                 = 'Visio Plan 1'
    'MCOEV'                             = 'Teams Phone Standard'
    'MCOMEETADV'                        = 'Microsoft 365 Audio Conferencing'
    'MCOPSTN1'                          = 'Skype for Business PSTN Domestic Calling'
    'PHONESYSTEM_VIRTUALUSER'           = 'Teams Phone Resource Account'
    'MCOCAP'                            = 'Teams Shared Devices'
    'Microsoft_Teams_Rooms_Pro'         = 'Teams Rooms Pro'
    'Microsoft_Teams_Rooms_Basic'       = 'Teams Rooms Basic'
    'TEAMS_EXPLORATORY'                 = 'Teams Exploratory'
    'Microsoft_Teams_Premium'           = 'Teams Premium'
    'Microsoft_365_Copilot'             = 'Microsoft 365 Copilot'
    'FLOW_FREE'                         = 'Power Automate (free)'
    'POWERAPPS_VIRAL'                   = 'Power Apps Plan 2 Trial'
    'WINDOWS_STORE'                     = 'Windows Store for Business'
    'RIGHTSMANAGEMENT'                  = 'Azure Information Protection Plan 1'
    'ATP_ENTERPRISE'                    = 'Defender for Office 365 (Plan 1)'
    'THREAT_INTELLIGENCE'               = 'Defender for Office 365 (Plan 2)'
    'IDENTITY_THREAT_PROTECTION'        = 'Microsoft 365 E5 Security'
    'INFORMATION_PROTECTION_COMPLIANCE' = 'Microsoft 365 E5 Compliance'
    'Microsoft_365_E3_(no_Teams)'       = 'Microsoft 365 E3 (no Teams)'
    'Microsoft_365_E5_(no_Teams)'       = 'Microsoft 365 E5 (no Teams)'
    'O365_w/o_Teams_Bundle_M3'          = 'Office 365 E3 (no Teams)'
    'O365_w/o_Teams_Bundle_M5'          = 'Office 365 E5 (no Teams)'
}

$friendlyNames = @{}
$csvUrl = 'https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv'
try {
    Write-Host 'Downloading Microsoft license name reference...' -ForegroundColor Cyan
    $csv = Invoke-RestMethod -Uri $csvUrl -UseBasicParsing | ConvertFrom-Csv
    foreach ($row in $csv) {
        if ($row.String_Id -and -not $friendlyNames.ContainsKey($row.String_Id)) {
            $friendlyNames[$row.String_Id] = $row.Product_Display_Name
        }
    }
    Write-Host ("  Loaded {0} SKU names." -f $friendlyNames.Count) -ForegroundColor Green
}
catch {
    Write-Warning "Could not download license reference CSV ($($_.Exception.Message)). Using built-in list."
}
foreach ($k in $fallbackNames.Keys) {
    if (-not $friendlyNames.ContainsKey($k)) { $friendlyNames[$k] = $fallbackNames[$k] }
}

# Short names like "E3" derived from the friendly name
function Get-ShortLicenseName {
    param([string]$Friendly, [string]$SkuPartNumber)
    switch -Regex ($Friendly) {
        'Business Premium'          { return 'Business Premium' }
        'Business Standard'         { return 'Business Standard' }
        'Business Basic'            { return 'Business Basic' }
        'Apps for Business'         { return 'Apps for Business' }
        'Apps for Enterprise'       { return 'Apps for Enterprise' }
        '365 E5 Security'           { return 'E5 Security' }
        '365 E5 Compliance'         { return 'E5 Compliance' }
        'Mobility \+ Security E3'   { return 'EMS E3' }
        'Mobility \+ Security E5'   { return 'EMS E5' }
        '\b([EFAG]\d)\b'            { return $Matches[1] }
        'Exchange Online \(Plan 1\)' { return 'EXO P1' }
        'Exchange Online \(Plan 2\)' { return 'EXO P2' }
        'Entra ID P1'               { return 'Entra P1' }
        'Entra ID P2'               { return 'Entra P2' }
        'Copilot'                   { return 'Copilot' }
    }
    return $Friendly
}

# --- Tenant SKUs: SkuId -> SkuPartNumber --------------------------------------
$skuMap = @{}
foreach ($sku in Get-MgSubscribedSku -All) {
    $skuMap[$sku.SkuId.ToString()] = $sku.SkuPartNumber
}

# --- Users in domain ----------------------------------------------------------
Write-Host "Retrieving users for @$Domain ..." -ForegroundColor Cyan
$props = 'id', 'displayName', 'userPrincipalName', 'mail', 'accountEnabled', 'userType', 'assignedLicenses', 'department', 'jobTitle', 'usageLocation'

$filter = "endswith(userPrincipalName,'@$Domain')"
if ($IncludeMail) { $filter += " or endswith(mail,'@$Domain')" }

$users = Get-MgUser -All -Filter $filter -Property $props -ConsistencyLevel eventual -CountVariable userCount |
    Select-Object $props
Write-Host ("  Found {0} users." -f @($users).Count) -ForegroundColor Green

# --- Build report -------------------------------------------------------------
$report = foreach ($u in $users) {
    $licenses = foreach ($lic in $u.AssignedLicenses) {
        $part = $skuMap[$lic.SkuId.ToString()]
        if (-not $part) { $part = $lic.SkuId.ToString() }
        $friendly = if ($friendlyNames.ContainsKey($part)) { $friendlyNames[$part] } else { $part }
        [pscustomobject]@{
            SkuPartNumber = $part
            Friendly      = $friendly
            Short         = Get-ShortLicenseName -Friendly $friendly -SkuPartNumber $part
        }
    }

    $base = [ordered]@{
        DisplayName       = $u.DisplayName
        UserPrincipalName = $u.UserPrincipalName
        Mail              = $u.Mail
        AccountEnabled    = $u.AccountEnabled
        UserType          = $u.UserType
        Department        = $u.Department
        JobTitle          = $u.JobTitle
        UsageLocation     = $u.UsageLocation
    }

    if ($OneRowPerLicense) {
        $rows = if ($licenses) { $licenses } else { [pscustomobject]@{ SkuPartNumber = ''; Friendly = 'Unlicensed'; Short = 'Unlicensed' } }
        foreach ($l in $rows) {
            $row = [ordered]@{}
            foreach ($k in $base.Keys) { $row[$k] = $base[$k] }
            $row.License       = $l.Friendly
            $row.LicenseShort  = $l.Short
            $row.SkuPartNumber = $l.SkuPartNumber
            [pscustomobject]$row
        }
    }
    else {
        $row = $base
        $row.IsLicensed     = [bool]$licenses
        $row.LicenseCount   = @($licenses).Count
        $row.Licenses       = if ($licenses) { ($licenses.Friendly | Sort-Object) -join '; ' } else { 'Unlicensed' }
        $row.LicensesShort  = if ($licenses) { ($licenses.Short | Sort-Object) -join '; ' } else { 'Unlicensed' }
        $row.SkuPartNumbers = ($licenses.SkuPartNumber | Sort-Object) -join '; '
        [pscustomobject]$row
    }
}

$report | Sort-Object DisplayName | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "Exported $(@($report).Count) rows to $OutputPath" -ForegroundColor Green

# --- Summary ------------------------------------------------------------------
Write-Host "`nLicense summary for @$Domain" -ForegroundColor Cyan
$allLic = foreach ($u in $users) {
    if (-not $u.AssignedLicenses) { 'Unlicensed' }
    foreach ($lic in $u.AssignedLicenses) {
        $part = $skuMap[$lic.SkuId.ToString()]
        if ($part -and $friendlyNames.ContainsKey($part)) { $friendlyNames[$part] } elseif ($part) { $part } else { $lic.SkuId }
    }
}
$allLic | Group-Object | Sort-Object Count -Descending | Format-Table @{n = 'License'; e = { $_.Name } }, Count -AutoSize

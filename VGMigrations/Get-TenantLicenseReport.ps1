#Requires -Version 7.0
<#
.SYNOPSIS
    Get-TenantLicenseReport.ps1 — Connects to a list of tenants and reports the licenses
    (subscribed SKUs) assigned in each one.

.DESCRIPTION
    Reads a list of tenants from an Excel workbook or CSV, connects to Microsoft Graph for
    each in turn, and pulls the tenant's subscribed SKUs — friendly product name,
    enabled/consumed/available unit counts. Writes one consolidated CSV covering every
    tenant plus a per-tenant summary in the log.

    Each tenant needs its own sign-in. If the signed-in account already has delegated
    admin access to every tenant in the list (e.g. a Partner/GDAP relationship), the
    browser will generally re-use SSO between tenants without asking for a password again.

.PARAMETER TenantsFile
    Path to the tenants workbook/CSV. Defaults to the standing Volaris tenant list.

.PARAMETER TenantIdColumn
    Excel column letter holding the tenant ID (.xlsx only). Default 'B'.

.PARAMETER SkipColumn
    Excel column letter holding the skip flag (.xlsx only) — rows where this column equals
    -SkipValue are excluded. Default 'L'.

.PARAMETER SkipValue
    Value in -SkipColumn that marks a tenant to be skipped (case-insensitive). Default 'Yes'.

.PARAMETER HeaderRow
    Row number the data starts after (.xlsx only) — row 1 is assumed to be headers. Default 1.

.PARAMETER Column
    CSV only: column name to read tenant identifiers from. Auto-detected if omitted
    (tries TenantId, Tenant, Domain, TenantDomain, Name).

.PARAMETER OutputPath
    Path for the consolidated CSV report. Defaults to a timestamped file next to TenantsFile.

.EXAMPLE
    .\Get-TenantLicenseReport.ps1
.EXAMPLE
    .\Get-TenantLicenseReport.ps1 -TenantsFile C:\tenants.csv
#>

param(
    [string]$TenantsFile = 'C:\Users\andyw\OneDrive - Volaris Group\GRP Data Security (Volaris Consolidated) - 3. Execution\M365 Migrations\Tenant IDs.xlsx',

    [string]$TenantIdColumn = 'B',
    [string]$SkipColumn     = 'L',
    [string]$SkipValue      = 'Yes',
    [int]$HeaderRow         = 1,

    [string]$Column,

    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

Write-Host "=== Get-TenantLicenseReport ===" -ForegroundColor Cyan
Write-Host "Tenants file: $TenantsFile" -ForegroundColor White

if (-not (Test-Path $TenantsFile)) {
    Write-Host "ERROR: Tenants file not found: $TenantsFile" -ForegroundColor Red
    exit 1
}

. (Join-Path $PSScriptRoot 'Ensure-GraphModules.ps1') -GraphModules @('Microsoft.Graph.Identity.DirectoryManagement')

if (-not $OutputPath) {
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputPath = Join-Path (Split-Path $TenantsFile -Parent) "TenantLicenseReport_$stamp.csv"
}

# Converts an Excel column letter (A, B, ..., Z, AA, AB, ...) to a 1-based index.
function ConvertFrom-ExcelColumnLetter([string]$Letter) {
    $Letter = $Letter.Trim().ToUpperInvariant()
    $index = 0
    foreach ($ch in $Letter.ToCharArray()) {
        $index = $index * 26 + ([int][char]$ch - [int][char]'A' + 1)
    }
    return $index
}

$ext = [System.IO.Path]::GetExtension($TenantsFile).ToLowerInvariant()
$tenants = [System.Collections.Generic.List[string]]::new()

if ($ext -in @('.xlsx', '.xlsm', '.xls')) {
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host "Installing ImportExcel module (CurrentUser)..." -ForegroundColor Yellow
        Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    }
    Import-Module ImportExcel -ErrorAction Stop

    $idIdx   = ConvertFrom-ExcelColumnLetter $TenantIdColumn
    $skipIdx = ConvertFrom-ExcelColumnLetter $SkipColumn
    Write-Host "Reading column $TenantIdColumn (tenant ID) — skipping rows where column $SkipColumn = '$SkipValue'." -ForegroundColor Gray

    $excelRows = @(Import-Excel -Path $TenantsFile -NoHeader -StartRow ($HeaderRow + 1) -ErrorAction Stop)
    $skippedCount = 0
    foreach ($row in $excelRows) {
        $idVal   = $row."P$idIdx"
        $skipVal = $row."P$skipIdx"
        if ([string]::IsNullOrWhiteSpace($idVal)) { continue }
        if ($skipVal -and $skipVal.ToString().Trim().Equals($SkipValue, [System.StringComparison]::OrdinalIgnoreCase)) {
            $skippedCount++
            continue
        }
        $tenants.Add($idVal.ToString().Trim())
    }
    if ($skippedCount -gt 0) { Write-Host "$skippedCount tenant(s) skipped (column $SkipColumn = '$SkipValue')." -ForegroundColor Yellow }
} else {
    # Detect CSV encoding from BOM — same helper as Import-FlyMappings.ps1 / search-domain.ps1.
    function Get-CsvEncoding([string]$Path) {
        try {
            $fs    = [System.IO.FileStream]::new($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            $bytes = New-Object byte[] 4
            $read  = $fs.Read($bytes, 0, 4)
            $fs.Close()
            if ($read -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { return 'UTF8' }
            if ($read -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) { return 'Unicode' }
            if ($read -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) { return 'BigEndianUnicode' }
        } catch {}
        return 'Default'
    }

    $encoding = Get-CsvEncoding $TenantsFile
    $rows = @(Import-Csv -Path $TenantsFile -Encoding $encoding)
    if ($rows.Count -eq 0) {
        Write-Host "Tenants file has no rows. Nothing to do." -ForegroundColor Yellow
        exit 0
    }

    $headers = $rows[0].PSObject.Properties.Name
    $tenantCol = $Column
    if (-not $tenantCol) {
        foreach ($candidate in @('TenantId','Tenant','Domain','TenantDomain','Name')) {
            if ($headers -contains $candidate) { $tenantCol = $candidate; break }
        }
    }
    if (-not $tenantCol) {
        Write-Host "ERROR: Could not find a tenant column. Tried TenantId/Tenant/Domain/TenantDomain/Name. Columns found: $($headers -join ', ')" -ForegroundColor Red
        Write-Host "Use -Column to specify one explicitly." -ForegroundColor Yellow
        exit 1
    }
    Write-Host "Using '$tenantCol' as the tenant identifier column." -ForegroundColor Gray

    foreach ($v in ($rows | ForEach-Object { $_.$tenantCol } | Where-Object { $_ })) { $tenants.Add($v.Trim()) }
}

$tenants = @($tenants | Where-Object { $_ } | Select-Object -Unique)
Write-Host "$($tenants.Count) unique tenant(s) to check." -ForegroundColor Cyan
Write-Host ""

# Same friendly-name lookup table search-domain.ps1 uses for its license sections.
$LicenseFriendlyMap = @{
    'SPE_E5'                              = 'Microsoft 365 E5'
    'SPE_E3'                              = 'Microsoft 365 E3'
    'SPE_F1'                              = 'Microsoft 365 F1'
    'SPB'                                 = 'Microsoft 365 Business Premium'
    'O365_BUSINESS_PREMIUM'               = 'Microsoft 365 Business Standard'
    'O365_BUSINESS_ESSENTIALS'            = 'Microsoft 365 Business Basic'
    'O365_BUSINESS'                       = 'Microsoft 365 Apps for Business'
    'OFFICESUBSCRIPTION'                  = 'Microsoft 365 Apps for Enterprise'
    'M365_F1'                             = 'Microsoft 365 F1'
    'M365_F3'                             = 'Microsoft 365 F3 (Frontline)'
    'ENTERPRISEPACK'                      = 'Office 365 E3'
    'ENTERPRISEPREMIUM'                   = 'Office 365 E5'
    'ENTERPRISEPACK_GOV'                  = 'Office 365 E3 (Government)'
    'ENTERPRISEPREMIUM_GOV'               = 'Office 365 E5 (Government)'
    'STANDARDPACK'                        = 'Office 365 E1'
    'DESKLESSPACK'                        = 'Office 365 F3'
    'EXCHANGESTANDARD'                    = 'Exchange Online (Plan 1)'
    'EXCHANGEENTERPRISE'                  = 'Exchange Online (Plan 2)'
    'EXCHANGEDESKLESS'                    = 'Exchange Online Kiosk'
    'EXCHANGEARCHIVE_ADDON'               = 'Exchange Online Archiving for Exchange Online'
    'EXCHANGEARCHIVE'                     = 'Exchange Online Archiving for Exchange Server'
    'EXCHANGEESSENTIALS'                  = 'Exchange Online Essentials'
    'SHAREPOINTSTANDARD'                  = 'SharePoint Online (Plan 1)'
    'SHAREPOINTENTERPRISE'                = 'SharePoint Online (Plan 2)'
    'MCOSTANDARD'                         = 'Skype for Business Online (Plan 2)'
    'MCOPSTN1'                            = 'Microsoft 365 Domestic Calling Plan'
    'MCOPSTN2'                            = 'Microsoft 365 Domestic and International Calling Plan'
    'MCOEV'                               = 'Microsoft Teams Phone Standard'
    'MCOEV_VIRTUALUSER'                   = 'Microsoft Teams Phone Resource Account'
    'MCOMEETADV'                          = 'Microsoft 365 Audio Conferencing'
    'TEAMS_EXPLORATORY'                   = 'Microsoft Teams Exploratory'
    'TEAMS_COMMERCIAL_TRIAL'              = 'Microsoft Teams Commercial Trial'
    'PROJECTPROFESSIONAL'                 = 'Project Plan 3'
    'PROJECTPREMIUM'                      = 'Project Plan 5'
    'PROJECT_P1'                          = 'Project Plan 1'
    'PROJECTESSENTIALS'                   = 'Project Online Essentials'
    'VISIOCLIENT'                         = 'Visio Plan 2'
    'VISIOONLINE_PLAN1'                   = 'Visio Plan 1'
    'POWER_BI_PRO'                        = 'Power BI Pro'
    'POWER_BI_STANDARD'                   = 'Power BI (Free)'
    'PBI_PREMIUM_PER_USER'                = 'Power BI Premium Per User'
    'POWERAPPS_PER_USER'                  = 'Power Apps per User Plan'
    'POWERAPPS_VIRAL'                     = 'Microsoft Power Apps Plan 2 Trial'
    'FLOW_FREE'                           = 'Microsoft Power Automate Free'
    'FLOW_PER_USER'                       = 'Power Automate per User Plan'
    'EMS'                                 = 'Enterprise Mobility + Security E3'
    'EMSPREMIUM'                          = 'Enterprise Mobility + Security E5'
    'AAD_PREMIUM'                         = 'Microsoft Entra ID P1'
    'AAD_PREMIUM_P2'                      = 'Microsoft Entra ID P2'
    'AAD_BASIC'                           = 'Microsoft Entra ID Basic'
    'INTUNE_A'                            = 'Microsoft Intune Plan 1'
    'INTUNE_A_VL'                         = 'Microsoft Intune Plan 1 (VL)'
    'IDENTITY_THREAT_PROTECTION'          = 'Microsoft 365 E5 Security'
    'INFORMATION_PROTECTION_COMPLIANCE'   = 'Microsoft 365 E5 Compliance'
    'WIN_DEF_ATP'                         = 'Microsoft Defender for Endpoint'
    'ATP_ENTERPRISE'                      = 'Microsoft Defender for Office 365 (Plan 1)'
    'THREAT_INTELLIGENCE'                 = 'Microsoft Defender for Office 365 (Plan 2)'
    'DYN365_ENTERPRISE_SALES'             = 'Dynamics 365 Sales Enterprise'
    'DYN365_ENTERPRISE_CUSTOMER_SERVICE'  = 'Dynamics 365 Customer Service Enterprise'
    'DYN365_BUSCENTRAL_ESSENTIAL'         = 'Dynamics 365 Business Central Essentials'
    'STREAM'                              = 'Microsoft Stream Trial'
    'WIN10_PRO_ENT_SUB'                   = 'Windows 10/11 Enterprise E3'
    'WIN10_VDA_E5'                        = 'Windows 10/11 Enterprise E5'
    'SHAREPOINTSTORAGE'                   = 'SharePoint Online Storage'
    'RIGHTSMANAGEMENT'                    = 'Azure Information Protection (Plan 1)'
}
function Get-LicenseFriendlyName([string]$SkuPartNumber) {
    if ($LicenseFriendlyMap.ContainsKey($SkuPartNumber)) { return $LicenseFriendlyMap[$SkuPartNumber] }
    return $SkuPartNumber
}

# Only report SKUs that actually grant a mailbox (Exchange Online) and/or OneDrive (bundled
# under the SharePoint service plan) — excludes add-on-only SKUs like EMS, Entra ID P1/P2,
# Power BI Pro, Defender, Visio, Project, Teams Phone, Audio Conferencing, etc.
function Test-SkuGrantsMailboxOrOneDrive($Sku) {
    foreach ($plan in $Sku.ServicePlans) {
        if ($plan.ServicePlanName -match '^EXCHANGE_S_' -or $plan.ServicePlanName -match '^SHAREPOINT') { return $true }
    }
    return $false
}

$allRows = [System.Collections.Generic.List[pscustomobject]]::new()
$ok = 0; $failed = 0

foreach ($tenant in $tenants) {
    Write-Host "--- $tenant ---" -ForegroundColor Cyan
    try {
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}

        Write-Host "  Connecting — sign in with the browser window that opens..." -ForegroundColor Gray
        Connect-MgGraph -TenantId $tenant -Scopes 'Organization.Read.All' -NoWelcome -ErrorAction Stop

        $org     = Get-MgOrganization -ErrorAction Stop | Select-Object -First 1
        $allSkus = @(Get-MgSubscribedSku -All -ErrorAction Stop)
        $skus    = @($allSkus | Where-Object { Test-SkuGrantsMailboxOrOneDrive $_ })
        $excludedCount = $allSkus.Count - $skus.Count

        if ($allSkus.Count -eq 0) {
            Write-Host "  No licenses found." -ForegroundColor Yellow
        } elseif ($skus.Count -eq 0) {
            Write-Host "  No mailbox/OneDrive-capable licenses found ($excludedCount other SKU(s) excluded)." -ForegroundColor Yellow
        } elseif ($excludedCount -gt 0) {
            Write-Host "  ($excludedCount SKU(s) without mailbox/OneDrive excluded)" -ForegroundColor DarkGray
        }

        foreach ($sku in $skus) {
            $friendly  = Get-LicenseFriendlyName $sku.SkuPartNumber
            $enabled   = $sku.PrepaidUnits.Enabled
            $consumed  = $sku.ConsumedUnits
            $available = $enabled - $consumed
            Write-Host ("  {0,-45} enabled={1,-6} consumed={2,-6} available={3}" -f $friendly, $enabled, $consumed, $available)

            $allRows.Add([pscustomobject]@{
                TenantInput   = $tenant
                TenantId      = $org.Id
                TenantName    = $org.DisplayName
                SkuPartNumber = $sku.SkuPartNumber
                FriendlyName  = $friendly
                Enabled       = $enabled
                Consumed      = $consumed
                Available     = $available
                Suspended     = $sku.PrepaidUnits.Suspended
                Warning       = $sku.PrepaidUnits.Warning
            })
        }

        Write-Host "  OK — $($skus.Count) SKU(s)." -ForegroundColor Green
        $ok++
    } catch {
        Write-Host "  FAILED: $($_.Exception.Message.Split([Environment]::NewLine)[0])" -ForegroundColor Red
        $failed++
    }
    Write-Host ""
}

try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}

if ($allRows.Count -gt 0) {
    $allRows | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8 -Force
    Write-Host "Report written: $OutputPath" -ForegroundColor Green
}

Write-Host ""
Write-Host "=== Done ===" -ForegroundColor Cyan
Write-Host "Tenants OK: $ok  |  Failed: $failed  |  License rows: $($allRows.Count)"

if ($failed -gt 0) { exit 1 }

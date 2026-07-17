#Requires -Version 7.0
<#
.SYNOPSIS
    Get-TenantLicenseReport.ps1 — Connects to a list of tenants and reports the licenses
    (subscribed SKUs) assigned in each one.

.DESCRIPTION
    Reads a list of tenants from an Excel workbook or CSV, connects to Microsoft Graph for
    each in turn, and pulls the tenant's subscribed SKUs — friendly product name,
    enabled/consumed/available unit counts. Writes one consolidated CSV covering every
    tenant plus a per-tenant summary in the log. Tenants with no mailbox/OneDrive-capable
    license still get a single row in the CSV (SkuPartNumber '(none)', all counts 0) so it's
    clear the tenant was checked rather than silently missing from the report.

    Two connection modes per tenant, chosen automatically:
      - App-only (no sign-in): used when the row has both an App ID and App Secret —
        authenticates as that app registration via client-credentials, no browser at all.
        Requires the app registration to already hold Organization.Read.All (application
        permission, admin-consented) in that tenant; falls back to interactive if it doesn't.
      - Interactive: used when App ID/Secret are missing, or app-only auth fails. The
        tenant's Domain is printed as a prominent banner right before the browser opens so
        you know which account to sign in with.

.PARAMETER TenantsFile
    Path to the tenants workbook/CSV. Defaults to the standing Volaris tenant list.

.PARAMETER DomainColumn
    Excel column letter holding the tenant's domain (.xlsx only, display only). Default 'A'.

.PARAMETER TenantIdColumn
    Excel column letter holding the tenant ID (.xlsx only). Default 'B'.

.PARAMETER AppIdColumn
    Excel column letter holding the app registration's Client/App ID (.xlsx only, optional —
    enables app-only auth for that row when paired with -AppSecretColumn). Default 'C'.

.PARAMETER AppSecretColumn
    Excel column letter holding the app registration's client secret (.xlsx only, optional).
    Default 'D'.

.PARAMETER SkipColumn
    Excel column letter holding the "cutover done" skip flag (.xlsx only) — rows where this
    column equals -SkipValue are excluded. Default 'L'.

.PARAMETER SkipValue
    Value in -SkipColumn / -LicensesOkColumn that marks a tenant to be skipped
    (case-insensitive). Default 'Yes'.

.PARAMETER LicensesOkColumn
    Excel column letter holding the "licenses already confirmed correct" flag (.xlsx only) —
    rows where this column equals -SkipValue are excluded, same as -SkipColumn. Default 'I'.

.PARAMETER HeaderRow
    Row number the data starts after (.xlsx only) — row 1 is assumed to be headers. Default 1.

.PARAMETER ExcludeDomains
    Domain substrings to exclude (case-insensitive) — matched against the Domain column.
    Default excludes 'ourvolaris' (the Volaris management tenant itself, not a customer).

.PARAMETER Column
    CSV only: column name to read tenant identifiers from. Auto-detected if omitted
    (tries TenantId, Tenant, Domain, TenantDomain, Name). CSV rows also get app-only auth
    automatically if the CSV has AppId/ClientId and AppSecret/ClientSecret columns.

.PARAMETER OutputPath
    Path for the consolidated CSV report. Defaults to a timestamped file next to TenantsFile.

.PARAMETER RefreshIntervalDays
    How often the local lookup copy is refreshed from -TenantsFile, in days. Default 3.5
    (twice a week). The script always reads from the copy, never the original, so the
    original stays free for you to keep open in Excel while a report runs.

.PARAMETER ForceRefreshLookup
    Refresh the lookup copy from -TenantsFile now, regardless of its age.

.EXAMPLE
    .\Get-TenantLicenseReport.ps1
.EXAMPLE
    .\Get-TenantLicenseReport.ps1 -TenantsFile C:\tenants.csv
#>

param(
    [string]$TenantsFile = 'C:\Users\andyw\OneDrive - Volaris Group\GRP Data Security (Volaris Consolidated) - 3. Execution\M365 Migrations\Tenant IDs.xlsx',

    [string]$DomainColumn    = 'A',
    [string]$TenantIdColumn  = 'B',
    [string]$AppIdColumn     = 'C',
    [string]$AppSecretColumn = 'D',
    [string]$SkipColumn      = 'L',
    [string]$LicensesOkColumn = 'I',
    [string]$SkipValue       = 'Yes',
    [int]$HeaderRow          = 1,

    [string[]]$ExcludeDomains = @('ourvolaris'),

    [string]$Column,

    [string]$OutputPath,

    [double]$RefreshIntervalDays = 3.5,
    [switch]$ForceRefreshLookup
)

$ErrorActionPreference = 'Stop'

# MSAL/Azure.Identity exceptions frequently put the actual detail on the second (or later)
# line, with a generic "X authentication failed:" on the first — taking only the first line
# (as this script used to) silently discards the real reason. Joins the first few non-blank
# lines into one readable summary instead.
function Get-CleanErrorMessage($ErrorRecord) {
    $lines = @($ErrorRecord.Exception.Message -split "`r?`n" | Where-Object { $_.Trim() })
    if ($lines.Count -eq 0) { return $ErrorRecord.Exception.GetType().Name }
    return ($lines | Select-Object -First 3) -join ' | '
}

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

# Always read from a local copy, never the original — keeps the original free for you to have
# open in Excel while a report runs, since Import-Excel can't open a file someone else has
# locked. The copy is refreshed from the original at most every -RefreshIntervalDays (default
# twice a week); run with -ForceRefreshLookup to sync it immediately regardless of age.
# Stored under %LOCALAPPDATA%\FlyMigration (same as shared-config.json) rather than
# $PSScriptRoot, which is read-only for non-admin users in a perMachine packaged install.
$origItem    = Get-Item -LiteralPath $TenantsFile
$lookupCacheDir = Join-Path $env:LOCALAPPDATA 'FlyMigration'
if (-not (Test-Path $lookupCacheDir)) { New-Item -ItemType Directory -Path $lookupCacheDir -Force | Out-Null }
$lookupFile = Join-Path $lookupCacheDir ("$($origItem.BaseName)-lookupcache$($origItem.Extension)")

$needsRefresh = $ForceRefreshLookup -or -not (Test-Path $lookupFile)
if (-not $needsRefresh) {
    $ageDays = ((Get-Date) - (Get-Item $lookupFile).LastWriteTime).TotalDays
    if ($ageDays -ge $RefreshIntervalDays) { $needsRefresh = $true }
}

if ($needsRefresh) {
    try {
        Copy-Item -LiteralPath $TenantsFile -Destination $lookupFile -Force -ErrorAction Stop
        Write-Host "Lookup copy refreshed from the original." -ForegroundColor Green
    } catch {
        if (Test-Path $lookupFile) {
            $ageDaysStale = [math]::Round(((Get-Date) - (Get-Item $lookupFile).LastWriteTime).TotalDays, 1)
            Write-Host "WARNING: Could not refresh lookup copy ($(Get-CleanErrorMessage $_)) — using existing copy ($ageDaysStale day(s) old)." -ForegroundColor Yellow
        } else {
            Write-Host "ERROR: Could not create the lookup copy and none exists yet: $(Get-CleanErrorMessage $_)" -ForegroundColor Red
            exit 1
        }
    }
} else {
    $ageDaysDisplay = [math]::Round(((Get-Date) - (Get-Item $lookupFile).LastWriteTime).TotalDays, 1)
    Write-Host "Using existing lookup copy ($ageDaysDisplay day(s) old — refreshes every $RefreshIntervalDays days)." -ForegroundColor Gray
}
Write-Host "Lookup copy: $lookupFile" -ForegroundColor White

# Converts an Excel column letter (A, B, ..., Z, AA, AB, ...) to a 1-based index.
function ConvertFrom-ExcelColumnLetter([string]$Letter) {
    $Letter = $Letter.Trim().ToUpperInvariant()
    $index = 0
    foreach ($ch in $Letter.ToCharArray()) {
        $index = $index * 26 + ([int][char]$ch - [int][char]'A' + 1)
    }
    return $index
}

$ext = [System.IO.Path]::GetExtension($lookupFile).ToLowerInvariant()
# Each record: Domain, TenantId, AppId, AppSecret (AppId/AppSecret may be blank)
$tenantRecords = [System.Collections.Generic.List[pscustomobject]]::new()

if ($ext -in @('.xlsx', '.xlsm', '.xls')) {
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host "Installing ImportExcel module (CurrentUser)..." -ForegroundColor Yellow
        Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    }
    Import-Module ImportExcel -ErrorAction Stop

    $domainIdx  = ConvertFrom-ExcelColumnLetter $DomainColumn
    $idIdx      = ConvertFrom-ExcelColumnLetter $TenantIdColumn
    $appIdIdx   = ConvertFrom-ExcelColumnLetter $AppIdColumn
    $appSecIdx  = ConvertFrom-ExcelColumnLetter $AppSecretColumn
    $skipIdx    = ConvertFrom-ExcelColumnLetter $SkipColumn
    $licOkIdx   = ConvertFrom-ExcelColumnLetter $LicensesOkColumn
    Write-Host "Reading column $TenantIdColumn (tenant ID), $DomainColumn (domain) — skipping rows where column $SkipColumn or $LicensesOkColumn = '$SkipValue'." -ForegroundColor Gray

    $excelRows = @(Import-Excel -Path $lookupFile -NoHeader -StartRow ($HeaderRow + 1) -ErrorAction Stop)
    $skippedCount = 0
    foreach ($row in $excelRows) {
        $idVal    = $row."P$idIdx"
        $skipVal  = $row."P$skipIdx"
        $licOkVal = $row."P$licOkIdx"
        if ([string]::IsNullOrWhiteSpace($idVal)) { continue }
        if ($skipVal -and $skipVal.ToString().Trim().Equals($SkipValue, [System.StringComparison]::OrdinalIgnoreCase)) {
            $skippedCount++
            continue
        }
        if ($licOkVal -and $licOkVal.ToString().Trim().Equals($SkipValue, [System.StringComparison]::OrdinalIgnoreCase)) {
            $skippedCount++
            continue
        }
        $tenantRecords.Add([pscustomobject]@{
            Domain    = $row."P$domainIdx"
            TenantId  = $idVal.ToString().Trim()
            AppId     = $row."P$appIdIdx"
            AppSecret = $row."P$appSecIdx"
        })
    }
    if ($skippedCount -gt 0) { Write-Host "$skippedCount tenant(s) skipped (column $SkipColumn or $LicensesOkColumn = '$SkipValue')." -ForegroundColor Yellow }
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

    $encoding = Get-CsvEncoding $lookupFile
    $rows = @(Import-Csv -Path $lookupFile -Encoding $encoding)
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

    $domainCol = @('Domain','TenantDomain') | Where-Object { $headers -contains $_ } | Select-Object -First 1
    $appIdCol  = @('AppId','ClientId') | Where-Object { $headers -contains $_ } | Select-Object -First 1
    $appSecCol = @('AppSecret','ClientSecret','Secret') | Where-Object { $headers -contains $_ } | Select-Object -First 1

    foreach ($row in $rows) {
        if (-not $row.$tenantCol) { continue }
        $tenantRecords.Add([pscustomobject]@{
            Domain    = if ($domainCol) { $row.$domainCol } else { $null }
            TenantId  = $row.$tenantCol.Trim()
            AppId     = if ($appIdCol)  { $row.$appIdCol }  else { $null }
            AppSecret = if ($appSecCol) { $row.$appSecCol } else { $null }
        })
    }
}

# Dedup by TenantId, keeping the first occurrence of each.
$seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$tenantRecords = @($tenantRecords | Where-Object { $_.TenantId -and $seen.Add($_.TenantId) })

# Exclude the Volaris management tenant itself (and any other -ExcludeDomains match) — it's
# the source/management tenant, not a customer to be checked.
if ($ExcludeDomains -and $ExcludeDomains.Count -gt 0) {
    $beforeCount = $tenantRecords.Count
    $tenantRecords = @($tenantRecords | Where-Object {
        $rec = $_
        -not ($ExcludeDomains | Where-Object { $rec.Domain -and $rec.Domain -match [regex]::Escape($_) })
    })
    $excludedByDomain = $beforeCount - $tenantRecords.Count
    if ($excludedByDomain -gt 0) {
        Write-Host "$excludedByDomain tenant(s) excluded (domain matches: $($ExcludeDomains -join ', '))." -ForegroundColor Yellow
    }
}

Write-Host "$($tenantRecords.Count) unique tenant(s) to check." -ForegroundColor Cyan
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
#
# Trap: many trial/add-on-only SKUs (Power Automate Free, Teams Exploratory, Defender for
# Endpoint, Windows Store for Business, etc.) bundle EXCHANGE_S_FOUNDATION purely as internal
# plumbing — it grants no real mailbox. SHAREPOINTWAC is Office-for-the-web viewing rights, not
# storage. Both look like real Exchange/SharePoint matches to a plain name-prefix check, so they
# have to be excluded explicitly rather than relying on the prefix alone.
function Test-SkuGrantsMailboxOrOneDrive($Sku) {
    foreach ($plan in $Sku.ServicePlans) {
        $name = $plan.ServicePlanName
        if ($name -match '_FOUNDATION$') { continue }
        if ($name -eq 'SHAREPOINTWAC' -or $name -eq 'SHAREPOINTWAC_EDU') { continue }
        if ($name -match '^EXCHANGE_S_' -or $name -match '^SHAREPOINT') { return $true }
    }
    return $false
}

$allRows = [System.Collections.Generic.List[pscustomobject]]::new()
$ok = 0; $failed = 0

foreach ($rec in $tenantRecords) {
    $label = if ($rec.Domain) { $rec.Domain } else { $rec.TenantId }
    Write-Host ""
    Write-Host "════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  TENANT: $label" -ForegroundColor Cyan
    if ($rec.Domain) { Write-Host "  ($($rec.TenantId))" -ForegroundColor DarkGray }
    Write-Host "════════════════════════════════════════════════════" -ForegroundColor Cyan

    try {
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}

        $hasAppCreds = -not [string]::IsNullOrWhiteSpace($rec.AppId) -and -not [string]::IsNullOrWhiteSpace($rec.AppSecret)
        $connected = $false

        if ($hasAppCreds) {
            Write-Host "  Connecting with the stored app registration — no sign-in needed..." -ForegroundColor Gray
            try {
                $secureSecret = ConvertTo-SecureString $rec.AppSecret -AsPlainText -Force
                $cred = [PSCredential]::new($rec.AppId, $secureSecret)
                Connect-MgGraph -TenantId $rec.TenantId -ClientSecretCredential $cred -NoWelcome -ErrorAction Stop | Out-Null
                # Client-credentials auth can issue a token even when the app was never
                # actually consented in this tenant (no service principal) — that only
                # surfaces on the first real Graph call, as Authorization_IdentityNotFound.
                # Verify here so a "successful" connect that doesn't actually work still
                # falls back to interactive instead of being reported as a hard failure.
                $org = Get-MgOrganization -ErrorAction Stop | Select-Object -First 1
                $connected = $true
            } catch {
                Write-Host "  App-only auth failed: $(Get-CleanErrorMessage $_)" -ForegroundColor Yellow
                Write-Host "  Falling back to interactive sign-in." -ForegroundColor Yellow
                try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
            } finally {
                $secureSecret = $null; $cred = $null
            }
        }

        if (-not $connected) {
            Write-Host ""
            Write-Host "  >>> SIGN IN TO: $label <<<" -ForegroundColor Yellow
            Write-Host "  A browser window will open — use the admin account for this tenant." -ForegroundColor Gray

            # MSAL/WAM broker diagnostics get written directly to .NET's Console.Out/Error,
            # bypassing every PowerShell stream — a plain .Exception.Message can end up
            # completely blank even though MSAL actually explained the failure (same issue
            # search-domain.ps1 documented and worked around for its own Graph/EXO connects).
            # Capture Console.Out/Error during the attempt and fold it into the thrown error
            # so a failure here actually shows a reason instead of "authentication failed: ".
            $consoleSwallow = New-Object System.IO.StringWriter
            $origConsoleOut = [Console]::Out
            $origConsoleErr = [Console]::Error
            [Console]::SetOut($consoleSwallow)
            [Console]::SetError($consoleSwallow)
            try {
                Connect-MgGraph -TenantId $rec.TenantId -Scopes 'Organization.Read.All' -NoWelcome -ErrorAction Stop | Out-Null
                $org = Get-MgOrganization -ErrorAction Stop | Select-Object -First 1
            } catch {
                [Console]::SetOut($origConsoleOut); [Console]::SetError($origConsoleErr)
                $swallowed = $consoleSwallow.ToString().Trim()
                if ($swallowed) {
                    throw "$(Get-CleanErrorMessage $_) — broker output: $($swallowed -replace '[\r\n]+', ' ')"
                }
                throw
            } finally {
                if ([Console]::Out -ne $origConsoleOut) { [Console]::SetOut($origConsoleOut) }
                if ([Console]::Error -ne $origConsoleErr) { [Console]::SetError($origConsoleErr) }
                $consoleSwallow.Dispose()
            }
        }

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

        if ($skus.Count -eq 0) {
            # Still record the tenant in the report — a 0 row makes it obvious this tenant
            # was checked and genuinely has no mailbox/OneDrive-capable license, rather than
            # silently missing from the CSV.
            $allRows.Add([pscustomobject]@{
                Domain        = $rec.Domain
                TenantId      = $org.Id
                TenantName    = $org.DisplayName
                SkuPartNumber = '(none)'
                FriendlyName  = '(none)'
                Enabled       = 0
                Consumed      = 0
                Available     = 0
                Suspended     = 0
                Warning       = 0
            })
        }

        foreach ($sku in $skus) {
            $friendly  = Get-LicenseFriendlyName $sku.SkuPartNumber
            $enabled   = $sku.PrepaidUnits.Enabled
            $consumed  = $sku.ConsumedUnits
            $available = $enabled - $consumed
            Write-Host ("  {0,-45} enabled={1,-6} consumed={2,-6} available={3}" -f $friendly, $enabled, $consumed, $available)

            $allRows.Add([pscustomobject]@{
                Domain        = $rec.Domain
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
        $errMsg = Get-CleanErrorMessage $_
        Write-Host "  FAILED: $errMsg" -ForegroundColor Red
        # Still record the tenant — without this, a connect/query failure makes the tenant
        # vanish from the CSV entirely instead of showing up as an obvious error row.
        $allRows.Add([pscustomobject]@{
            Domain        = $rec.Domain
            TenantId      = $rec.TenantId
            TenantName    = '(FAILED)'
            SkuPartNumber = '(ERROR)'
            FriendlyName  = $errMsg
            Enabled       = 0
            Consumed      = 0
            Available     = 0
            Suspended     = 0
            Warning       = 0
        })
        $failed++
    }
}

try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}

if ($allRows.Count -gt 0) {
    $allRows | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8 -Force
    Write-Host ""
    Write-Host "Report written: $OutputPath" -ForegroundColor Green
    # Signal the Electron app to open this file natively (same convention as
    # New-AzureAppRegistration.ps1) — pops the CSV open on screen when the run finishes.
    Write-Output "##OPEN_FILE:$OutputPath##"
}

Write-Host ""
Write-Host "=== Done ===" -ForegroundColor Cyan
Write-Host "Tenants OK: $ok  |  Failed: $failed  |  License rows: $($allRows.Count)"

if ($failed -gt 0) { exit 1 }

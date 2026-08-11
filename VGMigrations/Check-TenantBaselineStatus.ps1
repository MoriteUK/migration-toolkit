#Requires -Version 7.0
<#
.SYNOPSIS
    Check-TenantBaselineStatus.ps1 — Connects to a list of tenants and reports whether
    Invoke-TenantBaseline.ps1's configuration has already been applied in each one.

.DESCRIPTION
    Reads a list of tenants from an Excel workbook or CSV (same file/format as
    Get-TenantLicenseReport.ps1) and, for each one, checks — read-only, nothing is changed —
    whether every step of Invoke-TenantBaseline.ps1 is already in place: authorization policy,
    admin consent workflow, Intune MDM scope, dynamic device groups, blocked personal
    enrollment, BitLocker/Windows compliance policies, iOS/Android app protection, the
    Conditional Access policy set, and LAPS. Writes one consolidated CSV, one row per tenant,
    one column per check, plus an overall summary column.

    Connects using ONLY the existing per-tenant app registration already used by
    Get-TenantLicenseReport.ps1 (stored appcreds_<tenantid>.json, or the workbook's AppId/
    AppSecret columns) — no new app registration is created, and there is no interactive
    fallback. If a tenant has no stored app registration, or authentication fails outright, it's
    skipped with a clear reason. If authentication succeeds but an individual check comes back
    403 (the app registration was consented for license reporting, not necessarily for these
    broader read scopes), that check is reported as "Unknown (insufficient permission)" rather
    than failing the whole tenant — see NOTES for the permissions this needs.

    Exchange Online settings (Invoke-TenantBaseline.ps1 steps 12-13: EOP/ATP preset policies,
    transport rules, audit logging) are NOT checked here — app-only Graph credentials don't
    cover Exchange Online, and setting up certificate-based EXO app-only access for every
    customer tenant is a separate piece of work. Those two steps still need a manual/interactive
    check.

.PARAMETER TenantsFile
    Path to the tenants workbook/CSV. Defaults to the standing Volaris tenant list (same
    default as Get-TenantLicenseReport.ps1).

.PARAMETER DomainColumn
    Excel column letter holding the tenant's domain (.xlsx only, display only). Default 'A'.

.PARAMETER TenantIdColumn
    Excel column letter holding the tenant ID (.xlsx only). Default 'B'.

.PARAMETER AppIdColumn
    Excel column letter holding the app registration's Client/App ID (.xlsx only). Default 'C'.

.PARAMETER AppSecretColumn
    Excel column letter holding the app registration's client secret (.xlsx only). Default 'D'.

.PARAMETER SkipColumn
    Excel column letter holding the "cutover done" skip flag (.xlsx only) — rows where this
    column equals -SkipValue are excluded. Default 'N'.

.PARAMETER SkipValue
    Value in -SkipColumn that marks a tenant to be skipped (case-insensitive). Default 'Yes'.

.PARAMETER HeaderRow
    Row number the data starts after (.xlsx only) — row 1 is assumed to be headers. Default 1.

.PARAMETER ExcludeDomains
    Domain substrings to exclude (case-insensitive). Default excludes 'ourvolaris'/'Volaris'
    (the management tenant itself, not a customer).

.PARAMETER Column
    CSV only: column name to read tenant identifiers from. Auto-detected if omitted.

.PARAMETER OutputPath
    Path for the consolidated CSV report. Defaults to a timestamped file next to TenantsFile.

.PARAMETER RefreshIntervalDays
    How often the local lookup copy is refreshed from -TenantsFile, in days. Default 3.5.

.PARAMETER ForceRefreshLookup
    Refresh the lookup copy from -TenantsFile now, regardless of its age.

.PARAMETER ProcessAll
    Process all tenants, ignoring SkipColumn. Still respects ExcludeDomains.

.NOTES
    The existing per-tenant app registration only needs Organization.Read.All for
    Get-TenantLicenseReport.ps1. To get real answers (not "Unknown (insufficient permission)")
    from every check here, it also needs these application permissions, admin-consented, in
    each tenant: Policy.Read.All, Group.Read.All, DeviceManagementConfiguration.Read.All,
    DeviceManagementApps.Read.All, DeviceManagementServiceConfig.Read.All.

.EXAMPLE
    .\Check-TenantBaselineStatus.ps1
    .\Check-TenantBaselineStatus.ps1 -TenantsFile C:\tenants.csv
    .\Check-TenantBaselineStatus.ps1 -ProcessAll
#>

param(
    [string]$TenantsFile = 'C:\Users\andyw\OneDrive - Volaris Group\GRP Data Security (Volaris Consolidated) - 3. Execution\M365 Migrations\Tenant IDs.xlsx',

    [string]$DomainColumn    = 'A',
    [string]$TenantIdColumn  = 'B',
    [string]$AppIdColumn     = 'C',
    [string]$AppSecretColumn = 'D',
    [string]$SkipColumn      = 'N',
    [string]$SkipValue       = 'Yes',
    [int]$HeaderRow          = 1,

    [string[]]$ExcludeDomains = @('ourvolaris', 'Volaris'),

    [string]$Column,

    [string]$OutputPath,

    [double]$RefreshIntervalDays = 3.5,
    [switch]$ForceRefreshLookup,
    [switch]$ProcessAll
)

$ErrorActionPreference = 'Stop'

function Get-CleanErrorMessage($ErrorRecord) {
    $lines = @($ErrorRecord.Exception.Message -split "`r?`n" | Where-Object { $_.Trim() })
    if ($lines.Count -eq 0) { return $ErrorRecord.Exception.GetType().Name }
    return ($lines | Select-Object -First 3) -join ' | '
}

# A check that 403s means the app registration hasn't been consented for that permission in this
# tenant — distinct from the check genuinely finding nothing configured. Surfacing which is which
# saves a lot of confusion versus a flat "Error".
function Test-InsufficientPermission($ErrorRecord) {
    return $ErrorRecord.Exception.Message -match 'Authorization_RequestDenied|Insufficient privileges|403'
}

Write-Host "=== Check-TenantBaselineStatus ===" -ForegroundColor Cyan
Write-Host "Tenants file: $TenantsFile" -ForegroundColor White

if (-not (Test-Path $TenantsFile)) {
    Write-Host "ERROR: Tenants file not found: $TenantsFile" -ForegroundColor Red
    exit 1
}

. (Join-Path $PSScriptRoot 'Ensure-GraphModules.ps1')

if (-not $OutputPath) {
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputPath = Join-Path (Split-Path $TenantsFile -Parent) "TenantBaselineStatus_$stamp.csv"
}

# Same local-lookup-copy convention as Get-TenantLicenseReport.ps1 — never lock the original.
$origItem       = Get-Item -LiteralPath $TenantsFile
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

function ConvertFrom-ExcelColumnLetter([string]$Letter) {
    $Letter = $Letter.Trim().ToUpperInvariant()
    $index = 0
    foreach ($ch in $Letter.ToCharArray()) {
        $index = $index * 26 + ([int][char]$ch - [int][char]'A' + 1)
    }
    return $index
}

$ext = [System.IO.Path]::GetExtension($lookupFile).ToLowerInvariant()
$tenantRecords = [System.Collections.Generic.List[pscustomobject]]::new()

if ($ext -in @('.xlsx', '.xlsm', '.xls')) {
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host "Installing ImportExcel module (CurrentUser)..." -ForegroundColor Yellow
        Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    }
    Import-Module ImportExcel -ErrorAction Stop

    $domainIdx = ConvertFrom-ExcelColumnLetter $DomainColumn
    $idIdx     = ConvertFrom-ExcelColumnLetter $TenantIdColumn
    $appIdIdx  = ConvertFrom-ExcelColumnLetter $AppIdColumn
    $appSecIdx = ConvertFrom-ExcelColumnLetter $AppSecretColumn
    $skipIdx   = ConvertFrom-ExcelColumnLetter $SkipColumn

    $excelRows = @(Import-Excel -Path $lookupFile -NoHeader -StartRow ($HeaderRow + 1) -ErrorAction Stop)
    $skippedCount = 0
    foreach ($row in $excelRows) {
        $idVal   = $row."P$idIdx"
        $skipVal = $row."P$skipIdx"
        if ([string]::IsNullOrWhiteSpace($idVal)) { continue }

        if (-not $ProcessAll -and $skipVal -and $skipVal.ToString().Trim().Equals($SkipValue, [System.StringComparison]::OrdinalIgnoreCase)) {
            $skippedCount++
            continue
        }

        $tidTrimmed = $idVal.ToString().Trim()
        $finalAppId = $null
        $finalAppSecret = $null

        # Same stored-creds-first lookup as Get-TenantLicenseReport.ps1 — reuse whatever app
        # registration that script already found for this tenant.
        $credsFile = Join-Path $PSScriptRoot "appcreds_$tidTrimmed.json"
        if (Test-Path $credsFile) {
            try {
                $creds = Get-Content $credsFile -Raw | ConvertFrom-Json
                $finalAppId = $creds.AppId
                $finalAppSecret = $creds.ClientSecret
            } catch {}
        }
        if ([string]::IsNullOrWhiteSpace($finalAppId)) {
            $excelAppId = $row."P$appIdIdx"
            $excelAppSecret = $row."P$appSecIdx"
            if (-not [string]::IsNullOrWhiteSpace($excelAppId) -and -not [string]::IsNullOrWhiteSpace($excelAppSecret)) {
                $finalAppId = $excelAppId
                $finalAppSecret = $excelAppSecret
            }
        }

        $tenantRecords.Add([pscustomobject]@{
            Domain    = $row."P$domainIdx"
            TenantId  = $tidTrimmed
            AppId     = $finalAppId
            AppSecret = $finalAppSecret
        })
    }
    if ($skippedCount -gt 0) { Write-Host "$skippedCount tenant(s) skipped (column $SkipColumn = '$SkipValue')." -ForegroundColor Yellow }
} else {
    $rows = @(Import-Csv -Path $lookupFile -Encoding UTF8)
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
        Write-Host "ERROR: Could not find a tenant column. Use -Column to specify one explicitly. Columns found: $($headers -join ', ')" -ForegroundColor Red
        exit 1
    }
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

$seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$tenantRecords = @($tenantRecords | Where-Object { $_.TenantId -and $seen.Add($_.TenantId) })

if ($ExcludeDomains -and $ExcludeDomains.Count -gt 0) {
    $tenantRecords = @($tenantRecords | Where-Object {
        $rec = $_
        -not ($ExcludeDomains | Where-Object { $rec.Domain -and $rec.Domain -match [regex]::Escape($_) })
    })
}

Write-Host "$($tenantRecords.Count) unique tenant(s) to check." -ForegroundColor Cyan
Write-Host ""

# ── The expected Conditional Access policy set (kept in sync with Invoke-TenantBaseline.ps1) ──
$ExpectedCAPolicies = @(
    'Block Access Outside Approved Countries'
    'Block Legacy Authentication'
    'Block Device Code Flow'
    'Require MFA for Admin Portals - 8hr'
    'Require MFA for Admin Roles - 8hr'
    'MFA for All Users - Browser Only - 8hr'
    'Require MFA for All Users'
    'Require MFA for Guest Users - 8hr'
)

# The 8 dynamic device groups Invoke-TenantBaseline.ps1 creates in Step 5.
$ExpectedDynamicGroups = @(
    'dyn-byod-android-devices'
    'dyn-corp-android-devices'
    'dyn-autopilot-devices'
    'dyn-byod-iOSiPad-devices'
    'dyn-corp-iOSiPad-devices'
    'dyn-corp-win10-devices'
    'dyn-corp-win11-devices'
    'dyn-byod-macOS-devices'
)

# One check = one Graph GET + a boolean test. Returns 'Configured' / 'NotConfigured' /
# 'Unknown (insufficient permission)' / 'Unknown (error)' — never throws past this point, so one
# failing check can't take down the rest of the tenant's row.
function Invoke-BaselineCheck {
    param([scriptblock]$Test)
    try {
        $result = & $Test
        if ($result) { return 'Configured' } else { return 'NotConfigured' }
    } catch {
        if (Test-InsufficientPermission $_) { return 'Unknown (insufficient permission)' }
        return "Unknown (error: $(Get-CleanErrorMessage $_))"
    }
}

$allRows = [System.Collections.Generic.List[pscustomobject]]::new()
$ok = 0; $failed = 0; $skippedNoCreds = 0

foreach ($rec in $tenantRecords) {
    $label = if ($rec.Domain) { $rec.Domain } else { $rec.TenantId }
    Write-Host ""
    Write-Host "════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  TENANT: $label" -ForegroundColor Cyan
    Write-Host "════════════════════════════════════════════════════" -ForegroundColor Cyan

    $hasAppCreds = -not [string]::IsNullOrWhiteSpace($rec.AppId) -and -not [string]::IsNullOrWhiteSpace($rec.AppSecret)
    if (-not $hasAppCreds) {
        Write-Host "  No app registration on file for this tenant — skipped (no interactive fallback)." -ForegroundColor Yellow
        $allRows.Add([pscustomobject]@{
            Domain = $rec.Domain; TenantId = $rec.TenantId; TenantName = '(not checked)'
            AuthorizationPolicy = 'Skipped'; AdminConsentWorkflow = 'Skipped'; IntuneMDMScope = 'Skipped'
            DynamicDeviceGroups = 'Skipped'; BlockPersonalEnrollment = 'Skipped'; BitLockerPolicy = 'Skipped'
            WindowsCompliancePolicy = 'Skipped'; iOSAppProtection = 'Skipped'; AndroidAppProtection = 'Skipped'
            ConditionalAccessPolicies = 'Skipped'; LAPSPolicy = 'Skipped'
            OverallStatus = 'Skipped — no app registration on file'
        })
        $skippedNoCreds++
        continue
    }

    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}

    try {
        Write-Host "  Connecting with the stored app registration — no sign-in needed..." -ForegroundColor Gray
        $secureSecret = ConvertTo-SecureString $rec.AppSecret -AsPlainText -Force
        $cred = [PSCredential]::new($rec.AppId, $secureSecret)
        Connect-MgGraph -TenantId $rec.TenantId -ClientSecretCredential $cred -NoWelcome -ErrorAction Stop | Out-Null
        $org = Get-MgOrganization -ErrorAction Stop | Select-Object -First 1
    } catch {
        $errMsg = Get-CleanErrorMessage $_
        Write-Host "  FAILED to authenticate: $errMsg" -ForegroundColor Red
        $allRows.Add([pscustomobject]@{
            Domain = $rec.Domain; TenantId = $rec.TenantId; TenantName = '(FAILED)'
            AuthorizationPolicy = 'Skipped'; AdminConsentWorkflow = 'Skipped'; IntuneMDMScope = 'Skipped'
            DynamicDeviceGroups = 'Skipped'; BlockPersonalEnrollment = 'Skipped'; BitLockerPolicy = 'Skipped'
            WindowsCompliancePolicy = 'Skipped'; iOSAppProtection = 'Skipped'; AndroidAppProtection = 'Skipped'
            ConditionalAccessPolicies = 'Skipped'; LAPSPolicy = 'Skipped'
            OverallStatus = "Auth failed: $errMsg"
        })
        $failed++
        continue
    } finally {
        $secureSecret = $null; $cred = $null
    }

    # ── Authorization Policy ────────────────────────────────────────────────
    $authPolicyStatus = Invoke-BaselineCheck {
        $p = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/policies/authorizationPolicy'
        $p.defaultUserRolePermissions.allowedToCreateApps -eq $false -and
        $p.defaultUserRolePermissions.allowedToCreateSecurityGroups -eq $false -and
        $p.defaultUserRolePermissions.allowedToCreateTenants -eq $false -and
        $p.allowUserConsentForRiskyApps -eq $false -and
        $p.allowEmailVerifiedUsersToJoinOrg -eq $false -and
        $p.allowInvitesFrom -eq 'adminsAndGuestInviters' -and
        $p.allowUserConsentForApps -eq $false -and
        $p.guestUserRoleId -eq '2af84b1e-32c8-42b7-82bc-daa82404023b'
    }

    # ── Admin Consent Workflow ──────────────────────────────────────────────
    $adminConsentStatus = Invoke-BaselineCheck {
        $group = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq 'Admin Consent Reviewers'").value | Select-Object -First 1
        if (-not $group) { return $false }
        $policy = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/policies/adminConsentRequestPolicy'
        $policy.isEnabled -eq $true -and ($policy.reviewers | Where-Object { $_.query -eq "/groups/$($group.id)/transitiveMembers" })
    }

    # ── Intune MDM User Scope ───────────────────────────────────────────────
    $intuneScopeStatus = Invoke-BaselineCheck {
        $resp = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/beta/policies/mobileDeviceManagementPolicies'
        $policy = $resp.value | Where-Object { $_.displayName -eq 'Microsoft Intune' }
        $policy -and $policy.appliesTo -eq 'all'
    }

    # ── Dynamic Device Groups ───────────────────────────────────────────────
    $dynGroupsFound = 0
    try {
        foreach ($name in $ExpectedDynamicGroups) {
            $existing = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$name'").value
            if ($existing) { $dynGroupsFound++ }
        }
        $dynGroupsStatus = if ($dynGroupsFound -eq $ExpectedDynamicGroups.Count) { 'Configured' }
                           elseif ($dynGroupsFound -eq 0) { 'NotConfigured' }
                           else { "Partial ($dynGroupsFound/$($ExpectedDynamicGroups.Count))" }
    } catch {
        $dynGroupsStatus = if (Test-InsufficientPermission $_) { 'Unknown (insufficient permission)' } else { "Unknown (error: $(Get-CleanErrorMessage $_))" }
    }

    # ── Block Personal Device Enrollment ────────────────────────────────────
    $blockPersonalStatus = Invoke-BaselineCheck {
        $configs = (Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/deviceManagement/deviceEnrollmentConfigurations').value
        $default = $configs | Where-Object { $_.id -like '*_DefaultPlatformRestrictions' } | Select-Object -First 1
        if (-not $default) { return $false }
        $full = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/deviceManagement/deviceEnrollmentConfigurations/$($default.id)"
        $full.androidRestriction.personalDeviceEnrollmentBlocked -eq $true -and
        $full.androidForWorkRestriction.personalDeviceEnrollmentBlocked -eq $true -and
        $full.iosRestriction.personalDeviceEnrollmentBlocked -eq $true -and
        $full.macOSRestriction.personalDeviceEnrollmentBlocked -eq $true -and
        $full.windowsRestriction.personalDeviceEnrollmentBlocked -eq $true
    }

    # ── BitLocker Compliance Policy ─────────────────────────────────────────
    $bitLockerStatus = Invoke-BaselineCheck {
        $configs = (Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/deviceManagement/deviceConfigurations').value
        [bool]($configs | Where-Object { $_.displayName -eq 'Enforce BitLocker Encryption' })
    }

    # ── Windows Compliance Policy ───────────────────────────────────────────
    $winComplianceStatus = Invoke-BaselineCheck {
        $policies = (Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies').value
        [bool]($policies | Where-Object { $_.displayName -eq 'Baseline Compliance Policy - Windows 10/11' })
    }

    # ── iOS App Protection Policy ───────────────────────────────────────────
    $iosProtectionStatus = Invoke-BaselineCheck {
        $policies = (Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/deviceAppManagement/iosManagedAppProtections').value
        [bool]($policies | Where-Object { $_.displayName -eq 'iOS App Protection Policy' })
    }

    # ── Android App Protection Policy ───────────────────────────────────────
    $androidProtectionStatus = Invoke-BaselineCheck {
        $policies = (Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/deviceAppManagement/androidManagedAppProtections').value
        [bool]($policies | Where-Object { $_.displayName -eq 'Android App Protection Policy' })
    }

    # ── Conditional Access Policies ─────────────────────────────────────────
    $caFound = 0; $caWrongState = 0
    try {
        $allCAPolicies = (Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies').value
        foreach ($name in $ExpectedCAPolicies) {
            $policy = $allCAPolicies | Where-Object { $_.displayName -eq $name }
            if (-not $policy) { continue }
            $caFound++
            if ($policy.state -ne 'disabled') { $caWrongState++ }
        }
        $caStatus = if ($caFound -eq $ExpectedCAPolicies.Count -and $caWrongState -eq 0) { 'Configured' }
                    elseif ($caFound -eq 0) { 'NotConfigured' }
                    elseif ($caWrongState -gt 0) { "Partial ($caFound/$($ExpectedCAPolicies.Count) present, $caWrongState not disabled)" }
                    else { "Partial ($caFound/$($ExpectedCAPolicies.Count))" }
    } catch {
        $caStatus = if (Test-InsufficientPermission $_) { 'Unknown (insufficient permission)' } else { "Unknown (error: $(Get-CleanErrorMessage $_))" }
    }

    # ── LAPS (policy + remediation script) ──────────────────────────────────
    $lapsStatus = Invoke-BaselineCheck {
        $policyUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?`$filter=templateReference/TemplateDisplayName%20eq%20%27Local%20admin%20password%20solution%20(Windows%20LAPS)%27"
        $hasPolicy = [bool](Invoke-MgGraphRequest -Uri $policyUri -Method GET -OutputType PSObject).value
        $scripts = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts?`$filter=displayName eq 'Windows LAPS User'").value
        $hasPolicy -and $scripts.Count -gt 0
    }

    $checks = @($authPolicyStatus, $adminConsentStatus, $intuneScopeStatus, $dynGroupsStatus, $blockPersonalStatus,
                $bitLockerStatus, $winComplianceStatus, $iosProtectionStatus, $androidProtectionStatus, $caStatus, $lapsStatus)
    $configuredCount = @($checks | Where-Object { $_ -eq 'Configured' }).Count
    $unknownCount    = @($checks | Where-Object { $_ -like 'Unknown*' }).Count
    $overall = "$configuredCount/$($checks.Count) configured"
    if ($unknownCount -gt 0) { $overall += ", $unknownCount unknown (permissions)" }

    Write-Host "  $overall — $($org.DisplayName)" -ForegroundColor (if ($configuredCount -eq $checks.Count) { 'Green' } else { 'Yellow' })

    $allRows.Add([pscustomobject]@{
        Domain                    = $rec.Domain
        TenantId                  = $org.Id
        TenantName                = $org.DisplayName
        AuthorizationPolicy       = $authPolicyStatus
        AdminConsentWorkflow      = $adminConsentStatus
        IntuneMDMScope            = $intuneScopeStatus
        DynamicDeviceGroups       = $dynGroupsStatus
        BlockPersonalEnrollment   = $blockPersonalStatus
        BitLockerPolicy           = $bitLockerStatus
        WindowsCompliancePolicy   = $winComplianceStatus
        iOSAppProtection          = $iosProtectionStatus
        AndroidAppProtection      = $androidProtectionStatus
        ConditionalAccessPolicies = $caStatus
        LAPSPolicy                = $lapsStatus
        OverallStatus             = $overall
    })
    $ok++
}

try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}

if ($allRows.Count -gt 0) {
    $allRows | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8 -Force
    Write-Host ""
    Write-Host "Report written: $OutputPath" -ForegroundColor Green
    Write-Output "##OPEN_FILE:$OutputPath##"
}

Write-Host ""
Write-Host "=== Done ===" -ForegroundColor Cyan
Write-Host "Tenants checked: $ok  |  Auth failed: $failed  |  No app registration: $skippedNoCreds"
Write-Host "NOTE: Exchange Online settings (EOP/ATP preset policies, transport rules, audit logging) are not covered by this report — see script help." -ForegroundColor DarkGray

if ($failed -gt 0) { exit 1 }

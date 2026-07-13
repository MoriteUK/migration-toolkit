<#
.SYNOPSIS
    All-in-one baseline configuration script (combines scripts 1-15).
    Uses ONE Microsoft Graph login and ONE Exchange Online login for the entire run.
    Every step checks whether its policy/object already exists before creating anything,
    so it's safe to re-run against a tenant that's already been partially or fully configured.

.NOTES
    Run this single file. It will:
      1. Kill other pwsh/powershell sessions and relaunch itself in a clean pwsh 7 window
         (avoids the module/assembly conflicts you've hit before).
      2. Verify/install Microsoft.Graph, Microsoft.Graph.Beta, ExchangeOnlineManagement.
      3. Connect to Graph ONCE with the full scope set every step needs.
      4. Run steps 2-11 (Entra ID / Intune, Graph-based).
      5. Prompt once for the admin UPN, connect to Exchange Online ONCE.
      6. Run steps 12-13 (EOP/Defender preset policies, Exchange Online org settings).
      7. Run steps 14-15 (LAPS, Conditional Access verification) back on Graph.
      8. Print a pass/fail/skip summary and write one consolidated log file.
#>

param(
    [switch]$Restarted,
    [string]$AdminUPN,
    [string]$LogPath
)

# Default log locations live under the app's own writable log folder (not the
# script folder, which is read-only in a perMachine install).
$script:BaselineLogDir = Join-Path $env:APPDATA 'FlyMigration\Logs'
if (-not (Test-Path $script:BaselineLogDir)) { New-Item -ItemType Directory -Path $script:BaselineLogDir -Force | Out-Null }
if (-not $LogPath) { $LogPath = Join-Path $script:BaselineLogDir "BaselineFullRun_$(Get-Date -Format 'yyyyMMdd_HHmmss').log" }

# --- Color-coded logging ---
function Write-Log {
    param(
        [string]$msg,
        [ValidateSet("INFO","OK","ERROR","WARN","SKIP")][string]$Level = "INFO"
    )
    $global:logContent += $msg
    switch ($Level) {
        "INFO"  { Write-Host $msg -ForegroundColor Cyan }
        "OK"    { Write-Host $msg -ForegroundColor Green }
        "ERROR" { Write-Host $msg -ForegroundColor Red }
        "WARN"  { Write-Host $msg -ForegroundColor Yellow }
        "SKIP"  { Write-Host $msg -ForegroundColor DarkYellow }
        default { Write-Host $msg }
    }
}

$global:summary = [System.Collections.Generic.List[object]]::new()
function Add-Summary {
    param([int]$Num, [string]$Name, [string]$Result)
    $global:summary.Add([pscustomobject]@{ Num = $Num; Step = $Name; Result = $Result })
}

# --- Restart Logic (clean session, matches script 1) ---
if (-not $Restarted) {
    if (-not (Get-Command pwsh -ErrorAction SilentlyContinue)) {
        Write-Host "[ERROR] PowerShell 7 (pwsh) is not available in PATH." -ForegroundColor Red
        return
    }

    Write-Host "Terminating other PowerShell sessions for a clean environment..." -ForegroundColor Yellow
    $scriptPath = $MyInvocation.MyCommand.Definition
    $currentId = $PID

    $psProcesses = Get-Process -Name pwsh,powershell -ErrorAction SilentlyContinue | Where-Object {
        $_.Id -ne $currentId
    }
    foreach ($proc in $psProcesses) {
        try {
            Write-Host (" - Terminating PID {0} ({1})" -f $proc.Id, $proc.ProcessName) -ForegroundColor Red
            $proc.Kill()
        } catch {}
    }

    Write-Host "Launching a new PowerShell 7 session..." -ForegroundColor Green
    $relaunchArgs = @('-NoExit', '-File', "`"$scriptPath`"", '-Restarted', '-LogPath', "`"$LogPath`"")
    if ($AdminUPN) { $relaunchArgs += @('-AdminUPN', "`"$AdminUPN`"") }
    Start-Process pwsh -ArgumentList $relaunchArgs
    exit
}

# --- Log Init ---
$global:logContent = @()
$global:logContent += "Timestamp: $(Get-Date -Format 's')"
$startTime = Get-Date

Write-Host ""
Write-Host "=================================================" -ForegroundColor Cyan
Write-Host "  Baseline Full Run (all steps, one login each)" -ForegroundColor Cyan
Write-Host "  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
Write-Host "=================================================" -ForegroundColor Cyan
Write-Host ""

# --- PowerShell Version Check ---
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Log "[ERROR] PowerShell $($PSVersionTable.PSVersion) is not supported. PowerShell 7+ required." "ERROR"
    $global:logContent -join "`n" | Out-File -FilePath $LogPath -Encoding utf8
    return
} else {
    Write-Log "[OK] PowerShell $($PSVersionTable.PSVersion) — OK" "OK"
}

# --- Module Handler ---
function Ensure-Module {
    param([string]$ModuleName, [string]$MinVersion = "2.28.0")
    $mod = Get-Module -ListAvailable -Name $ModuleName | Sort-Object Version -Descending | Select-Object -First 1
    if (!$mod -or ([version]$mod.Version -lt [version]$MinVersion)) {
        Write-Log "[INFO] $ModuleName — installing/updating to >= $MinVersion" "WARN"
        try {
            Install-Module $ModuleName -MinimumVersion $MinVersion -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            $newMod = Get-Module -ListAvailable -Name $ModuleName | Sort-Object Version -Descending | Select-Object -First 1
            Write-Log "[OK] $ModuleName installed — version $($newMod.Version)" "OK"
        } catch {
            Write-Log "[ERROR] Failed to install ${ModuleName}: $($_.Exception.Message)" "ERROR"
            $global:logContent -join "`n" | Out-File -FilePath $LogPath -Encoding utf8
            return
        }
    } else {
        Write-Log "[OK] $ModuleName $($mod.Version) — up to date" "OK"
    }
}

Write-Host "Verifying PowerShell modules..." -ForegroundColor Cyan
Ensure-Module -ModuleName "Microsoft.Graph" -MinVersion "2.28.0"
Ensure-Module -ModuleName "Microsoft.Graph.Beta" -MinVersion "2.28.0"
Ensure-Module -ModuleName "ExchangeOnlineManagement" -MinVersion "3.0.0"
Write-Log "[OK] All modules verified." "OK"

# --- Single Graph connection for the whole run ---
$AllScopes = @(
    "DeviceManagementConfiguration.ReadWrite.All"
    "DeviceManagementManagedDevices.ReadWrite.All"
    "Policy.ReadWrite.DeviceConfiguration"
    "Policy.ReadWrite.ConditionalAccess"
    "Policy.ReadWrite.MobilityManagement"
    "DeviceManagementRBAC.ReadWrite.All"
    "DeviceManagementApps.ReadWrite.All"
    "DeviceManagementServiceConfig.ReadWrite.All"
    "DeviceManagementManagedDevices.PrivilegedOperations.All"
    "Directory.ReadWrite.All"
    "User.ReadWrite.All"
    "Group.ReadWrite.All"
    "Application.ReadWrite.All"
    "RoleManagement.ReadWrite.Directory"
    "Directory.AccessAsUser.All"
    "SecurityEvents.ReadWrite.All"
    "SecurityActions.ReadWrite.All"
    "Reports.Read.All"
    "Mail.ReadWrite"
    "MailboxSettings.ReadWrite"
    "Sites.ReadWrite.All"
    "Files.ReadWrite.All"
    "Calendars.ReadWrite"
    "Contacts.ReadWrite"
    "Team.ReadBasic.All"
    "ChannelSettings.ReadWrite.All"
    "User.Read"
    "User.ReadBasic.All"
    "offline_access"
    "Policy.ReadWrite.ConsentRequest"
    "DeviceManagementScripts.ReadWrite.All"
    "Policy.Read.All"
)

Write-Host "`nConnecting to Microsoft Graph ($($AllScopes.Count) scopes)..." -ForegroundColor Cyan
try {
    Connect-MgGraph -Scopes $AllScopes -NoWelcome -ErrorAction Stop
} catch {
    if ($_.Exception.Message -match "InteractiveBrowserCredential|WithLogging|NullReference|WAM|broker") {
        Write-Log "[WARN] Interactive auth failed (MSAL conflict), retrying with device code..." "WARN"
        try {
            Connect-MgGraph -Scopes $AllScopes -NoWelcome -UseDeviceAuthentication -ErrorAction Stop
        } catch {
            Write-Log "[ERROR] Device code auth also failed: $($_.Exception.Message)" "ERROR"
            $global:logContent -join "`n" | Out-File -FilePath $LogPath -Encoding utf8
            return
        }
    } else {
        Write-Log "[ERROR] Exception during Connect-MgGraph: $($_.Exception.Message)" "ERROR"
        $global:logContent -join "`n" | Out-File -FilePath $LogPath -Encoding utf8
        return
    }
}

$context = Get-MgContext
if (-not $context) {
    Write-Log "[ERROR] Graph connection failed. No context returned." "ERROR"
    $global:logContent -join "`n" | Out-File -FilePath $LogPath -Encoding utf8
    return
}
Write-Log "[OK] Connected to Graph as: $($context.Account) (tenant $($context.TenantId))" "OK"

# --- Rename the log file to include the tenant name, now that we know it ---
try {
    $org = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/organization?`$select=displayName" -ErrorAction Stop
    $tenantName = $org.value[0].displayName -replace '[^a-zA-Z0-9]+', '-'
    $tenantName = $tenantName.Trim('-')
} catch {
    $tenantName = $context.TenantId
}
$LogPath = Join-Path $script:BaselineLogDir "BaselineFullRun_${tenantName}_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
Write-Log "[INFO] Log file for this run: $LogPath" "INFO"

# =========================================================================
# STEP 2 — Authorization Policy (Entra ID user settings)
# =========================================================================
Write-Host "`n=== Step 2: Configure Entra ID User Settings ===" -ForegroundColor Cyan
try {
    $authPolicyUri = "https://graph.microsoft.com/v1.0/policies/authorizationPolicy"
    $currentAuthPolicy = Invoke-MgGraphRequest -Method GET -Uri $authPolicyUri

    $alreadySet = (
        $currentAuthPolicy.defaultUserRolePermissions.allowedToCreateApps -eq $false -and
        $currentAuthPolicy.defaultUserRolePermissions.allowedToCreateSecurityGroups -eq $false -and
        $currentAuthPolicy.defaultUserRolePermissions.allowedToCreateTenants -eq $false -and
        $currentAuthPolicy.allowUserConsentForRiskyApps -eq $false -and
        $currentAuthPolicy.allowEmailVerifiedUsersToJoinOrg -eq $false -and
        $currentAuthPolicy.allowInvitesFrom -eq "adminsAndGuestInviters" -and
        $currentAuthPolicy.allowUserConsentForApps -eq $false -and
        $currentAuthPolicy.guestUserRoleId -eq "2af84b1e-32c8-42b7-82bc-daa82404023b"
    )

    if ($alreadySet) {
        Write-Log "[SKIP] Step 2: Authorization policy already matches baseline. Skipping." "SKIP"
        Add-Summary 2 "Configure_EntraID_User_Settings" "Skipped (already applied)"
    } else {
        $authPayload = @{
            defaultUserRolePermissions = @{
                allowedToCreateApps             = $false
                allowedToCreateSecurityGroups   = $false
                permissionGrantPoliciesAssigned = @()
                allowedToCreateTenants          = $false
            }
            allowUserConsentForRiskyApps     = $false
            allowEmailVerifiedUsersToJoinOrg = $false
            allowInvitesFrom                 = "adminsAndGuestInviters"
            allowUserConsentForApps          = $false
            guestUserRoleId                  = "2af84b1e-32c8-42b7-82bc-daa82404023b"
        }
        Invoke-MgGraphRequest -Method PATCH -Uri $authPolicyUri `
            -Body ($authPayload | ConvertTo-Json -Depth 5) -ContentType "application/json" -ErrorAction Stop

        Write-Log "[OK] Step 2: Authorization policy updated." "OK"
        Add-Summary 2 "Configure_EntraID_User_Settings" "Success"
    }
} catch {
    Write-Log "[ERROR] Step 2 failed: $($_.Exception.Message)" "ERROR"
    Add-Summary 2 "Configure_EntraID_User_Settings" "Error: $($_.Exception.Message)"
}

# =========================================================================
# STEP 3 — Admin User Consent Workflow
# =========================================================================
Write-Host "`n=== Step 3: Configure Admin User Consent Workflow ===" -ForegroundColor Cyan
try {
    $reviewerGroupName         = "Admin Consent Reviewers"
    $reviewerGroupMailNickname = "AdminConsentReviewers"

    $existingGroup = Get-MgGroup -Filter "displayName eq '$reviewerGroupName'" -ErrorAction Stop | Select-Object -First 1
    if ($existingGroup) {
        Write-Log "[SKIP] Step 3: Group '$reviewerGroupName' already exists (ID: $($existingGroup.Id))." "SKIP"
        $group = $existingGroup
    } else {
        $group = New-MgGroup -BodyParameter @{
            DisplayName     = $reviewerGroupName
            MailEnabled     = $true
            MailNickname    = $reviewerGroupMailNickname
            SecurityEnabled = $true
            GroupTypes      = @("Unified")
            Visibility      = "Private"
        }
        Write-Log "[OK] Step 3: Created group '$reviewerGroupName' (ID: $($group.Id))." "OK"
    }

    $adminConsentPolicyUri  = "https://graph.microsoft.com/v1.0/policies/adminConsentRequestPolicy"
    $authorizationPolicyUri = "https://graph.microsoft.com/v1.0/policies/authorizationPolicy"
    $currentConsentPolicy   = Invoke-MgGraphRequest -Method GET -Uri $adminConsentPolicyUri

    $consentAlreadySet = (
        $currentConsentPolicy.isEnabled -eq $true -and
        ($currentConsentPolicy.reviewers | Where-Object { $_.query -eq "/groups/$($group.Id)/transitiveMembers" })
    )

    if ($consentAlreadySet) {
        Write-Log "[SKIP] Step 3: Admin Consent Request Policy already enabled with this reviewer group." "SKIP"
        Add-Summary 3 "Configure_Admin_User_Consent_Workflow" "Skipped (already applied)"
    } else {
        $consentPayload = @{
            isEnabled             = $true
            notifyReviewers       = $true
            remindersEnabled      = $true
            requestDurationInDays = 5
            reviewers             = @(@{ query = "/groups/$($group.Id)/transitiveMembers"; queryType = "MicrosoftGraph" })
        }
        $consentBody = $consentPayload | ConvertTo-Json -Depth 5
        $success = $false
        for ($attempt = 1; $attempt -le 3; $attempt++) {
            try {
                Invoke-MgGraphRequest -Method PUT -Uri $adminConsentPolicyUri -Body $consentBody -ContentType "application/json" -ErrorAction Stop
                $success = $true
                break
            } catch {
                if ($attempt -lt 3) { Start-Sleep -Seconds 5 }
            }
        }

        $authzPayload = @{ defaultUserRolePermissions = @{ permissionGrantPoliciesAssigned = @() } }
        Invoke-MgGraphRequest -Method PATCH -Uri $authorizationPolicyUri `
            -Body ($authzPayload | ConvertTo-Json -Depth 5) -ContentType "application/json" -ErrorAction Stop

        if ($success) {
            Write-Log "[OK] Step 3: Admin Consent Workflow enabled, user consent disabled." "OK"
            Add-Summary 3 "Configure_Admin_User_Consent_Workflow" "Success"
        } else {
            Write-Log "[ERROR] Step 3: Admin Consent Policy PUT failed after 3 attempts." "ERROR"
            Add-Summary 3 "Configure_Admin_User_Consent_Workflow" "Error: PUT failed after retries"
        }
    }
} catch {
    Write-Log "[ERROR] Step 3 failed: $($_.Exception.Message)" "ERROR"
    Add-Summary 3 "Configure_Admin_User_Consent_Workflow" "Error: $($_.Exception.Message)"
}

# =========================================================================
# STEP 4 — Intune MDM User Scope
# =========================================================================
Write-Host "`n=== Step 4: Set Intune MDM User Scope ===" -ForegroundColor Cyan
try {
    $uri = "https://graph.microsoft.com/beta/policies/mobileDeviceManagementPolicies"
    $response = Invoke-MgGraphRequest -Method GET -Uri $uri
    $policy = $response.value | Where-Object { $_.displayName -eq "Microsoft Intune" }

    if (-not $policy) {
        Write-Log "[ERROR] Step 4: Microsoft Intune MDM policy not found." "ERROR"
        Add-Summary 4 "Set_Intune_MDM_User_Scope" "Error: policy not found"
    } elseif ($policy.appliesTo -eq "all") {
        Write-Log "[SKIP] Step 4: MDM user scope already set to ALL." "SKIP"
        Add-Summary 4 "Set_Intune_MDM_User_Scope" "Skipped (already applied)"
    } else {
        $policyUri = "https://graph.microsoft.com/beta/policies/mobileDeviceManagementPolicies/$($policy.id)"
        $body = @{ appliesTo = "all" } | ConvertTo-Json
        Invoke-MgGraphRequest -Method PATCH -Uri $policyUri -Body $body -ContentType "application/json" -ErrorAction Stop
        Write-Log "[OK] Step 4: MDM user scope set to ALL users." "OK"
        Add-Summary 4 "Set_Intune_MDM_User_Scope" "Success"
    }
} catch {
    Write-Log "[ERROR] Step 4 failed: $($_.Exception.Message)" "ERROR"
    Add-Summary 4 "Set_Intune_MDM_User_Scope" "Error: $($_.Exception.Message)"
}

# =========================================================================
# STEP 5 — Dynamic Device Groups
# =========================================================================
Write-Host "`n=== Step 5: Create Dynamic Device Groups ===" -ForegroundColor Cyan
function New-DynamicGroupIfMissing {
    param([string]$DisplayName, [string]$Description, [string]$MembershipRule)
    $checkUri = "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$DisplayName'"
    try {
        $existing = (Invoke-MgGraphRequest -Method GET -Uri $checkUri).value
    } catch {
        Write-Log "[ERROR] Step 5: Failed to check group '$DisplayName': $($_.Exception.Message)" "ERROR"
        return "Error"
    }
    if ($existing) {
        Write-Log "[SKIP] Step 5: Group '$DisplayName' already exists." "SKIP"
        return "Skipped"
    }
    $mailNickname = ($DisplayName -replace '[^a-zA-Z0-9]', '').ToLower()
    $body = @{
        displayName     = $DisplayName
        description     = $Description
        mailEnabled     = $false
        mailNickname    = $mailNickname
        securityEnabled = $true
        groupTypes      = @("DynamicMembership")
        membershipRule  = $MembershipRule
        membershipRuleProcessingState = "On"
    } | ConvertTo-Json -Depth 5
    try {
        Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groups" -Body $body -ContentType "application/json" -ErrorAction Stop
        Write-Log "[OK] Step 5: Created group '$DisplayName'." "OK"
        return "Success"
    } catch {
        Write-Log "[ERROR] Step 5: Failed to create '$DisplayName': $($_.Exception.Message)" "ERROR"
        return "Error"
    }
}

$dynGroupDefs = @(
    @{ DisplayName = "dyn-byod-android-devices";  Description = "All Android BYOD devices will be automatically added"; MembershipRule = '(device.managementType -eq "MDM") and (device.deviceOwnership -eq "Personal") and (device.deviceOSType -eq "AndroidForWork")' }
    @{ DisplayName = "dyn-corp-android-devices";  Description = "All Android Corporate devices will be automatically added"; MembershipRule = '(device.managementType -eq "MDM") and (device.deviceOwnership -eq "Company") and (device.deviceOSType -startsWith "Android")' }
    @{ DisplayName = "dyn-autopilot-devices";     Description = "All Autopilot devices will be automatically added"; MembershipRule = '(device.devicePhysicalIDs -any (_ -contains "[ZTDID]"))' }
    @{ DisplayName = "dyn-byod-iOSiPad-devices";  Description = "All iOS BYOD devices will be automatically added"; MembershipRule = '(device.managementType -eq "MDM") and (device.deviceOwnership -eq "Personal") and ((device.deviceOSType -eq "iOS") or (device.deviceOSType -eq "iPad"))' }
    @{ DisplayName = "dyn-corp-iOSiPad-devices";  Description = "All iOS Corporate devices will be automatically added"; MembershipRule = '(device.managementType -eq "MDM") and (device.deviceOwnership -eq "Company") and ((device.deviceOSType -eq "iOS") or (device.deviceOSType -eq "iPad"))' }
    @{ DisplayName = "dyn-corp-win10-devices";    Description = "All Windows 10 Corporate devices will be automatically added"; MembershipRule = '(device.deviceOSVersion -startsWith "10.0") and (device.managementType -eq "MDM") and (device.deviceOwnership -eq "Company")' }
    @{ DisplayName = "dyn-corp-win11-devices";    Description = "All Windows 11 Corporate devices will be automatically added"; MembershipRule = '(device.deviceOSVersion -startsWith "10.0.2") and (device.managementType -eq "MDM") and (device.deviceOwnership -eq "Company")' }
    @{ DisplayName = "dyn-byod-macOS-devices";    Description = "All MacOS BYOD devices will be automatically added"; MembershipRule = '(device.managementType -eq "MDM") and (device.deviceOwnership -eq "Personal") and (device.deviceOSType -eq "MacMDM")' }
)

$step5Results = foreach ($def in $dynGroupDefs) {
    New-DynamicGroupIfMissing -DisplayName $def.DisplayName -Description $def.Description -MembershipRule $def.MembershipRule
}
if ($step5Results -contains "Error") {
    Add-Summary 5 "Create_Dynamic_Device_Groups" "Completed with errors (see log)"
} elseif ($step5Results -contains "Success") {
    Add-Summary 5 "Create_Dynamic_Device_Groups" "Success (some groups created)"
} else {
    Add-Summary 5 "Create_Dynamic_Device_Groups" "Skipped (all groups already existed)"
}

# =========================================================================
# STEP 6 — Block Personal Device Enrollment
# =========================================================================
Write-Host "`n=== Step 6: Block Personal Device Enrollment ===" -ForegroundColor Cyan
try {
    $defaultConfig = Get-MgDeviceManagementDeviceEnrollmentConfiguration | Where-Object { $_.Id -like "*_DefaultPlatformRestrictions" }
    if (-not $defaultConfig) {
        Write-Log "[ERROR] Step 6: Default platform restrictions configuration not found." "ERROR"
        Add-Summary 6 "Block_Personal_Device_Enrollment" "Error: config not found"
    } else {
        $currentConfig = Get-MgDeviceManagementDeviceEnrollmentConfiguration -DeviceEnrollmentConfigurationId $defaultConfig.Id
        $alreadyBlocked = (
            $currentConfig.AdditionalProperties.androidRestriction.personalDeviceEnrollmentBlocked -eq $true -and
            $currentConfig.AdditionalProperties.androidForWorkRestriction.personalDeviceEnrollmentBlocked -eq $true -and
            $currentConfig.AdditionalProperties.iosRestriction.personalDeviceEnrollmentBlocked -eq $true -and
            $currentConfig.AdditionalProperties.macOSRestriction.personalDeviceEnrollmentBlocked -eq $true -and
            $currentConfig.AdditionalProperties.windowsRestriction.personalDeviceEnrollmentBlocked -eq $true
        )

        if ($alreadyBlocked) {
            Write-Log "[SKIP] Step 6: Personal device enrollment already blocked on all platforms." "SKIP"
            Add-Summary 6 "Block_Personal_Device_Enrollment" "Skipped (already applied)"
        } else {
            $body = @{
                "@odata.type" = "#microsoft.graph.deviceEnrollmentPlatformRestrictionsConfiguration"
                displayName = $defaultConfig.displayName
                description = "Block personal device enrollment for all platforms"
                androidRestriction        = @{ platformBlocked = $false; personalDeviceEnrollmentBlocked = $true; osMinimumVersion = ""; osMaximumVersion = "" }
                androidForWorkRestriction = @{ platformBlocked = $false; personalDeviceEnrollmentBlocked = $true; osMinimumVersion = $null; osMaximumVersion = $null }
                iosRestriction            = @{ platformBlocked = $false; personalDeviceEnrollmentBlocked = $true; osMinimumVersion = ""; osMaximumVersion = "" }
                macOSRestriction          = @{ platformBlocked = $false; personalDeviceEnrollmentBlocked = $true; osMinimumVersion = $null; osMaximumVersion = $null }
                windowsRestriction        = @{ platformBlocked = $false; personalDeviceEnrollmentBlocked = $true; osMinimumVersion = ""; osMaximumVersion = "" }
            }
            Update-MgDeviceManagementDeviceEnrollmentConfiguration -DeviceEnrollmentConfigurationId $defaultConfig.Id -BodyParameter $body -ErrorAction Stop
            Write-Log "[OK] Step 6: Personal device enrollment blocked on all platforms." "OK"
            Add-Summary 6 "Block_Personal_Device_Enrollment" "Success"
        }
    }
} catch {
    Write-Log "[ERROR] Step 6 failed: $($_.Exception.Message)" "ERROR"
    Add-Summary 6 "Block_Personal_Device_Enrollment" "Error: $($_.Exception.Message)"
}

# =========================================================================
# STEP 7 — BitLocker Compliance Policy
# =========================================================================
Write-Host "`n=== Step 7: Deploy BitLocker Compliance Policy ===" -ForegroundColor Cyan
try {
    $policyDisplayName = "Enforce BitLocker Encryption"
    $existingPolicies = Get-MgDeviceManagementDeviceConfiguration | Where-Object { $_.displayName -eq $policyDisplayName }

    if ($existingPolicies.Count -gt 0) {
        Write-Log "[SKIP] Step 7: Policy '$policyDisplayName' already exists (ID: $($existingPolicies[0].Id))." "SKIP"
        Add-Summary 7 "Deploy_BitLocker_Compliance_Policy" "Skipped (already applied)"
    } else {
        $bitLockerConfig = @{
            "@odata.type" = "#microsoft.graph.windows10EndpointProtectionConfiguration"
            displayName = $policyDisplayName
            description = "Requires BitLocker encryption and configures recovery key backup to Azure AD."
            bitLockerEncryptDevice = $true
            bitLockerSystemDriveEncryptionMethod = "xtsAes256"
            bitLockerSystemDrivePolicy = @{
                encryptionMethod = "xtsAes256"; requireEncryption = $true
                recoveryOptions = @{ blockDataRecoveryAgent = $false; recoveryPasswordEnabled = $true; recoveryKeyUsage = "required"; recoveryInformationToStore = "storeRecoveryPasswordAndKeyPackage"; enableRecoveryInformationSaveToAzureAd = $true }
            }
            bitLockerFixedDrivePolicy = @{
                encryptionMethod = "xtsAes256"; requireEncryption = $false
                recoveryOptions = @{ recoveryPasswordEnabled = $true; recoveryKeyUsage = "optional"; enableRecoveryInformationSaveToAzureAd = $true }
            }
        }
        $policy = New-MgDeviceManagementDeviceConfiguration -BodyParameter $bitLockerConfig -ErrorAction Stop
        Write-Log "[OK] Step 7: BitLocker policy created (ID: $($policy.Id))." "OK"
        Add-Summary 7 "Deploy_BitLocker_Compliance_Policy" "Success"
    }
} catch {
    Write-Log "[ERROR] Step 7 failed: $($_.Exception.Message)" "ERROR"
    Add-Summary 7 "Deploy_BitLocker_Compliance_Policy" "Error: $($_.Exception.Message)"
}

# =========================================================================
# STEP 8 — Windows Compliance Policy
# =========================================================================
Write-Host "`n=== Step 8: Deploy Windows Compliance Policy ===" -ForegroundColor Cyan
try {
    $policyName = "Baseline Compliance Policy - Windows 10/11"
    $existingPolicies = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies"
    $alreadyExists = $existingPolicies.value | Where-Object { $_.displayName -eq $policyName }

    if ($alreadyExists) {
        Write-Log "[SKIP] Step 8: Policy '$policyName' already exists (ID: $($alreadyExists.id))." "SKIP"
        Add-Summary 8 "Deploy_Windows_Compliance_Policy" "Skipped (already applied)"
    } else {
        $compliancePolicy = @{
            "@odata.type" = "#microsoft.graph.windows10CompliancePolicy"
            displayName = $policyName
            description = "Enforces baseline security settings for Windows devices."
            passwordRequired = $true
            passwordMinimumLength = 8
            passwordRequiredToUnlockFromIdle = $true
            passwordMinutesOfInactivityBeforeLock = 15
            passwordExpirationDays = 90
            passwordPreviousPasswordBlockCount = 5
            osMinimumVersion = "10.0.19041.0"
            bitLockerEnabled = $true
            secureBootEnabled = $true
            codeIntegrityEnabled = $true
            storageRequireEncryption = $true
            deviceThreatProtectionEnabled = $true
            deviceThreatProtectionRequiredSecurityLevel = "secured"
            defenderEnabled = $true
            defenderSignatureUpToDate = $true
            firewallEnabled = $true
            scheduledActionsForRule = @(@{ ruleName = "PasswordRequired"; scheduledActionConfigurations = @(@{ actionType = "block"; gracePeriodHours = 0 }) })
        }
        $response = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies" `
            -Body ($compliancePolicy | ConvertTo-Json -Depth 5 -Compress) -ContentType "application/json" -ErrorAction Stop

        Write-Log "[OK] Step 8: Compliance policy created (ID: $($response.id))." "OK"
        Add-Summary 8 "Deploy_Windows_Compliance_Policy" "Success"
    }
} catch {
    Write-Log "[ERROR] Step 8 failed: $($_.Exception.Message)" "ERROR"
    Add-Summary 8 "Deploy_Windows_Compliance_Policy" "Error: $($_.Exception.Message)"
}

# =========================================================================
# STEP 9 — iOS App Protection Policy
# =========================================================================
Write-Host "`n=== Step 9: Configure iOS App Protection Policy ===" -ForegroundColor Cyan
try {
    $desiredPolicyName = "iOS App Protection Policy"
    $currentPolicies = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/deviceAppManagement/iosManagedAppProtections"
    $existingPolicy = $currentPolicies.value | Where-Object { $_.displayName -eq $desiredPolicyName }

    if ($existingPolicy) {
        Write-Log "[SKIP] Step 9: Policy '$desiredPolicyName' already exists (ID: $($existingPolicy.id))." "SKIP"
        Add-Summary 9 "Configure_iOS_App_Protection_Policy" "Skipped (already applied)"
    } else {
        $body = @{
            "@odata.type" = "#microsoft.graph.iosManagedAppProtection"
            displayName = $desiredPolicyName
            description = "App Protection (MAM) policy for iOS with strict access and PIN requirements"
            periodOfflineBeforeAccessCheck = "P1D"; periodOnlineBeforeAccessCheck = "PT30M"
            pinRequired = $true; simplePinBlocked = $true; minimumPinLength = 6; pinCharacterSet = "numeric"
            fingerprintBlocked = $false; faceIdBlocked = $false; maximumPinRetries = 5
            periodBeforePinReset = "PT0S"; appPinWhenDevicePinIsSet = $true; deviceComplianceRequired = $true
            managedBrowserToOpenLinksRequired = $false; managedBrowser = "notConfigured"; minimumRequiredOsVersion = "18.0"
            recheckAccessAfterInactivityInMinutes = 30
            allowedInboundDataTransferSources = "managedApps"; allowedOutboundDataTransferDestinations = "managedApps"
            allowedOutboundClipboardSharingLevel = "managedApps"; allowedDataStorageLocations = @("oneDriveForBusiness", "sharePoint")
            dataBackupBlocked = $true; saveAsBlocked = $true; periodOfflineBeforeWipeIsEnforced = "P90D"
            contactSyncBlocked = $true; printBlocked = $true; disableAppPinIfDevicePinIsSet = $false
            organizationalCredentialsRequired = $false
        }
        $jsonBody = $body | ConvertTo-Json -Depth 10
        $response = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/deviceAppManagement/iosManagedAppProtections" -Body $jsonBody -ContentType "application/json" -ErrorAction Stop
        $policyId = $response.id
        $targetAppsBody = @{ apps = @(); appGroupType = "allCoreMicrosoftApps" } | ConvertTo-Json -Depth 4
        Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/deviceAppManagement/iosManagedAppProtections/$policyId/targetApps" -Body $targetAppsBody -ContentType "application/json" -ErrorAction Stop

        Write-Log "[OK] Step 9: iOS MAM policy created (ID: $policyId) and targeted to core Microsoft apps." "OK"
        Add-Summary 9 "Configure_iOS_App_Protection_Policy" "Success"
    }
} catch {
    Write-Log "[ERROR] Step 9 failed: $($_.Exception.Message)" "ERROR"
    Add-Summary 9 "Configure_iOS_App_Protection_Policy" "Error: $($_.Exception.Message)"
}

# =========================================================================
# STEP 10 — Android App Protection Policy
# =========================================================================
Write-Host "`n=== Step 10: Configure Android App Protection Policy ===" -ForegroundColor Cyan
try {
    $desiredPolicyName = "Android App Protection Policy"
    $currentPolicies = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/deviceAppManagement/androidManagedAppProtections"
    $existingPolicy = $currentPolicies.value | Where-Object { $_.displayName -eq $desiredPolicyName }

    if ($existingPolicy) {
        Write-Log "[SKIP] Step 10: Policy '$desiredPolicyName' already exists (ID: $($existingPolicy.id))." "SKIP"
        Add-Summary 10 "Configure_Android_App_Protection_Policy" "Skipped (already applied)"
    } else {
        $body = @{
            "@odata.type" = "#microsoft.graph.androidManagedAppProtection"
            displayName = $desiredPolicyName
            description = "App Protection (MAM) policy for Android with strict access and PIN requirements"
            periodOfflineBeforeAccessCheck = "P1D"; periodOnlineBeforeAccessCheck = "PT30M"
            pinRequired = $true; simplePinBlocked = $true; minimumPinLength = 6; pinCharacterSet = "numeric"
            fingerprintBlocked = $false; maximumPinRetries = 5; periodBeforePinReset = "PT0S"
            appPinWhenDevicePinIsSet = $true; deviceComplianceRequired = $true
            managedBrowserToOpenLinksRequired = $false; managedBrowser = "notConfigured"; minimumRequiredOsVersion = "14.0"
            recheckAccessAfterInactivityInMinutes = 30
            allowedInboundDataTransferSources = "managedApps"; allowedOutboundDataTransferDestinations = "managedApps"
            allowedOutboundClipboardSharingLevel = "managedApps"; allowedDataStorageLocations = @("oneDriveForBusiness", "sharePoint")
            dataBackupBlocked = $true; saveAsBlocked = $true; periodOfflineBeforeWipeIsEnforced = "P90D"
            contactSyncBlocked = $true; printBlocked = $true; disableAppPinIfDevicePinIsSet = $false
            organizationalCredentialsRequired = $false
        }
        $jsonBody = $body | ConvertTo-Json -Depth 10
        $response = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/deviceAppManagement/androidManagedAppProtections" -Body $jsonBody -ContentType "application/json" -ErrorAction Stop
        $policyId = $response.id
        $targetAppsBody = @{ apps = @(); appGroupType = "allCoreMicrosoftApps" } | ConvertTo-Json -Depth 4
        Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/deviceAppManagement/androidManagedAppProtections/$policyId/targetApps" -Body $targetAppsBody -ContentType "application/json" -ErrorAction Stop

        Write-Log "[OK] Step 10: Android MAM policy created (ID: $policyId) and targeted to core Microsoft apps." "OK"
        Add-Summary 10 "Configure_Android_App_Protection_Policy" "Success"
    }
} catch {
    Write-Log "[ERROR] Step 10 failed: $($_.Exception.Message)" "ERROR"
    Add-Summary 10 "Configure_Android_App_Protection_Policy" "Error: $($_.Exception.Message)"
}

# =========================================================================
# STEP 11 — Conditional Access Policies
# =========================================================================
Write-Host "`n=== Step 11: Deploy Conditional Access Policies ===" -ForegroundColor Cyan
try {
    $trustedCountries = @("CA", "US", "GB")
    $locationMap = @{}
    $existingLocations = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations").value

    foreach ($code in $trustedCountries) {
        $name = "Trusted Country - $code"
        $existing = $existingLocations | Where-Object { $_.displayName -eq $name }
        if ($existing) {
            $locationMap[$code] = $existing.id
            Write-Log "[SKIP] Step 11: Named location '$name' already exists." "SKIP"
        } else {
            $body = @{ "@odata.type" = "#microsoft.graph.countryNamedLocation"; displayName = $name; countriesAndRegions = @($code); includeUnknownCountriesAndRegions = $false; isTrusted = $true } | ConvertTo-Json -Depth 10
            $newLoc = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations" -Body $body -ContentType "application/json" -ErrorAction Stop
            $locationMap[$code] = $newLoc.id
            Write-Log "[OK] Step 11: Created named location '$name'." "OK"
            Start-Sleep -Seconds 10
        }
    }

    $existingPolicies = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies").value
    $existingNames = $existingPolicies.displayName

    $policies = @(
        '{"displayName":"Block Access Outside Approved Countries","state":"disabled","conditions":{"users":{"includeUsers":["All"]},"applications":{"includeApplications":["All"]},"locations":{"includeLocations":["All"],"excludeLocations":[]}},"grantControls":{"operator":"OR","builtInControls":["block"]}}'
        '{"displayName":"Block Legacy Authentication","state":"disabled","conditions":{"users":{"includeUsers":["All"]},"applications":{"includeApplications":["All"]},"clientAppTypes":["exchangeActiveSync","other"],"locations":{"includeLocations":["All"]}},"grantControls":{"operator":"OR","builtInControls":["block"]}}'
        '{"displayName":"Require MFA for Admin Portals - 8hr","state":"disabled","conditions":{"users":{"includeUsers":["All"]},"applications":{"includeApplications":["MicrosoftAdminPortals"]},"locations":{"includeLocations":["All"]}},"grantControls":{"operator":"OR","builtInControls":["mfa"]},"sessionControls":{"signInFrequency":{"value":8,"type":"hours","isEnabled":true}}}'
        '{"displayName":"Require MFA for Admin Roles - 8hr","state":"disabled","conditions":{"users":{"includeRoles":["62e90394-69f5-4237-9190-012177145e10"]},"applications":{"includeApplications":["All"]},"locations":{"includeLocations":["All"]}},"grantControls":{"operator":"OR","builtInControls":["mfa"]},"sessionControls":{"signInFrequency":{"value":8,"type":"hours","isEnabled":true}}}'
        '{"displayName":"MFA for All Users - Browser Only - 8hr","state":"disabled","conditions":{"users":{"includeUsers":["all"]},"applications":{"includeApplications":["All"]},"clientAppTypes":["browser"],"locations":{"includeLocations":["All"]}},"grantControls":{"operator":"OR","builtInControls":["mfa"]},"sessionControls":{"signInFrequency":{"value":8,"type":"hours","isEnabled":true}}}'
        '{"displayName":"Require MFA for All Users","state":"disabled","conditions":{"users":{"includeUsers":["All"]},"applications":{"includeApplications":["All"]},"locations":{"includeLocations":["All"]}},"grantControls":{"operator":"OR","builtInControls":["mfa"]}}'
        '{"displayName":"Require MFA for Guest Users - 8hr","state":"disabled","conditions":{"users":{"includeGuestsOrExternalUsers":{"guestOrExternalUserTypes":"internalGuest,b2bCollaborationGuest,b2bCollaborationMember,b2bDirectConnectUser,otherExternalUser,serviceProvider","externalTenants":{"membershipKind":"all"}}},"applications":{"includeApplications":["All"]},"clientAppTypes":["all"],"locations":{"includeLocations":["All"]}},"grantControls":{"operator":"OR","builtInControls":["mfa"]},"sessionControls":{"signInFrequency":{"isEnabled":true,"type":"hours","value":8,"authenticationType":"primaryAndSecondaryAuthentication","frequencyInterval":"timeBased"}}}'
    )

    $createdCount = 0; $skippedCount = 0
    foreach ($rawJson in $policies) {
        $policyObj = $rawJson | ConvertFrom-Json
        $policyName = $policyObj.displayName

        if ($existingNames -contains $policyName) {
            Write-Log "[SKIP] Step 11: '$policyName' already exists." "SKIP"
            $skippedCount++
            continue
        }
        if ($policyName -eq "Block Access Outside Approved Countries") {
            $excludeLocationIds = @()
            foreach ($code in $trustedCountries) { if ($locationMap.ContainsKey($code)) { $excludeLocationIds += $locationMap[$code] } }
            $policyObj.conditions.locations.excludeLocations = $excludeLocationIds
        }
        $body = $policyObj | ConvertTo-Json -Depth 10
        Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies" -Body $body -ContentType "application/json" -ErrorAction Stop
        Write-Log "[OK] Step 11: Created '$policyName'." "OK"
        $createdCount++
    }

    # Verify all are disabled
    $allPolicies = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies").value
    foreach ($rawJson in $policies) {
        $name = ($rawJson | ConvertFrom-Json).displayName
        $policy = $allPolicies | Where-Object { $_.displayName -eq $name }
        if ($policy -and $policy.state -ne "disabled") {
            Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies/$($policy.id)" -Body '{"state":"disabled"}' -ContentType "application/json" -ErrorAction Stop
            Write-Log "[OK] Step 11: Corrected '$name' to disabled." "OK"
        }
    }

    if ($createdCount -eq 0) {
        Add-Summary 11 "Deploy_Conditional_Access_Policies" "Skipped (all policies already existed)"
    } else {
        Add-Summary 11 "Deploy_Conditional_Access_Policies" "Success ($createdCount created, $skippedCount already existed)"
    }
} catch {
    Write-Log "[ERROR] Step 11 failed: $($_.Exception.Message)" "ERROR"
    Add-Summary 11 "Deploy_Conditional_Access_Policies" "Error: $($_.Exception.Message)"
}

# =========================================================================
# Single Exchange Online connection for Steps 12 & 13
# =========================================================================
Write-Host "`nConnecting to Exchange Online (used once, for steps 12 & 13)..." -ForegroundColor Cyan
try {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    if (-not $AdminUPN) { $AdminUPN = Read-Host "Enter admin UPN for Exchange Online (e.g. admin@contoso.com)" }

    try {
        Connect-ExchangeOnline -UserPrincipalName $AdminUPN -ShowBanner:$false -ErrorAction Stop
        Write-Log "[OK] Connected to Exchange Online as $AdminUPN" "OK"
    } catch [System.NullReferenceException] {
        Write-Log "[WARN] WAM broker unavailable, retrying with device code..." "WARN"
        Connect-ExchangeOnline -UserPrincipalName $AdminUPN -Device -ShowBanner:$false -ErrorAction Stop
        Write-Log "[OK] Connected to Exchange Online as $AdminUPN (device code)" "OK"
    }
} catch {
    Write-Log "[ERROR] Could not connect to Exchange Online: $($_.Exception.Message). Steps 12 & 13 will be skipped." "ERROR"
    Add-Summary 12 "Enable_EOP_Default_Security_Policies" "Skipped (no EXO connection)"
    Add-Summary 13 "Configure_Exchange_Online_Settings" "Skipped (no EXO connection)"
    $AdminUPN = $null
}

# =========================================================================
# STEP 12 — EOP/ATP Standard Preset Security Policies
# =========================================================================
if ($AdminUPN) {
    Write-Host "`n=== Step 12: Enable EOP Default Security Policies ===" -ForegroundColor Cyan
    $presetNotInitialized = $false
    function Test-PresetNotFoundError { param($ErrorRecord) return $ErrorRecord.Exception.Message -like "*couldn't be found*" }

    try {
        $eopRule = Get-EOPProtectionPolicyRule -Identity "Standard Preset Security Policy" -ErrorAction Stop
        if ($eopRule.Enabled -eq $false) {
            Enable-EOPProtectionPolicyRule -Identity "Standard Preset Security Policy" -ErrorAction Stop
            Write-Log "[OK] Step 12: EOP Standard Preset Policy enabled." "OK"
        } else {
            Write-Log "[SKIP] Step 12: EOP Standard Preset Policy already enabled." "SKIP"
        }
    } catch {
        if (Test-PresetNotFoundError $_) { $presetNotInitialized = $true; Write-Log "[WARN] Step 12: EOP preset rule never initialized in Defender portal." "WARN" }
        else { Write-Log "[ERROR] Step 12 (EOP): $($_.Exception.Message)" "ERROR" }
    }

    try {
        $atpRule = Get-ATPProtectionPolicyRule -Identity "Standard Preset Security Policy" -ErrorAction Stop
        if ($atpRule.Enabled -eq $false) {
            Enable-ATPProtectionPolicyRule -Identity "Standard Preset Security Policy" -ErrorAction Stop
            Write-Log "[OK] Step 12: ATP Standard Preset Policy enabled." "OK"
        } else {
            Write-Log "[SKIP] Step 12: ATP Standard Preset Policy already enabled." "SKIP"
        }
    } catch {
        if (Test-PresetNotFoundError $_) { $presetNotInitialized = $true; Write-Log "[WARN] Step 12: ATP preset rule never initialized in Defender portal." "WARN" }
        else { Write-Log "[ERROR] Step 12 (ATP): $($_.Exception.Message)" "ERROR" }
    }

    if ($presetNotInitialized) {
        Write-Log "[ACTION REQUIRED] Step 12: Standard preset security policy must be turned on once via security.microsoft.com > Email & collaboration > Policies & rules > Threat policies > Preset Security Policies > Standard protection > Manage." "WARN"
        Add-Summary 12 "Enable_EOP_Default_Security_Policies" "Action required: initialize preset in Defender portal"
    } else {
        Add-Summary 12 "Enable_EOP_Default_Security_Policies" "Success"
    }
}

# =========================================================================
# STEP 13 — Exchange Online Org Settings
# =========================================================================
if ($AdminUPN) {
    Write-Host "`n=== Step 13: Configure Exchange Online Settings ===" -ForegroundColor Cyan
    $auditLog = Join-Path $script:BaselineLogDir "OrgConfigAuditLog_${tenantName}_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    Write-Log "[INFO] Step 13 audit log: $auditLog" "INFO"
    Add-Content $auditLog "=== Script Run: $(Get-Date -Format 's') by $AdminUPN (combined baseline script) ==="

    try {
        $roleMgmtAssignments = Get-ManagementRoleAssignment -RoleAssignee $AdminUPN -Role "Role Management" -ErrorAction SilentlyContinue
        $hasUnscopedRoleMgmt = $roleMgmtAssignments | Where-Object { $_.RecipientWriteScope -in @("Organization", "None") }
        if (-not $hasUnscopedRoleMgmt) {
            Write-Log "[WARN] Step 13: $AdminUPN lacks unscoped Role Management permission. Continuing — the role assignment policy step will be skipped automatically if it fails." "WARN"
            Add-Content $auditLog "PRE-FLIGHT WARNING: $AdminUPN lacks unscoped Role Management permission."
        }

        # Block automatic forwarding to external domains (idempotent - safe to always set)
        Set-RemoteDomain Default -AutoForwardEnabled $false
        Add-Content $auditLog "Blocked automatic forwarding to external domains."

        # Block SMTP forwarding rule
        $existingRule = Get-TransportRule | Where-Object { $_.Name -eq "Block External SMTP Forwarding" }
        if ($existingRule) {
            Write-Log "[SKIP] Step 13: Transport rule 'Block External SMTP Forwarding' already exists." "SKIP"
        } else {
            New-TransportRule -Name "Block External SMTP Forwarding" -SentToScope NotInOrganization -FromScope InOrganization `
                -MessageTypeMatches AutoForward -RejectMessageReasonText "Automatic forwarding to external domains is not allowed." -Enabled $true
            Write-Log "[OK] Step 13: Created transport rule 'Block External SMTP Forwarding'." "OK"
        }

        # Role Assignment Policy - Prevent Add-ins
        try {
            Enable-OrganizationCustomization -ErrorAction Stop
            Write-Log "[INFO] Step 13: Organisation customisation enabled. Waiting 60s to propagate..." "INFO"
            Start-Sleep -Seconds 60
        } catch {
            if ($_.Exception.Message -notmatch "already enabled|already customized|AlreadyCustomized") { throw }
        }

        $availableRoles = Get-ManagementRole | Where-Object { $_.Name -like "My*" } | Select-Object -ExpandProperty Name
        $revisedRoles = @('MyTextMessaging','MyDistributionGroups','MyMailSubscriptions','MyBaseOptions','MyVoiceMail','MyProfileInformation','MyContactInformation','MyRetentionPolicies','MyDistributionGroupMembership') |
            Where-Object { $availableRoles -contains $_ }
        $newPolicyName = 'Role Assignment Policy - Prevent Add-ins'
        $policyExists = Get-RoleAssignmentPolicy | Where-Object { $_.Name -eq $newPolicyName }

        $policyCreated = $false
        if (-not $policyExists) {
            try {
                New-RoleAssignmentPolicy -Name $newPolicyName -Roles $revisedRoles -ErrorAction Stop
                Write-Log "[OK] Step 13: Created role assignment policy '$newPolicyName'." "OK"
                $policyCreated = $true
            } catch {
                Write-Log "[ERROR] Step 13: Failed to create '$newPolicyName': $($_.Exception.Message)" "ERROR"
            }
        } else {
            Write-Log "[SKIP] Step 13: Role assignment policy '$newPolicyName' already exists." "SKIP"
            $policyCreated = $true
        }

        if ($policyCreated) {
            Set-RoleAssignmentPolicy -id $newPolicyName -IsDefault -ErrorAction Stop
            Get-Mailbox -ResultSize Unlimited | Set-Mailbox -RoleAssignmentPolicy $newPolicyName
            Write-Log "[OK] Step 13: Set '$newPolicyName' as default and assigned to all mailboxes." "OK"
        }

        # Enable mailbox audit logging (idempotent - safe to always set)
        Get-Mailbox -ResultSize Unlimited | Set-Mailbox -AuditEnabled $true
        Set-OrganizationConfig -AuditDisabled $false
        Write-Log "[OK] Step 13: Mailbox audit logging enabled." "OK"

        Add-Content $auditLog "=== END SCRIPT RUN (combined) ===`n"
        Add-Summary 13 "Configure_Exchange_Online_Settings" "Success"
    } catch {
        Write-Log "[ERROR] Step 13 failed: $($_.Exception.Message)" "ERROR"
        Add-Content $auditLog "ERROR: $($_.Exception.Message)"
        Add-Summary 13 "Configure_Exchange_Online_Settings" "Error: $($_.Exception.Message)"
    }
}

# =========================================================================
# STEP 14 — Entra LAPS
# =========================================================================
Write-Host "`n=== Step 14: Enable LAPS Policy (Entra ID) ===" -ForegroundColor Cyan
$accountname = "Local-Admin"
try {
    $ConfigCheckUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?`$filter=templateReference/TemplateDisplayName%20eq%20%27Local%20admin%20password%20solution%20(Windows%20LAPS)%27"
    $CurrentConfigPolicy = Invoke-MgGraphRequest -Uri $ConfigCheckUri -Method GET -OutputType PSObject -ContentType "application/json" | Select-Object -ExpandProperty Value

    if ($CurrentConfigPolicy) {
        Write-Log "[SKIP] Step 14: Existing LAPS config policy detected — skipping creation." "SKIP"
    } else {
        $ConfigUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies"
        $configbody = @"
{
  "name": "Windows LAPS",
  "description": "Windows LAPS",
  "platforms": "windows10",
  "technologies": "mdm",
  "roleScopeTagIds": ["0"],
  "settings": [
    { "@odata.type": "#microsoft.graph.deviceManagementConfigurationSetting", "settingInstance": { "@odata.type": "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance", "settingDefinitionId": "device_vendor_msft_laps_policies_backupdirectory", "choiceSettingValue": { "@odata.type": "#microsoft.graph.deviceManagementConfigurationChoiceSettingValue", "value": "device_vendor_msft_laps_policies_backupdirectory_1", "children": [ { "@odata.type": "#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance", "settingDefinitionId": "device_vendor_msft_laps_policies_passwordagedays_aad", "simpleSettingValue": { "@odata.type": "#microsoft.graph.deviceManagementConfigurationIntegerSettingValue", "value": 7 } } ], "settingValueTemplateReference": { "settingValueTemplateId": "4d90f03d-e14c-43c4-86da-681da96a2f92" } }, "settingInstanceTemplateReference": { "settingInstanceTemplateId": "a3270f64-e493-499d-8900-90290f61ed8a" } } },
    { "@odata.type": "#microsoft.graph.deviceManagementConfigurationSetting", "settingInstance": { "@odata.type": "#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance", "settingDefinitionId": "device_vendor_msft_laps_policies_administratoraccountname", "simpleSettingValue": { "@odata.type": "#microsoft.graph.deviceManagementConfigurationStringSettingValue", "value": "$accountname", "settingValueTemplateReference": { "settingValueTemplateId": "992c7fce-f9e4-46ab-ac11-e167398859ea" } }, "settingInstanceTemplateReference": { "settingInstanceTemplateId": "d3d7d492-0019-4f56-96f8-1967f7deabeb" } } },
    { "@odata.type": "#microsoft.graph.deviceManagementConfigurationSetting", "settingInstance": { "@odata.type": "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance", "settingDefinitionId": "device_vendor_msft_laps_policies_passwordcomplexity", "choiceSettingValue": { "@odata.type": "#microsoft.graph.deviceManagementConfigurationChoiceSettingValue", "value": "device_vendor_msft_laps_policies_passwordcomplexity_4", "children": [], "settingValueTemplateReference": { "settingValueTemplateId": "aa883ab5-625e-4e3b-b830-a37a4bb8ce01" } }, "settingInstanceTemplateReference": { "settingInstanceTemplateId": "8a7459e8-1d1c-458a-8906-7b27d216de52" } } },
    { "@odata.type": "#microsoft.graph.deviceManagementConfigurationSetting", "settingInstance": { "@odata.type": "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance", "settingDefinitionId": "device_vendor_msft_laps_policies_postauthenticationactions", "choiceSettingValue": { "@odata.type": "#microsoft.graph.deviceManagementConfigurationChoiceSettingValue", "value": "device_vendor_msft_laps_policies_postauthenticationactions_1", "children": [], "settingValueTemplateReference": { "settingValueTemplateId": "68ff4f78-baa8-4b32-bf3d-5ad5566d8142" } }, "settingInstanceTemplateReference": { "settingInstanceTemplateId": "d9282eb1-d187-42ae-b366-7081f32dcfff" } } },
    { "@odata.type": "#microsoft.graph.deviceManagementConfigurationSetting", "settingInstance": { "@odata.type": "#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance", "settingDefinitionId": "device_vendor_msft_laps_policies_postauthenticationresetdelay", "simpleSettingValue": { "@odata.type": "#microsoft.graph.deviceManagementConfigurationIntegerSettingValue", "value": 1, "settingValueTemplateReference": { "settingValueTemplateId": "0deb6aee-8dac-40c4-a9dd-c3718e5c1d52" } }, "settingInstanceTemplateReference": { "settingInstanceTemplateId": "a9e21166-4055-4042-9372-efaf3ef41868" } } }
  ],
  "templateReference": { "templateId": "adc46e5a-f4aa-4ff6-aeff-4f27bc525796_1" }
}
"@
        Invoke-MgGraphRequest -Method POST -Uri $ConfigUri -Body $configbody -ContentType "application/json" -ErrorAction Stop | Out-Null
        Write-Log "[OK] Step 14: Windows LAPS config policy created." "OK"
    }

    # Remediation script — check for existing "Windows LAPS User" health script first
    $existingScripts = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts?`$filter=displayName eq 'Windows LAPS User'"
    if ($existingScripts.value.Count -gt 0) {
        Write-Log "[SKIP] Step 14: Remediation script 'Windows LAPS User' already exists." "SKIP"
        Add-Summary 14 "Enable_LAPS_Policy_EntraID" "Skipped (already applied)"
    } else {
        $Remediationscript = @'
#Add system.web assembly
Add-Type -AssemblyName 'System.Web'
$Userexist = (Get-LocalUser).Name -Contains ">Placeholder<"
if (!$userexist) {
    $password = [System.Web.Security.Membership]::GeneratePassword(20,5)
    $Securepassword = ConvertTo-SecureString $Password -AsPlainText -force
    $params = @{ Name = ">Placeholder<"; Password = $Securepassword }
    New-LocalUser @params
}
Add-LocalGroupMember -Group "Administrators" -Member ">Placeholder<"
'@
        $Remediationscript = $Remediationscript -replace ">Placeholder<", "$accountname"
        $RemediationBase64 = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($Remediationscript))

        $Detectionscript = @'
$Userexist = (Get-LocalUser).Name -Contains ">Placeholder<"
if ($userexist) { Write-Host ">Placeholder< exists" } Else { Write-Host ">Placeholder< does not Exists"; Exit 1 }
$localadmins = ([ADSI]"WinNT://./Administrators").psbase.Invoke('Members') | % { ([ADSI]$_).InvokeGet('AdsPath') }
if ($localadmins -like "*>Placeholder<*") { Write-Host ">Placeholder< is a member of local admins"; exit 0 } else { Write-Host ">Placeholder< is NOT a member of local admins"; exit 1 }
'@
        $Detectionscript = $Detectionscript -replace ">Placeholder<", "$accountname"
        $DetectionBase64 = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($Detectionscript))

        $ScriptURI = "https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts"
        $scriptbody = @"
{
  "displayName": "Windows LAPS User",
  "description": "Checks for account \"$accountname\". If it doesn't exist, it will create it with a random password and add it to the local administrators group.",
  "publisher": "",
  "runAs32Bit": false,
  "runAsAccount": "system",
  "enforceSignatureCheck": false,
  "detectionScriptContent": "$DetectionBase64",
  "remediationScriptContent": "$RemediationBase64",
  "roleScopeTagIds": ["0"]
}
"@
        Invoke-MgGraphRequest -Method POST -Uri $ScriptURI -Body $scriptbody -ContentType "application/json" -ErrorAction Stop | Out-Null
        Write-Log "[OK] Step 14: Remediation package 'Windows LAPS User' created." "OK"
        Add-Summary 14 "Enable_LAPS_Policy_EntraID" "Success"
    }
} catch {
    Write-Log "[ERROR] Step 14 failed: $($_.Exception.Message)" "ERROR"
    Add-Summary 14 "Enable_LAPS_Policy_EntraID" "Error: $($_.Exception.Message)"
}

# =========================================================================
# STEP 15 — Verify / Disable Conditional Access Policies
# =========================================================================
Write-Host "`n=== Step 15: Verify Conditional Access Policies Disabled ===" -ForegroundColor Cyan
try {
    $expectedPolicyNames = @(
        "Block Access Outside Approved Countries"
        "Block Legacy Authentication"
        "Require MFA for Admin Portals - 8hr"
        "Require MFA for Admin Roles - 8hr"
        "MFA for All Users - Browser Only - 8hr"
        "Require MFA for All Users"
        "Require MFA for Guest Users - 8hr"
    )
    $allPolicies = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies").value
    $missing = 0; $fixed = 0; $ok = 0

    foreach ($name in $expectedPolicyNames) {
        $policy = $allPolicies | Where-Object { $_.displayName -eq $name }
        if (-not $policy) {
            Write-Log "[WARN] Step 15: Missing policy '$name'." "WARN"
            $missing++
            continue
        }
        if ($policy.state -ne "disabled") {
            Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies/$($policy.id)" -Body '{"state":"disabled"}' -ContentType "application/json" -ErrorAction Stop
            Write-Log "[OK] Step 15: Fixed '$name' — set to disabled." "OK"
            $fixed++
        } else {
            Write-Log "[SKIP] Step 15: '$name' already disabled." "SKIP"
            $ok++
        }
    }

    if ($missing -gt 0) {
        Add-Summary 15 "Verify_Disable_CA_Policies" "Completed with $missing missing polic(y/ies)"
    } elseif ($fixed -gt 0) {
        Add-Summary 15 "Verify_Disable_CA_Policies" "Success ($fixed corrected)"
    } else {
        Add-Summary 15 "Verify_Disable_CA_Policies" "Skipped (all already disabled)"
    }
} catch {
    Write-Log "[ERROR] Step 15 failed: $($_.Exception.Message)" "ERROR"
    Add-Summary 15 "Verify_Disable_CA_Policies" "Error: $($_.Exception.Message)"
}

# =========================================================================
# Finish
# =========================================================================
$elapsed = (Get-Date) - $startTime
$global:logContent -join "`n" | Out-File -FilePath $LogPath -Encoding utf8

Write-Host ""
Write-Host "=================================================" -ForegroundColor Green
Write-Host "  Run complete in $([math]::Round($elapsed.TotalMinutes, 1)) minutes" -ForegroundColor Green
Write-Host "  Log: $LogPath" -ForegroundColor Green
Write-Host "=================================================" -ForegroundColor Green
Write-Host ""
Write-Host "========== SUMMARY ==========" -ForegroundColor Cyan
$global:summary | Sort-Object Num | Format-Table -AutoSize
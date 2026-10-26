#Requires -Version 7.0
<#
.SYNOPSIS
    Exports Microsoft 365 user license assignments with friendly names.

.DESCRIPTION
    Connects to Microsoft Graph and exports user license assignments to CSV,
    including friendly license names (e.g., "Microsoft 365 E3" instead of "SPE_E3").

    Creates two CSV files:
    - UserLicenses_<timestamp>.csv - One row per user with all assigned licenses
    - LicenseCounts_<timestamp>.csv - Aggregate count per license SKU

.PARAMETER OutputPath
    Directory to save the CSV files. Defaults to current directory.

.PARAMETER Domain
    Optional domain filter (e.g., "contoso.com"). Only exports users with UPN ending in this domain.

.PARAMETER SearchTerm
    Optional search term to filter users by DisplayName.

.EXAMPLE
    .\Export-UserLicenses.ps1

    Exports all users' licenses to the current directory.

.EXAMPLE
    .\Export-UserLicenses.ps1 -Domain "contoso.com" -OutputPath "C:\Reports"

    Exports licenses for users with @contoso.com UPNs to C:\Reports.

.EXAMPLE
    .\Export-UserLicenses.ps1 -SearchTerm "Contoso"

    Exports licenses for users whose DisplayName contains "Contoso".
#>

[CmdletBinding()]
param(
    [string]$OutputPath = (Get-Location).Path,
    [string]$Domain,
    [string]$SearchTerm
)

# -----------------------------------------------------------------------
# License Friendly Names Mapping
# -----------------------------------------------------------------------
$script:LicenseFriendlyNames = @{
    # Microsoft 365 Enterprise
    'SPE_E3'                        = 'Microsoft 365 E3'
    'SPE_E5'                        = 'Microsoft 365 E5'
    'SPE_F1'                        = 'Microsoft 365 F1'
    'SPE_F3'                        = 'Microsoft 365 F3'
    'Microsoft_365_E3'              = 'Microsoft 365 E3'
    'Microsoft_365_E5'              = 'Microsoft 365 E5'
    'M365EDU_A3_FACULTY'            = 'Microsoft 365 A3 for Faculty'
    'M365EDU_A3_STUDENT'            = 'Microsoft 365 A3 for Students'
    'M365EDU_A5_FACULTY'            = 'Microsoft 365 A5 for Faculty'
    'M365EDU_A5_STUDENT'            = 'Microsoft 365 A5 for Students'

    # Office 365 Enterprise
    'ENTERPRISEPACK'                = 'Office 365 E3'
    'ENTERPRISEPREMIUM'             = 'Office 365 E5'
    'ENTERPRISEPACK_USGOV_DOD'      = 'Office 365 E3 (US Gov DoD)'
    'ENTERPRISEPACK_USGOV_GCCHIGH'  = 'Office 365 E3 (US Gov GCC High)'
    'ENTERPRISEPREMIUM_NOPSTNCONF'  = 'Office 365 E5 without Audio Conferencing'
    'DESKLESSPACK'                  = 'Office 365 F3'
    'O365_BUSINESS'                 = 'Office 365 Business'
    'O365_BUSINESS_ESSENTIALS'      = 'Office 365 Business Essentials'
    'O365_BUSINESS_PREMIUM'         = 'Office 365 Business Premium'
    'SMB_BUSINESS'                  = 'Microsoft 365 Apps for Business'
    'OFFICESUBSCRIPTION'            = 'Microsoft 365 Apps for Enterprise'
    'LITEPACK'                      = 'Office 365 Small Business'

    # Microsoft 365 Business
    'SPB'                           = 'Microsoft 365 Business Premium'
    'O365_BUSINESS_PREMIUM'         = 'Microsoft 365 Business Standard'
    'SMB_BUSINESS_PREMIUM'          = 'Microsoft 365 Business Premium'
    'BUSINESS_ESSENTIALS'           = 'Microsoft 365 Business Basic'
    'SMB_BUSINESS_ESSENTIALS'       = 'Microsoft 365 Business Basic'

    # Exchange Online
    'EXCHANGESTANDARD'              = 'Exchange Online Plan 1'
    'EXCHANGEENTERPRISE'            = 'Exchange Online Plan 2'
    'EXCHANGEARCHIVE'               = 'Exchange Online Archiving'
    'EXCHANGEARCHIVE_ADDON'         = 'Exchange Online Archiving for Exchange Online'
    'EXCHANGEDESKLESS'              = 'Exchange Online Kiosk'
    'EXCHANGE_S_ESSENTIALS'         = 'Exchange Online Essentials'
    'EOP_ENTERPRISE'                = 'Exchange Online Protection'

    # SharePoint
    'SHAREPOINTSTANDARD'            = 'SharePoint Online Plan 1'
    'SHAREPOINTENTERPRISE'          = 'SharePoint Online Plan 2'
    'SHAREPOINTSTORAGE'             = 'SharePoint Online Storage'
    'SHAREPOINTWAC'                 = 'Office Online'

    # Teams
    'TEAMS_EXPLORATORY'             = 'Microsoft Teams Exploratory'
    'TEAMS1'                        = 'Microsoft Teams'
    'TEAMS_COMMERCIAL_TRIAL'        = 'Microsoft Teams Commercial Cloud Trial'
    'MCOEV'                         = 'Microsoft Teams Phone Standard'
    'MCOEV_DOD'                     = 'Microsoft Teams Phone Standard for DoD'
    'MCOEV_GOV'                     = 'Microsoft Teams Phone Standard for Government'
    'MCOPSTN1'                      = 'Domestic Calling Plan (120 min)'
    'MCOPSTN2'                      = 'Domestic and International Calling Plan'
    'MCOPSTN5'                      = 'Domestic Calling Plan (3000 min)'
    'MCOPSTN_5'                     = 'Domestic Calling Plan (3000 min)'
    'MCOPSTN6'                      = 'Domestic Calling Plan (1200 min)'
    'MCOMEETADV'                    = 'Microsoft 365 Audio Conferencing'

    # Power Platform
    'FLOW_FREE'                     = 'Microsoft Power Automate Free'
    'FLOW_P1'                       = 'Microsoft Power Automate Plan 1'
    'FLOW_P2'                       = 'Microsoft Power Automate Plan 2'
    'POWER_BI_STANDARD'             = 'Power BI Free'
    'POWER_BI_PRO'                  = 'Power BI Pro'
    'POWER_BI_ADDON'                = 'Power BI for Office 365 Add-On'
    'POWER_BI_INDIVIDUAL_USER'      = 'Power BI Individual'
    'POWERAPPS_VIRAL'               = 'Microsoft Power Apps Plan 2 Trial'
    'POWERAPPS_PER_USER'            = 'Power Apps Per User Plan'
    'POWERAPPS_PER_APP'             = 'Power Apps Per App Plan'

    # Project & Visio
    'PROJECTPREMIUM'                = 'Project Plan 5'
    'PROJECTONLINE_PLAN_1'          = 'Project Plan 1'
    'PROJECTESSENTIALS'             = 'Project Plan 3'
    'PROJECTPROFESSIONAL'           = 'Project Plan 5'
    'VISIOONLINE_PLAN1'             = 'Visio Plan 1'
    'VISIOCLIENT'                   = 'Visio Plan 2'

    # Security & Compliance
    'EMS'                           = 'Enterprise Mobility + Security E3'
    'EMSPREMIUM'                    = 'Enterprise Mobility + Security E5'
    'AAD_PREMIUM'                   = 'Azure Active Directory Premium P1'
    'AAD_PREMIUM_P2'                = 'Azure Active Directory Premium P2'
    'ADALLOM_S_STANDALONE'          = 'Microsoft Defender for Cloud Apps'
    'ATP_ENTERPRISE'                = 'Microsoft Defender for Office 365 Plan 1'
    'THREAT_INTELLIGENCE'           = 'Microsoft Defender for Office 365 Plan 2'
    'INFORMATION_PROTECTION_COMPLIANCE' = 'Microsoft 365 E5 Compliance'
    'IDENTITY_THREAT_PROTECTION'    = 'Microsoft 365 E5 Security'
    'M365_SECURITY_COMPLIANCE_FOR_FLW' = 'Microsoft 365 F5 Security + Compliance'
    'RIGHTSMANAGEMENT'              = 'Azure Rights Management'
    'RIGHTSMANAGEMENT_ADHOC'        = 'Rights Management Adhoc'
    'INFORMATION_BARRIERS'          = 'Information Barriers'
    'EQUIVIO_ANALYTICS'             = 'Office 365 Advanced eDiscovery'
    'LOCKBOX_ENTERPRISE'            = 'Customer Lockbox'

    # Intune & Device Management
    'INTUNE_A'                      = 'Microsoft Intune'
    'INTUNE_A_D'                    = 'Microsoft Intune Device'
    'INTUNE_A_VL'                   = 'Microsoft Intune (Volume License)'
    'WIN10_PRO_ENT_SUB'             = 'Windows 10/11 Enterprise E3'
    'WIN10_VDA_E3'                  = 'Windows 10/11 Enterprise E3 VDA'
    'WIN10_VDA_E5'                  = 'Windows 10/11 Enterprise E5'

    # Standalone Services
    'STREAM'                        = 'Microsoft Stream'
    'COMMUNICATIONS_DLP'            = 'Microsoft Communications DLP'
    'COMMUNICATIONS_COMPLIANCE'     = 'Microsoft Communications Compliance'
    'YAMMER_ENTERPRISE'             = 'Yammer Enterprise'
    'RMS_S_ENTERPRISE'              = 'Azure Active Directory Rights Management'
    'CRMSTANDARD'                   = 'Microsoft Dynamics CRM Online'
    'CRMPLAN2'                      = 'Microsoft Dynamics CRM Online Basic'

    # Developer & Other
    'DEVELOPERPACK'                 = 'Microsoft 365 E3 Developer'
    'DEVELOPERPACK_E5'              = 'Microsoft 365 E5 Developer'
    'WIN_DEF_ATP'                   = 'Microsoft Defender for Endpoint'
    'MDATP_Server'                  = 'Microsoft Defender for Endpoint Server'
    'WDATP_UNIFIED_SERVER'          = 'Microsoft Defender for Endpoint Server'
}

function Get-FriendlyLicenseName {
    param([string]$SkuPartNumber)
    if ($script:LicenseFriendlyNames.ContainsKey($SkuPartNumber)) {
        return $script:LicenseFriendlyNames[$SkuPartNumber]
    }
    return $SkuPartNumber
}

# -----------------------------------------------------------------------
# Main Script
# -----------------------------------------------------------------------

Write-Host ''
Write-Host 'Microsoft 365 License Export' -ForegroundColor Cyan
Write-Host ('=' * 50) -ForegroundColor Cyan
Write-Host ''

# Check prerequisites
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    Write-Host '[ERROR] Microsoft.Graph.Authentication module not found.' -ForegroundColor Red
    Write-Host '        Install it with: Install-Module Microsoft.Graph -Scope CurrentUser' -ForegroundColor Yellow
    exit 1
}

# Import required modules
Write-Host '[INFO] Importing Microsoft.Graph.Authentication...' -ForegroundColor Gray
Import-Module Microsoft.Graph.Authentication -DisableNameChecking -ErrorAction Stop

# Connect to Microsoft Graph
$scopes = @('User.Read.All', 'Organization.Read.All')
Write-Host '[INFO] Connecting to Microsoft Graph...' -ForegroundColor Gray
Write-Host "        Scopes: $($scopes -join ', ')" -ForegroundColor DarkGray

try {
    Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop
    Write-Host '[OK]   Connected to Microsoft Graph' -ForegroundColor Green
}
catch {
    Write-Host "[ERROR] Failed to connect to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

try {
    # Fetch tenant subscribed SKUs
    Write-Host ''
    Write-Host '[INFO] Fetching tenant license SKUs...' -ForegroundColor Gray
    $skuResponse = Invoke-MgGraphRequest -Uri 'https://graph.microsoft.com/v1.0/subscribedSkus' -Method GET -ErrorAction Stop
    $subscribedSkus = $skuResponse.value

    $skuLookup = @{}
    foreach ($sku in $subscribedSkus) {
        if ($sku.skuId) {
            $skuLookup[$sku.skuId.ToString()] = $sku.skuPartNumber
        }
    }
    Write-Host "[OK]   Found $($subscribedSkus.Count) tenant SKU(s)" -ForegroundColor Green

    # Build user filter
    Write-Host ''
    Write-Host '[INFO] Fetching users...' -ForegroundColor Gray
    $filterParts = @("userType eq 'Member'")

    if ($Domain) {
        $filterParts += "endsWith(userPrincipalName, '@$Domain')"
        Write-Host "        Domain filter: @$Domain" -ForegroundColor DarkGray
    }

    if ($SearchTerm) {
        $filterParts += "startsWith(displayName, '$SearchTerm')"
        Write-Host "        Name filter: $SearchTerm" -ForegroundColor DarkGray
    }

    $filter = $filterParts -join ' and '
    $uri = "https://graph.microsoft.com/v1.0/users?`$filter=$filter&`$select=id,userPrincipalName,displayName&`$top=999"

    $users = [System.Collections.Generic.List[object]]::new()
    do {
        $response = Invoke-MgGraphRequest -Uri $uri -Method GET -ErrorAction Stop
        foreach ($u in $response.value) {
            $users.Add($u)
        }
        $uri = $response.'@odata.nextLink'
    } while ($uri)

    Write-Host "[OK]   Found $($users.Count) user(s)" -ForegroundColor Green

    if ($users.Count -eq 0) {
        Write-Host '[WARN] No users found matching the criteria' -ForegroundColor Yellow
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        exit 0
    }

    # Fetch licenses for each user
    Write-Host ''
    Write-Host '[INFO] Fetching license assignments...' -ForegroundColor Gray

    $userLicenses = [System.Collections.Generic.List[PSCustomObject]]::new()
    $licenseCount = @{}
    $licenseFriendlyByPart = @{}
    $current = 0

    foreach ($user in $users) {
        $current++
        Write-Progress -Activity 'Fetching license details' -Status "User $current of $($users.Count): $($user.userPrincipalName)" -PercentComplete (($current / $users.Count) * 100)

        try {
            $upnEnc = [Uri]::EscapeDataString($user.userPrincipalName)
            $licResponse = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/users/$upnEnc/licenseDetails" -Method GET -ErrorAction Stop
            $licenses = $licResponse.value

            $skuParts = [System.Collections.Generic.List[string]]::new()
            $skuFriendly = [System.Collections.Generic.List[string]]::new()
            $skuIds = [System.Collections.Generic.List[string]]::new()

            foreach ($lic in $licenses) {
                $skuId = $lic.skuId.ToString()
                $skuPart = $lic.skuPartNumber

                if (-not $skuPart -and $skuLookup.ContainsKey($skuId)) {
                    $skuPart = $skuLookup[$skuId]
                }

                if ($skuPart) {
                    $friendly = Get-FriendlyLicenseName -SkuPartNumber $skuPart

                    $skuParts.Add($skuPart)
                    $skuFriendly.Add($friendly)
                    $skuIds.Add($skuId)

                    if ($licenseCount.ContainsKey($skuPart)) {
                        $licenseCount[$skuPart]++
                    } else {
                        $licenseCount[$skuPart] = 1
                    }

                    if (-not $licenseFriendlyByPart.ContainsKey($skuPart)) {
                        $licenseFriendlyByPart[$skuPart] = $friendly
                    }
                }
            }

            $userLicenses.Add([PSCustomObject]@{
                DisplayName       = $user.displayName
                UserPrincipalName = $user.userPrincipalName
                LicenseCount      = $skuParts.Count
                SkuPartNumbers    = ($skuParts -join '; ')
                LicenseNames      = ($skuFriendly -join '; ')
                SkuIds            = ($skuIds -join '; ')
            })
        }
        catch {
            Write-Host "[WARN] Could not fetch licenses for $($user.userPrincipalName): $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    Write-Progress -Activity 'Fetching license details' -Completed

    # Build aggregate counts
    $licenseCounts = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($skuPart in ($licenseCount.Keys | Sort-Object)) {
        $friendly = if ($licenseFriendlyByPart.ContainsKey($skuPart)) {
            $licenseFriendlyByPart[$skuPart]
        } else {
            Get-FriendlyLicenseName -SkuPartNumber $skuPart
        }

        $licenseCounts.Add([PSCustomObject]@{
            SkuPartNumber = $skuPart
            LicenseName   = $friendly
            AssignedCount = $licenseCount[$skuPart]
        })
    }

    # Sort by count descending
    $licenseCounts = $licenseCounts | Sort-Object AssignedCount -Descending

    # Export to CSV
    Write-Host ''
    Write-Host '[INFO] Writing CSV files...' -ForegroundColor Gray

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $userLicensesPath = Join-Path $OutputPath "UserLicenses_$timestamp.csv"
    $licenseCountsPath = Join-Path $OutputPath "LicenseCounts_$timestamp.csv"

    $userLicenses | Export-Csv -Path $userLicensesPath -NoTypeInformation -Encoding UTF8
    $licenseCounts | Export-Csv -Path $licenseCountsPath -NoTypeInformation -Encoding UTF8

    Write-Host ''
    Write-Host '[OK]   Export complete!' -ForegroundColor Green
    Write-Host "       User Licenses: $userLicensesPath" -ForegroundColor Gray
    Write-Host "       License Counts: $licenseCountsPath" -ForegroundColor Gray
    Write-Host ''
    Write-Host "       Users exported: $($userLicenses.Count)" -ForegroundColor Cyan
    Write-Host "       Unique licenses: $($licenseCounts.Count)" -ForegroundColor Cyan
    Write-Host "       Users with licenses: $(($userLicenses | Where-Object { $_.LicenseCount -gt 0 }).Count)" -ForegroundColor Cyan
    Write-Host ''
}
catch {
    Write-Host ''
    Write-Host "[ERROR] Export failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "        $($_.ScriptStackTrace)" -ForegroundColor DarkGray
    exit 1
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}

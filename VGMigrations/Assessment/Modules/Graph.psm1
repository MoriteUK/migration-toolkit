#Requires -Version 7.0

Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -DisableNameChecking -Force -Global

# Graph API base
$script:GraphBase = 'https://graph.microsoft.com/v1.0'

# Static skuPartNumber -> friendly product name lookup, falls back to the raw skuPartNumber
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

# -----------------------------------------------------------------------
# Private functions
# -----------------------------------------------------------------------

# Paginates all unified groups tenant-wide. Does NOT use $expand (BadRequest when combined with $filter).
# resourceProvisioningOptions is included so callers can identify Teams.
<#
.SYNOPSIS
    Paginates all unified groups tenant-wide via Invoke-MgGraphRequest.
#>
function Get-UnifiedGroupsWithMembers {
    $select = 'id,displayName,description,mail,proxyAddresses,mailNickname,visibility,createdDateTime,resourceProvisioningOptions'
    # groupTypes/any(...) is an ADVANCED query: without ConsistencyLevel: eventual + $count=true,
    # Graph silently returns an incomplete result set with a prematurely-terminating nextLink.
    $uri     = "$script:GraphBase/groups?`$filter=groupTypes/any(c:c+eq+'Unified')&`$select=$select&`$count=true&`$top=999"
    $headers = @{ 'ConsistencyLevel' = 'eventual' }
    $groups  = [System.Collections.Generic.List[object]]::new()

    do {
        # Header must ride along on every page - nextLink requests need it too
        $response = Invoke-MgGraphRequest -Uri $uri -Method GET -Headers $headers -ErrorAction Stop
        foreach ($g in $response.value) { $groups.Add($g) }
        $uri = $response.'@odata.nextLink'
    } while ($uri)

    return $groups.ToArray()
}

<#
.SYNOPSIS
    Fetches owner or member UPNs for a group, returning an empty array on error.
#>
function Resolve-GroupMemberUpns {
    param(
        [string]$GroupId,
        [ValidateSet('owners','members')][string]$RelType
    )
    try {
        $uri    = "$script:GraphBase/groups/$GroupId/$RelType`?`$select=userPrincipalName&`$top=999"
        $result = Invoke-MgGraphRequest -Uri $uri -Method GET -ErrorAction Stop
        return @($result.value | ForEach-Object { $_.userPrincipalName } | Where-Object { $_ })
    }
    catch { return @() }
}

<#
.SYNOPSIS
    Gets the SharePoint root site URL for a group, or null if unavailable.
#>
function Resolve-GroupSiteUrl {
    param([string]$GroupId)
    try {
        $result = Invoke-MgGraphRequest -Uri "$script:GraphBase/groups/$GroupId/sites/root?`$select=webUrl" -Method GET -ErrorAction Stop
        return $result.webUrl
    }
    catch { return $null }
}

<#
.SYNOPSIS
    Builds VBU-scoped Team records with owners, members, channel counts, and site URL from the pre-fetched unified groups.
#>
function Get-TeamData {
    param(
        [PSCustomObject]$Context,
        [object[]]$AllUnifiedGroups
    )

    $searchTerm = $Context.VBUSearchTerm
    $domain     = $Context.VBUDomain
    $records    = [System.Collections.Generic.List[PSCustomObject]]::new()

    $teamGroups = $AllUnifiedGroups | Where-Object {
        $_.resourceProvisioningOptions -contains 'Team'
    }

    $scoped = $teamGroups | Where-Object {
        (Test-WholeWordMatch -Text $_.displayName -Term $searchTerm) -or
        ($_.mail        -like "*$domain*")     -or
        ($_.proxyAddresses | Where-Object { $_ -like "*$domain*" })
    }

    foreach ($g in $scoped) {
        $ownerUpns  = Resolve-GroupMemberUpns -GroupId $g.id -RelType 'owners'
        $memberUpns = Resolve-GroupMemberUpns -GroupId $g.id -RelType 'members'
        $siteUrl    = Resolve-GroupSiteUrl   -GroupId $g.id

        $channels           = @()
        $channelCount       = 0
        $privateChannelCount = 0
        $sharedChannelCount  = 0

        try {
            $chanResult          = Invoke-MgGraphRequest -Uri "$script:GraphBase/teams/$($g.id)/channels?`$select=id,membershipType" -Method GET -ErrorAction Stop
            $channels            = $chanResult.value
            $channelCount        = $channels.Count
            $privateChannelCount = ($channels | Where-Object { $_.membershipType -eq 'private' }).Count
            $sharedChannelCount  = ($channels | Where-Object { $_.membershipType -eq 'shared' }).Count
        }
        catch { }

        $records.Add([PSCustomObject]@{
            TeamId               = $g.id
            GroupId              = $g.id
            DisplayName          = $g.displayName
            Description          = $g.description
            Visibility           = $g.visibility
            TeamType             = 'Standard'
            CreatedDateTime      = $g.createdDateTime
            Owners               = $ownerUpns  -join '|'
            Members              = $memberUpns -join '|'
            OwnerCount           = $ownerUpns.Count
            MemberCount          = $memberUpns.Count
            ChannelCount         = $channelCount
            PrivateChannelCount  = $privateChannelCount
            SharedChannelCount   = $sharedChannelCount
            SharePointSiteUrl    = $siteUrl
            AssociatedObject     = $null
            MigrationObjectType  = $null
        })
    }

    return $records.ToArray()
}

<#
.SYNOPSIS
    Builds VBU-scoped M365 Group records with owners and members from the pre-fetched unified groups.
#>
function Get-M365GroupData {
    param(
        [PSCustomObject]$Context,
        [object[]]$AllUnifiedGroups
    )

    $searchTerm = $Context.VBUSearchTerm
    $domain     = $Context.VBUDomain
    $records    = [System.Collections.Generic.List[PSCustomObject]]::new()

    $scoped = $AllUnifiedGroups | Where-Object {
        (Test-WholeWordMatch -Text $_.displayName -Term $searchTerm) -or
        ($_.mail        -like "*$domain*")     -or
        ($_.proxyAddresses | Where-Object { $_ -like "*$domain*" })
    }

    foreach ($g in $scoped) {
        $ownerUpns  = Resolve-GroupMemberUpns -GroupId $g.id -RelType 'owners'
        $memberUpns = Resolve-GroupMemberUpns -GroupId $g.id -RelType 'members'

        $records.Add([PSCustomObject]@{
            GroupId             = $g.id
            DisplayName         = $g.displayName
            MailNickname        = $g.mailNickname
            PrimarySmtpAddress  = $g.mail
            Visibility          = $g.visibility
            CreatedDateTime     = $g.createdDateTime
            Owners              = $ownerUpns  -join '|'
            Members             = $memberUpns -join '|'
            OwnerCount          = $ownerUpns.Count
            MemberCount         = $memberUpns.Count
            SharePointSiteUrl   = $null
            AssociatedObject    = $null
            MigrationObjectType = $null
        })
    }

    return $records.ToArray()
}

<#
.SYNOPSIS
    Retrieves Entra registered devices per in-scope user via /users/{upn}/registeredDevices, cross-referenced against ADUsers.json.
#>
function Get-IntuneDeviceData {
    param([PSCustomObject]$Context)

    $adUsersPath = Join-Path $Context.RawPath 'ADUsers.json'
    if (-not (Test-Path $adUsersPath)) {
        Write-Host ($PREFIX_WARN + 'ADUsers.json not found - Intune device cross-reference skipped') -ForegroundColor Yellow
        return @()
    }

    $adUsers = Get-Content $adUsersPath -Raw | ConvertFrom-Json
    $records = [System.Collections.Generic.List[PSCustomObject]]::new()
    $seen    = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

    $select = 'id,deviceId,displayName,operatingSystem,operatingSystemVersion,trustType,approximateLastSignInDateTime'

    foreach ($u in $adUsers) {
        $upn = $u.UserPrincipalName
        if (-not $upn) { continue }

        $uri = "$script:GraphBase/users/$([Uri]::EscapeDataString($upn))/registeredDevices?`$select=$select"
        try {
            do {
                $response = Invoke-MgGraphRequest -Uri $uri -Method GET -ErrorAction Stop
                foreach ($d in $response.value) {
                    $key = if ($d.deviceId) { $d.deviceId } else { $d.id }
                    if ($seen.Contains($key)) { continue }
                    [void]$seen.Add($key)
                    $records.Add([PSCustomObject]@{
                        DeviceName                    = $d.displayName
                        DeviceId                      = if ($d.deviceId) { $d.deviceId } else { $d.id }
                        OperatingSystem               = $d.operatingSystem
                        OperatingSystemVersion        = $d.operatingSystemVersion
                        TrustType                     = $d.trustType
                        JoinType                      = $d.trustType
                        Enabled                       = $true
                        IsManaged                     = $true
                        IsCompliant                   = $null
                        ApproximateLastSignInDateTime = $d.approximateLastSignInDateTime
                        PrimaryUserUPN                = $upn
                        RegisteredOwners              = $upn
                        RegisteredUsers               = $upn
                        Source                        = 'Intune'
                    })
                }
                $uri = $response.'@odata.nextLink'
            } while ($uri)
        }
        catch {
            Write-Host ($PREFIX_WARN + "Could not retrieve devices for '$upn': " + $_.Exception.Message) -ForegroundColor Yellow
        }
    }

    return $records.ToArray()
}

<#
.SYNOPSIS
    Paginates all app registrations and filters to VBU scope on display name, identifier URIs, and redirect URIs, fetching owners per matched app.
#>
function Get-AppRegistrationData {
    param([PSCustomObject]$Context)

    $t      = $Context.VBUSearchTerm
    $select = 'id,appId,displayName,identifierUris,web,publicClient,spa,createdDateTime'
    $uri    = "$script:GraphBase/applications?`$select=$select&`$top=999"
    $apps   = [System.Collections.Generic.List[object]]::new()

    do {
        $response = Invoke-MgGraphRequest -Uri $uri -Method GET -ErrorAction Stop
        foreach ($a in $response.value) { $apps.Add($a) }
        $uri = $response.'@odata.nextLink'
    } while ($uri)

    $records = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($app in $apps) {
        $identifierUris = @($app.identifierUris | Where-Object { $_ })

        $replyUrls = [System.Collections.Generic.List[string]]::new()
        foreach ($platform in @($app.web, $app.publicClient, $app.spa)) {
            if ($platform -and $platform.redirectUris) {
                foreach ($url in $platform.redirectUris) { $replyUrls.Add($url) }
            }
        }

        $matched = (Test-WholeWordMatch -Text $app.displayName -Term $t) -or
                   ($identifierUris | Where-Object { $_ -like "*$t*" }) -or
                   ($replyUrls      | Where-Object { $_ -like "*$t*" })
        if (-not $matched) { continue }

        $owners = @()
        try {
            $ownerResult = Invoke-MgGraphRequest -Uri "$script:GraphBase/applications/$($app.id)/owners?`$select=userPrincipalName" -Method GET -ErrorAction Stop
            $owners = @($ownerResult.value | ForEach-Object { $_.userPrincipalName } | Where-Object { $_ })
        }
        catch { }

        $records.Add([PSCustomObject]@{
            DisplayName     = $app.displayName
            AppId           = $app.appId
            ObjectId        = $app.id
            IdentifierUris  = ($identifierUris | Sort-Object) -join '|'
            ReplyUrls       = ($replyUrls | Sort-Object) -join '|'
            Owners          = $owners -join '|'
            CreatedDateTime = $app.createdDateTime
        })
    }

    return $records.ToArray()
}

<#
.SYNOPSIS
    Paginates all enterprise apps (service principals) and filters to VBU scope on display name, homepage, and reply URLs.
#>
function Get-EnterpriseAppData {
    param([PSCustomObject]$Context)

    $t      = $Context.VBUSearchTerm
    $select = 'id,appId,displayName,homepage,replyUrls,servicePrincipalType,accountEnabled,tags'
    $uri    = "$script:GraphBase/servicePrincipals?`$select=$select&`$top=999"
    $sps    = [System.Collections.Generic.List[object]]::new()

    do {
        $response = Invoke-MgGraphRequest -Uri $uri -Method GET -ErrorAction Stop
        foreach ($sp in $response.value) { $sps.Add($sp) }
        $uri = $response.'@odata.nextLink'
    } while ($uri)

    $records = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($sp in $sps) {
        $replyUrls = @($sp.replyUrls | Where-Object { $_ })

        $matched = (Test-WholeWordMatch -Text $sp.displayName -Term $t) -or
                   ($sp.homepage     -like "*$t*") -or
                   ($replyUrls | Where-Object { $_ -like "*$t*" })
        if (-not $matched) { continue }

        $records.Add([PSCustomObject]@{
            DisplayName          = $sp.displayName
            AppId                = $sp.appId
            ObjectId             = $sp.id
            Homepage             = $sp.homepage
            ReplyUrls            = ($replyUrls | Sort-Object) -join '|'
            ServicePrincipalType = $sp.servicePrincipalType
            AccountEnabled       = $sp.accountEnabled
        })
    }

    return $records.ToArray()
}

<#
.SYNOPSIS
    Resolves a skuPartNumber to a friendly product name, falling back to the raw value if unknown.
#>
function Get-FriendlyLicenseName {
    param([string]$SkuPartNumber)
    if ($script:LicenseFriendlyNames.ContainsKey($SkuPartNumber)) { return $script:LicenseFriendlyNames[$SkuPartNumber] }
    return $SkuPartNumber
}

<#
.SYNOPSIS
    Retrieves tenant-wide subscribed SKUs (licenses available on the tenant).
#>
function Get-SubscribedSkuData {
    try {
        Write-Host ($PREFIX_INFO + 'Fetching tenant subscribed SKUs...') -ForegroundColor DarkGray
        $response = Invoke-MgGraphRequest -Uri "$script:GraphBase/subscribedSkus" -Method GET -ErrorAction Stop
        $skus = @($response.value)

        Write-Host ($PREFIX_INFO + "Found $($skus.Count) subscribed SKU(s):") -ForegroundColor DarkGray
        foreach ($sku in $skus) {
            $friendly = Get-FriendlyLicenseName -SkuPartNumber $sku.skuPartNumber
            $available = $sku.prepaidUnits.enabled - $sku.consumedUnits
            Write-Host ($PREFIX_INFO + "  - $friendly ($($sku.skuPartNumber)): $($sku.consumedUnits) assigned, $available available") -ForegroundColor DarkGray
        }

        return $skus
    }
    catch {
        Write-Host ($PREFIX_WARN + "Failed to fetch subscribed SKUs: $($_.Exception.Message)") -ForegroundColor Yellow
        Write-Host ($PREFIX_INFO + 'Check that the required Graph permission (Organization.Read.All) is granted') -ForegroundColor DarkGray
        throw
    }
}

<#
.SYNOPSIS
    Retrieves assigned license details for each in-scope user (from ADUsers.json) via per-user Graph lookups.
#>
function Get-UserLicenseData {
    param(
        [PSCustomObject]$Context,
        [object[]]$SubscribedSkus
    )

    $adUsersPath = Join-Path $Context.RawPath 'ADUsers.json'
    if (-not (Test-Path $adUsersPath)) {
        Write-Host ($PREFIX_WARN + 'ADUsers.json not found - user license lookup skipped') -ForegroundColor Yellow
        Write-Host ($PREFIX_INFO + "Expected path: $adUsersPath") -ForegroundColor DarkGray
        return @()
    }

    Write-Host ($PREFIX_INFO + 'Building SKU lookup table from subscribed SKUs...') -ForegroundColor DarkGray
    $skuLookup = @{}
    foreach ($sku in $SubscribedSkus) {
        if ($sku.skuId) {
            $skuLookup[$sku.skuId.ToString()] = $sku.skuPartNumber
        }
    }
    Write-Host ($PREFIX_INFO + "SKU lookup table: $($skuLookup.Count) SKU(s)") -ForegroundColor DarkGray

    $adUsers = Get-Content $adUsersPath -Raw | ConvertFrom-Json
    Write-Host ($PREFIX_INFO + "Processing licenses for $($adUsers.Count) user(s)...") -ForegroundColor DarkGray

    $records = [System.Collections.Generic.List[PSCustomObject]]::new()
    $successCount = 0
    $errorCount = 0
    $usersWithLicenses = 0
    $usersWithoutLicenses = 0

    $current = 0
    foreach ($u in $adUsers) {
        $current++
        $upn = $u.UserPrincipalName
        if (-not $upn) {
            Write-Host ($PREFIX_WARN + "Skipping user without UPN (DisplayName: $($u.DisplayName))") -ForegroundColor Yellow
            continue
        }

        # Show progress every 50 users
        if ($current % 50 -eq 0) {
            Write-Host ($PREFIX_INFO + "Progress: $current of $($adUsers.Count) users processed (Success: $successCount, Errors: $errorCount)") -ForegroundColor DarkGray
        }

        try {
            $result = Invoke-MgGraphRequest -Uri "$script:GraphBase/users/$([Uri]::EscapeDataString($upn))?`$select=assignedLicenses,displayName" -Method GET -ErrorAction Stop
            $successCount++
        }
        catch {
            $errorCount++
            $errorMsg = $_.Exception.Message
            if ($errorMsg -like '*Request_ResourceNotFound*') {
                Write-Host ($PREFIX_WARN + "User not found in Entra ID: '$upn'") -ForegroundColor Yellow
            }
            elseif ($errorMsg -like '*Authorization_RequestDenied*' -or $errorMsg -like '*Forbidden*') {
                Write-Host ($PREFIX_WARN + "Authorization denied reading licenses for '$upn' - check Graph permissions") -ForegroundColor Yellow
            }
            else {
                Write-Host ($PREFIX_WARN + "Could not retrieve licenses for '$upn': $errorMsg") -ForegroundColor Yellow
            }
            continue
        }

        $skuPartNumbers = [System.Collections.Generic.List[string]]::new()
        $licenseNames   = [System.Collections.Generic.List[string]]::new()
        $skuIds         = [System.Collections.Generic.List[string]]::new()

        foreach ($lic in @($result.assignedLicenses)) {
            $skuId         = $lic.skuId.ToString()
            $skuPartNumber = if ($skuLookup.ContainsKey($skuId)) { $skuLookup[$skuId] } else { $skuId }
            $skuPartNumbers.Add($skuPartNumber)
            $licenseNames.Add((Get-FriendlyLicenseName -SkuPartNumber $skuPartNumber))
            $skuIds.Add($skuId)
        }

        if ($skuIds.Count -gt 0) { $usersWithLicenses++ } else { $usersWithoutLicenses++ }

        $records.Add([PSCustomObject]@{
            DisplayName       = $result.displayName
            UserPrincipalName = $upn
            LicenseCount      = $skuIds.Count
            SkuPartNumbers    = $skuPartNumbers -join '|'
            LicenseNames      = $licenseNames -join '|'
            SkuIds            = $skuIds -join '|'
        })
    }

    Write-Host ($PREFIX_INFO + "License lookup complete:") -ForegroundColor DarkGray
    Write-Host ($PREFIX_INFO + "  Total users processed: $current") -ForegroundColor DarkGray
    Write-Host ($PREFIX_INFO + "  Successful lookups: $successCount") -ForegroundColor DarkGray
    Write-Host ($PREFIX_INFO + "  Failed lookups: $errorCount") -ForegroundColor DarkGray
    Write-Host ($PREFIX_INFO + "  Users with licenses: $usersWithLicenses") -ForegroundColor DarkGray
    Write-Host ($PREFIX_INFO + "  Users without licenses: $usersWithoutLicenses") -ForegroundColor DarkGray

    if ($errorCount -gt 0) {
        Write-Host ($PREFIX_WARN + "License collection had $errorCount error(s) - check Graph API permissions") -ForegroundColor Yellow
        Write-Host ($PREFIX_INFO + "Required scopes: User.Read.All, Organization.Read.All") -ForegroundColor DarkGray
    }

    return $records.ToArray()
}

<#
.SYNOPSIS
    Builds per-SKU assigned license counts from the tenant's subscribed SKUs.
#>
function Get-LicenseCountData {
    param([object[]]$SubscribedSkus)

    return @($SubscribedSkus | ForEach-Object {
        [PSCustomObject]@{
            SkuPartNumber = $_.skuPartNumber
            LicenseName   = Get-FriendlyLicenseName -SkuPartNumber $_.skuPartNumber
            AssignedCount = $_.consumedUnits
        }
    })
}

<#
.SYNOPSIS
    Retrieves Planner plans for each VBU-scoped Team and M365 Group, recording a placeholder row when the plan list is unavailable due to permissions.
#>
function Get-PlannerPlanData {
    param(
        [object[]]$Teams,
        [object[]]$M365Groups,
        [object[]]$AllUnifiedGroups
    )

    $groupMap = @{}
    foreach ($g in $AllUnifiedGroups) { $groupMap[$g.id] = $g }

    $seen     = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $groupIds = [System.Collections.Generic.List[string]]::new()
    foreach ($t in $Teams)      { if ($t.GroupId -and $seen.Add($t.GroupId)) { $groupIds.Add($t.GroupId) } }
    foreach ($g in $M365Groups) { if ($g.GroupId -and $seen.Add($g.GroupId)) { $groupIds.Add($g.GroupId) } }

    $records = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($groupId in $groupIds) {
        $group       = $groupMap[$groupId]
        $displayName = if ($group) { $group.displayName } else { $null }
        $email       = if ($group) { $group.mail } else { $null }

        try {
            $response = Invoke-MgGraphRequest -Uri "$script:GraphBase/groups/$groupId/planner/plans" -Method GET -ErrorAction Stop
            $plans    = @($response.value)

            foreach ($plan in $plans) {
                $records.Add([PSCustomObject]@{
                    GroupDisplayName = $displayName
                    GroupEmail       = $email
                    GraphGroupId     = $groupId
                    PlanTitle        = $plan.title
                    PlanId           = $plan.id
                    CreatedDateTime  = $plan.createdDateTime
                    Status           = 'Success'
                })
            }
        }
        catch {
            $records.Add([PSCustomObject]@{
                GroupDisplayName = $displayName
                GroupEmail       = $email
                GraphGroupId     = $groupId
                PlanTitle        = '(plan list unavailable)'
                PlanId           = $null
                CreatedDateTime  = $null
                Status           = 'PermissionDenied'
            })
        }
    }

    return $records.ToArray()
}

# -----------------------------------------------------------------------
# Public functions
# -----------------------------------------------------------------------

<#
.SYNOPSIS
    Orchestrates all Microsoft Graph collection workloads.
.DESCRIPTION
    Fetches unified groups tenant-wide once, derives VBU-scoped Teams and M365 Groups
    from that set, then collects Intune devices, app registrations, enterprise apps,
    user licenses, license counts, and Planner plans, writing Teams.json, M365Groups.json,
    IntuneDevices.json, AppRegistrations.json, EnterpriseApps.json, UserLicenses.json,
    LicenseCounts.json, and PlannerPlans.json.
    Failures are non-critical and are returned as a failed collector result.
#>
function Invoke-GraphCollection {
    [CmdletBinding()]
    param([Parameter(Mandatory)][PSCustomObject]$Context)

    $start = Get-Date
    Write-SectionHeader 'Microsoft Graph'

    try {
        Write-Host ($PREFIX_INFO + 'Fetching all unified groups (tenant-wide)...') -ForegroundColor DarkGray
        $allGroups = Get-UnifiedGroupsWithMembers
        Write-ProgressLine -Label 'Unified groups (total)' -Count $allGroups.Count

        Write-Host ($PREFIX_INFO + 'Collecting Teams...') -ForegroundColor DarkGray
        $teams = Get-TeamData -Context $Context -AllUnifiedGroups $allGroups
        Write-ProgressLine -Label 'Teams (VBU-scoped)' -Count $teams.Count

        Write-Host ($PREFIX_INFO + 'Collecting M365 Groups...') -ForegroundColor DarkGray
        $m365Groups = Get-M365GroupData -Context $Context -AllUnifiedGroups $allGroups
        Write-ProgressLine -Label 'M365 Groups (VBU-scoped)' -Count $m365Groups.Count

        Write-Host ($PREFIX_INFO + 'Collecting Intune devices...') -ForegroundColor DarkGray
        $intuneDevices = Get-IntuneDeviceData -Context $Context
        Write-ProgressLine -Label 'Intune Devices' -Count $intuneDevices.Count

        Write-Host ($PREFIX_INFO + 'Collecting app registrations...') -ForegroundColor DarkGray
        $appRegistrations = Get-AppRegistrationData -Context $Context
        Write-ProgressLine -Label 'App Registrations (VBU-scoped)' -Count $appRegistrations.Count

        Write-Host ($PREFIX_INFO + 'Collecting enterprise apps...') -ForegroundColor DarkGray
        $enterpriseApps = Get-EnterpriseAppData -Context $Context
        Write-ProgressLine -Label 'Enterprise Apps (VBU-scoped)' -Count $enterpriseApps.Count

        Write-Host ($PREFIX_INFO + 'Fetching subscribed SKUs...') -ForegroundColor DarkGray
        $subscribedSkus = Get-SubscribedSkuData
        Write-ProgressLine -Label 'Subscribed SKUs (tenant-wide)' -Count $subscribedSkus.Count

        Write-Host ($PREFIX_INFO + 'Collecting user licenses...') -ForegroundColor DarkGray
        $userLicenses = Get-UserLicenseData -Context $Context -SubscribedSkus $subscribedSkus
        Write-ProgressLine -Label 'User Licenses (VBU-scoped)' -Count $userLicenses.Count

        Write-Host ($PREFIX_INFO + 'Collecting license counts...') -ForegroundColor DarkGray
        $licenseCounts = Get-LicenseCountData -SubscribedSkus $subscribedSkus
        Write-ProgressLine -Label 'License Counts (tenant-wide)' -Count $licenseCounts.Count

        Write-Host ($PREFIX_INFO + 'Collecting Planner plans...') -ForegroundColor DarkGray
        $plannerPlans = Get-PlannerPlanData -Teams $teams -M365Groups $m365Groups -AllUnifiedGroups $allGroups
        Write-ProgressLine -Label 'Planner Plans (VBU-scoped)' -Count $plannerPlans.Count

        Write-JsonOutput -FileName 'Teams.json'            -Data $teams            -RawPath $Context.RawPath
        Write-JsonOutput -FileName 'M365Groups.json'       -Data $m365Groups       -RawPath $Context.RawPath
        Write-JsonOutput -FileName 'IntuneDevices.json'    -Data $intuneDevices    -RawPath $Context.RawPath
        Write-JsonOutput -FileName 'AppRegistrations.json' -Data $appRegistrations -RawPath $Context.RawPath
        Write-JsonOutput -FileName 'EnterpriseApps.json'   -Data $enterpriseApps   -RawPath $Context.RawPath
        Write-JsonOutput -FileName 'UserLicenses.json'     -Data $userLicenses     -RawPath $Context.RawPath
        Write-JsonOutput -FileName 'LicenseCounts.json'    -Data $licenseCounts    -RawPath $Context.RawPath
        Write-JsonOutput -FileName 'PlannerPlans.json'     -Data $plannerPlans     -RawPath $Context.RawPath

        $counts = @{
            TeamCount            = $teams.Count
            M365GroupCount       = $m365Groups.Count
            IntuneDeviceCount    = $intuneDevices.Count
            AppRegistrationCount = $appRegistrations.Count
            EnterpriseAppCount   = $enterpriseApps.Count
            UserLicenseCount     = $userLicenses.Count
            LicenseCountCount    = $licenseCounts.Count
            PlannerPlanCount     = $plannerPlans.Count
        }
        $msg = "Teams: $($counts.TeamCount), M365 Groups: $($counts.M365GroupCount), " +
               "Intune Devices: $($counts.IntuneDeviceCount), App Registrations: $($counts.AppRegistrationCount), " +
               "Enterprise Apps: $($counts.EnterpriseAppCount), User Licenses: $($counts.UserLicenseCount), " +
               "License SKUs: $($counts.LicenseCountCount), Planner Plans: $($counts.PlannerPlanCount)"
        Update-CollectorStatus -CollectorName 'Graph Data' -Status 'Complete' `
            -RawPath $Context.RawPath -StartTime $start -Message $msg

        return New-CollectorResult -Success $true -Counts $counts
    }
    catch {
        Write-Host ($PREFIX_FAIL + 'Graph collection failed: ' + $_.Exception.Message) -ForegroundColor Red
        Update-CollectorStatus -CollectorName 'Graph Data' -Status 'Failed' `
            -RawPath $Context.RawPath -StartTime $start -Message $_.Exception.Message
        return New-CollectorResult -Success $false -ErrorMessage $_.Exception.Message
    }
}

# -----------------------------------------------------------------------
# Exports
# -----------------------------------------------------------------------

Export-ModuleMember -Function 'Invoke-GraphCollection'

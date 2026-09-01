#Requires -Version 7.0

Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -DisableNameChecking -Force -Global

# Graph API base
$script:GraphBase = 'https://graph.microsoft.com/v1.0'

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
    $uri    = "$script:GraphBase/groups?`$filter=groupTypes/any(c:c+eq+'Unified')&`$select=$select&`$top=999"
    $groups = [System.Collections.Generic.List[object]]::new()

    do {
        $response = Invoke-MgGraphRequest -Uri $uri -Method GET -ErrorAction Stop
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
        ($_.displayName -like "*$searchTerm*") -or
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
        ($_.displayName -like "*$searchTerm*") -or
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
            ProxyAddresses      = ($g.proxyAddresses | Sort-Object) -join '|'
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
                        # Entra object ID (the 'id' property) - distinct from the hardware DeviceId
                        # above. Update-MgDevice/Remove-MgDevice's -DeviceId parameter actually
                        # expects this object ID despite the parameter name; the Domain Removal CSV
                        # compatibility layer (LegacyExport.psm1) needs it kept separate.
                        EntraObjectId                  = $d.id
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

        $matched = ($app.displayName -like "*$t*") -or
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

# -----------------------------------------------------------------------
# Public functions
# -----------------------------------------------------------------------

<#
.SYNOPSIS
    Orchestrates all Microsoft Graph collection workloads.
.DESCRIPTION
    Fetches unified groups tenant-wide once, derives VBU-scoped Teams and M365 Groups
    from that set, then collects Intune devices and app registrations, writing
    Teams.json, M365Groups.json, IntuneDevices.json, and AppRegistrations.json.
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

        Write-JsonOutput -FileName 'Teams.json'            -Data $teams            -RawPath $Context.RawPath
        Write-JsonOutput -FileName 'M365Groups.json'       -Data $m365Groups       -RawPath $Context.RawPath
        Write-JsonOutput -FileName 'IntuneDevices.json'    -Data $intuneDevices    -RawPath $Context.RawPath
        Write-JsonOutput -FileName 'AppRegistrations.json' -Data $appRegistrations -RawPath $Context.RawPath

        $counts = @{
            TeamCount            = $teams.Count
            M365GroupCount       = $m365Groups.Count
            IntuneDeviceCount    = $intuneDevices.Count
            AppRegistrationCount = $appRegistrations.Count
        }
        $msg = "Teams: $($counts.TeamCount), M365 Groups: $($counts.M365GroupCount), " +
               "Intune Devices: $($counts.IntuneDeviceCount), App Registrations: $($counts.AppRegistrationCount)"
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

#Requires -Version 7.0
<#
.SYNOPSIS
    Update-EntraCache.ps1 — Refreshes the tenant-wide Entra ID users and M365 Groups cache

.DESCRIPTION
    Caches all Entra ID users and M365 Groups (unified groups) from Microsoft Graph. This allows
    assessments to run without connecting to Graph for user/group data, significantly speeding up
    the process and avoiding Graph API throttling.

    The cache includes:
    - All Entra ID users with key properties (UPN, DisplayName, Mail, Licenses, etc.)
    - All M365 Groups with members

.PARAMETER SharePointAdminUrl
    SPO admin URL (e.g., https://ourvolaris-admin.sharepoint.com) - used as the cache key
    to match what Run-Assessment.ps1 uses, so both agree on which cache folder to use.

.PARAMETER CacheRoot
    Base folder for cache files. Defaults to Assessment\Cache next to this script - same default
    Run-Assessment.ps1 uses.

.EXAMPLE
    .\Update-EntraCache.ps1 -SharePointAdminUrl https://ourvolaris-admin.sharepoint.com
#>

param(
    [Parameter(Mandatory)]
    [string]$SharePointAdminUrl,

    [string]$CacheRoot
)

$ErrorActionPreference = 'Stop'

$moduleRoot = Join-Path $PSScriptRoot 'Assessment\Modules'
Import-Module (Join-Path $moduleRoot 'Common.psm1') -Force -DisableNameChecking -Global

$cacheRootPath = if ($CacheRoot) { $CacheRoot } else { Join-Path $PSScriptRoot 'Assessment\Cache' }
$cacheFolder   = Get-DiscoveryCacheFolder -SharePointAdminUrl $SharePointAdminUrl -CacheRoot $cacheRootPath
New-Item -ItemType Directory -Path $cacheFolder -Force | Out-Null

Write-Host ''
Write-Host 'Entra ID Users & M365 Groups Cache Refresh' -ForegroundColor Cyan
Write-Host ('=' * 50) -ForegroundColor Cyan
Write-Host ''

# Pin Microsoft.Graph.* to 2.33.0 (last version before mandatory WAM)
. (Join-Path $PSScriptRoot 'Ensure-GraphModules.ps1') -GraphModules @('Microsoft.Graph.Users', 'Microsoft.Graph.Groups')

try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
Write-Host ($PREFIX_INFO + 'Connecting to Microsoft Graph...') -ForegroundColor DarkGray
try {
    Connect-MgGraph -Scopes @(
        'User.Read.All', 'Group.Read.All', 'GroupMember.Read.All'
    ) -NoWelcome -ErrorAction Stop
    Write-Host ($PREFIX_OK + 'Connected') -ForegroundColor Green
}
catch {
    Write-Host ($PREFIX_FAIL + 'Graph connection failed: ' + $_.Exception.Message) -ForegroundColor Red
    return
}

try {
    # Step 1: Get all Entra ID users
    Write-Host ($PREFIX_INFO + 'Collecting all Entra ID users...') -ForegroundColor DarkGray

    $userProps = @(
        'id', 'userPrincipalName', 'displayName', 'givenName', 'surname', 'mail',
        'jobTitle', 'department', 'officeLocation', 'mobilePhone', 'businessPhones',
        'companyName', 'userType', 'accountEnabled', 'assignedLicenses',
        'proxyAddresses', 'onPremisesSamAccountName', 'onPremisesDistinguishedName',
        'onPremisesSyncEnabled', 'createdDateTime'
    )

    $allUsers = @(Get-MgUser -All -Property ($userProps -join ',') -ErrorAction Stop)
    Write-Host ($PREFIX_OK + "Found $($allUsers.Count) Entra ID users") -ForegroundColor Green

    # Build user records
    $userRecords = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($user in $allUsers) {
        $userRecords.Add([PSCustomObject]@{
            Id                      = $user.Id
            UserPrincipalName       = $user.UserPrincipalName
            DisplayName             = $user.DisplayName
            GivenName               = $user.GivenName
            Surname                 = $user.Surname
            Mail                    = $user.Mail
            JobTitle                = $user.JobTitle
            Department              = $user.Department
            OfficeLocation          = $user.OfficeLocation
            MobilePhone             = $user.MobilePhone
            BusinessPhones          = if ($user.BusinessPhones) { $user.BusinessPhones -join '|' } else { '' }
            CompanyName             = $user.CompanyName
            UserType                = $user.UserType
            AccountEnabled          = $user.AccountEnabled
            AssignedLicenses        = if ($user.AssignedLicenses) { ($user.AssignedLicenses | ForEach-Object { $_.SkuId }) -join '|' } else { '' }
            ProxyAddresses          = if ($user.ProxyAddresses) { $user.ProxyAddresses -join '|' } else { '' }
            OnPremisesSamAccountName = $user.OnPremisesSamAccountName
            OnPremisesDistinguishedName = $user.OnPremisesDistinguishedName
            OnPremisesSyncEnabled   = $user.OnPremisesSyncEnabled
            CreatedDateTime         = $user.CreatedDateTime
        })
    }

    $usersFile = Join-Path $cacheFolder 'EntraUsers-All.json'
    $userRecords.ToArray() | ConvertTo-Json -Depth 10 | Set-Content -Path $usersFile -Encoding UTF8
    Write-Host ($PREFIX_OK + "Cached $($userRecords.Count) Entra ID users to $usersFile") -ForegroundColor Green

    # Step 2: Get all M365 Groups
    Write-Host ''
    Write-Host ($PREFIX_INFO + 'Collecting all M365 Groups...') -ForegroundColor DarkGray

    $groupProps = @(
        'id', 'displayName', 'mail', 'mailEnabled', 'mailNickname',
        'description', 'visibility', 'groupTypes', 'securityEnabled',
        'createdDateTime', 'proxyAddresses', 'resourceProvisioningOptions'
    )

    $allGroups = @(Get-MgGroup -All -Property ($groupProps -join ',') -Filter "groupTypes/any(c:c eq 'Unified')" -ErrorAction Stop)
    Write-Host ($PREFIX_OK + "Found $($allGroups.Count) M365 Groups") -ForegroundColor Green

    # Build group records with members
    Write-Host ($PREFIX_INFO + 'Collecting group members...') -ForegroundColor DarkGray
    $groupRecords = [System.Collections.Generic.List[PSCustomObject]]::new()
    $processed = 0
    $total = $allGroups.Count

    foreach ($group in $allGroups) {
        $processed++
        if ($processed % 50 -eq 0 -or $processed -eq $total) {
            $pct = [math]::Round(($processed / $total) * 100, 0)
            Write-Host ($PREFIX_INFO + "  $processed / $total ($pct%)") -ForegroundColor DarkGray
        }

        # Get group members
        $members = @()
        try {
            $members = @(Get-MgGroupMember -GroupId $group.Id -All -ErrorAction Stop |
                Select-Object -ExpandProperty AdditionalProperties |
                ForEach-Object { $_['userPrincipalName'] } |
                Where-Object { $_ })
        }
        catch {
            # Some groups may not have members or may error - continue
        }

        $groupRecords.Add([PSCustomObject]@{
            Id                      = $group.Id
            DisplayName             = $group.DisplayName
            Mail                    = $group.Mail
            MailEnabled             = $group.MailEnabled
            MailNickname            = $group.MailNickname
            Description             = $group.Description
            Visibility              = $group.Visibility
            GroupTypes              = if ($group.GroupTypes) { $group.GroupTypes -join '|' } else { '' }
            SecurityEnabled         = $group.SecurityEnabled
            CreatedDateTime         = $group.CreatedDateTime
            ProxyAddresses          = if ($group.ProxyAddresses) { $group.ProxyAddresses -join '|' } else { '' }
            ResourceProvisioningOptions = if ($group.ResourceProvisioningOptions) { $group.ResourceProvisioningOptions -join '|' } else { '' }
            Members                 = $members -join '|'
            MemberCount             = $members.Count
        })
    }

    $groupsFile = Join-Path $cacheFolder 'M365Groups-All.json'
    $groupRecords.ToArray() | ConvertTo-Json -Depth 10 | Set-Content -Path $groupsFile -Encoding UTF8
    Write-Host ($PREFIX_OK + "Cached $($groupRecords.Count) M365 Groups with members to $groupsFile") -ForegroundColor Green

    # Update cache completion status
    Update-CacheCompletionStatus -CacheFolder $cacheFolder -CacheType 'EntraUsersAndGroups'

    Write-Host ''
    Write-Host ($PREFIX_OK + "Entra cache refreshed: $cacheFolder") -ForegroundColor Green
}
catch {
    Write-Host ($PREFIX_FAIL + 'Cache refresh failed: ' + $_.Exception.Message) -ForegroundColor Red
}
finally {
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
}

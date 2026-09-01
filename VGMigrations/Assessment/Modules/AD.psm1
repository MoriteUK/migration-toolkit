#Requires -Version 7.0

Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -DisableNameChecking -Force -Global

# ActiveDirectory is only imported when actually needed (Context.SkipAD is false) - see
# Invoke-ADCollection. Importing it unconditionally at module-load time would fail on a
# machine without RSAT before Invoke-ADCollection gets a chance to check SkipAD.

$script:ManagerCache = @{}

$script:ADUserProperties = @(
    'DisplayName', 'GivenName', 'Surname', 'UserPrincipalName', 'Mail',
    'Department', 'Title', 'Company', 'Manager', 'EmployeeID',
    'extensionAttribute6', 'extensionAttribute7', 'proxyAddresses',
    'ObjectGUID', 'SID', 'Enabled', 'DistinguishedName', 'SamAccountName'
)

$script:ADGroupProperties = @(
    'DisplayName', 'Name', 'SamAccountName', 'Description', 'Mail',
    'proxyAddresses', 'extensionAttribute6', 'extensionAttribute7',
    'GroupScope', 'GroupCategory', 'DistinguishedName'
)

$script:ADComputerProperties = @(
    'Name', 'ObjectGUID', 'OperatingSystem', 'OperatingSystemVersion',
    'LastLogonDate', 'ManagedBy', 'Description', 'SamAccountName',
    'DistinguishedName', 'Enabled'
)

# -----------------------------------------------------------------------
# Private functions
# -----------------------------------------------------------------------

<#
.SYNOPSIS
    Resolves a manager distinguished name to a display name using a session-scoped cache.
#>
function Resolve-ManagerName {
    param([string]$ManagerDN)
    if ([string]::IsNullOrEmpty($ManagerDN)) { return '' }
    if ($script:ManagerCache.ContainsKey($ManagerDN)) { return $script:ManagerCache[$ManagerDN] }
    try {
        $mgr  = Get-ADUser -Identity $ManagerDN -Properties DisplayName -ErrorAction Stop
        $name = if ($mgr.DisplayName) { $mgr.DisplayName } else { $mgr.Name }
    }
    catch { $name = '' }
    $script:ManagerCache[$ManagerDN] = $name
    return $name
}

<#
.SYNOPSIS
    Retrieves all AD users and filters client-side to the VBU scope on name, UPN, proxy addresses, and extension attributes.
#>
function Get-VBUScopedUsers {
    param([PSCustomObject]$Context)

    $t = $Context.VBUSearchTerm
    $d = $Context.VBUDomain
    $i = $Context.VBUId

    # proxyAddresses -like with wildcards is unreliable in LDAP filters on multi-valued attributes.
    # Full client-side filter avoids double-query deduplication and catches all match paths correctly.
    $raw = Get-ADUser -Filter * -Properties $script:ADUserProperties -ErrorAction Stop |
        Where-Object {
            ($_.DisplayName         -like "*$t*") -or
            ($_.UserPrincipalName   -like "*$t*") -or
            ($_.extensionAttribute6 -eq $d)        -or
            ($_.extensionAttribute7 -eq $i)        -or
            ($_.proxyAddresses -and ($_.proxyAddresses | Where-Object { $_ -like "*$t*" }))
        }

    return @($raw | ForEach-Object {
        [PSCustomObject]@{
            DisplayName         = $_.DisplayName
            GivenName           = $_.GivenName
            Surname             = $_.Surname
            UserPrincipalName   = $_.UserPrincipalName
            Mail                = $_.Mail
            Department          = $_.Department
            Title               = $_.Title
            Company             = $_.Company
            Manager             = Resolve-ManagerName -ManagerDN $_.Manager
            EmployeeID          = $_.EmployeeID
            ExtensionAttribute6 = $_.extensionAttribute6
            ExtensionAttribute7 = $_.extensionAttribute7
            ProxyAddresses      = ($_.proxyAddresses | Sort-Object) -join '|'
            ObjectGUID          = $_.ObjectGUID.ToString()
            SID                 = $_.SID.ToString()
            Enabled             = $_.Enabled
            # Included for device correlation in Get-ADDeviceData and the Domain Removal
            # CSV compatibility layer (LegacyExport.psm1) - not a workbook field
            DistinguishedName   = $_.DistinguishedName
            SamAccountName      = $_.SamAccountName
        }
    })
}

<#
.SYNOPSIS
    Retrieves AD groups matching the VBU search term on Name, mail, or proxyAddresses via server-side filter.
#>
function Get-VBUScopedGroups {
    param([PSCustomObject]$Context)

    $t      = $Context.VBUSearchTerm
    $filter = "(Name -like '*$t*') -or (mail -like '*$t*') -or (proxyAddresses -like '*$t*')"

    $raw = Get-ADGroup -Filter $filter -Properties $script:ADGroupProperties -ErrorAction Stop

    return @($raw | ForEach-Object {
        [PSCustomObject]@{
            DisplayName         = $_.DisplayName
            Name                = $_.Name
            SamAccountName      = $_.SamAccountName
            Description         = $_.Description
            Mail                = $_.Mail
            ProxyAddresses      = ($_.proxyAddresses | Sort-Object) -join '|'
            ExtensionAttribute6 = $_.extensionAttribute6
            ExtensionAttribute7 = $_.extensionAttribute7
            GroupScope          = $_.GroupScope.ToString()
            GroupCategory       = $_.GroupCategory.ToString()
            DistinguishedName   = $_.DistinguishedName
        }
    })
}

<#
.SYNOPSIS
    Enumerates the members of each matched AD group and returns flattened membership records.
#>
function Get-ADGroupMembershipData {
    param([PSCustomObject[]]$Groups)

    $records = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($group in $Groups) {
        $groupName = if ($group.DisplayName) { $group.DisplayName } else { $group.Name }
        try {
            $members = Get-ADGroupMember -Identity $group.DistinguishedName -ErrorAction Stop
            foreach ($member in $members) {
                $memberType = switch ($member.objectClass) {
                    'user'     { 'User' }
                    'group'    { 'Group' }
                    'computer' { 'Computer' }
                    default    { $member.objectClass }
                }
                $memberName = $member.Name
                $memberUPN  = ''
                if ($member.objectClass -eq 'user') {
                    try {
                        $u          = Get-ADUser -Identity $member.DistinguishedName -Properties DisplayName -ErrorAction Stop
                        $memberName = if ($u.DisplayName) { $u.DisplayName } else { $u.Name }
                        $memberUPN  = $u.UserPrincipalName
                    }
                    catch { }
                }
                $records.Add([PSCustomObject]@{
                    GroupName  = $groupName
                    MemberName = $memberName
                    MemberType = $memberType
                    MemberUPN  = $memberUPN
                })
            }
        }
        catch {
            Write-Host ($PREFIX_WARN + "Could not enumerate members of '$groupName': " + $_.Exception.Message) -ForegroundColor Yellow
        }
    }

    return $records.ToArray()
}

<#
.SYNOPSIS
    Retrieves all AD computers and keeps those correlated to in-scope users via ManagedBy DN or SAM account in Description.
#>
function Get-ADDeviceData {
    param(
        [PSCustomObject[]]$InScopeUsers
    )

    $dnToUser  = @{}
    $samToUser = @{}
    foreach ($u in $InScopeUsers) {
        if ($u.DistinguishedName) { $dnToUser[$u.DistinguishedName]        = $u }
        if ($u.SamAccountName)    { $samToUser[$u.SamAccountName.ToLower()] = $u }
    }

    $allComputers = Get-ADComputer -Filter * -Properties $script:ADComputerProperties -ErrorAction Stop
    $records      = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($computer in $allComputers) {
        $managedByUser   = $null
        $descriptionUser = $null

        if ($computer.ManagedBy -and $dnToUser.ContainsKey($computer.ManagedBy)) {
            $managedByUser = $dnToUser[$computer.ManagedBy]
        }

        if ($computer.Description) {
            $descLower = $computer.Description.ToLower()
            foreach ($sam in $samToUser.Keys) {
                if ($descLower.Contains($sam)) {
                    $descriptionUser = $samToUser[$sam]
                    break
                }
            }
        }

        if ($null -eq $managedByUser -and $null -eq $descriptionUser) { continue }

        $primaryUPN       = if ($managedByUser)   { $managedByUser.UserPrincipalName }   else { $descriptionUser.UserPrincipalName }
        $registeredOwners = if ($managedByUser)   { $managedByUser.UserPrincipalName }   else { '' }
        $registeredUsers  = if ($descriptionUser) { $descriptionUser.UserPrincipalName } else { '' }

        $records.Add([PSCustomObject]@{
            DeviceName                    = $computer.Name
            DeviceId                      = $computer.ObjectGUID.ToString()
            OperatingSystem               = $computer.OperatingSystem
            OperatingSystemVersion        = $computer.OperatingSystemVersion
            TrustType                     = 'DomainJoined'
            JoinType                      = 'AD Joined'
            Enabled                       = $computer.Enabled
            IsManaged                     = $false
            IsCompliant                   = $null
            ApproximateLastSignInDateTime = $computer.LastLogonDate
            PrimaryUserUPN                = $primaryUPN
            RegisteredOwners              = $registeredOwners
            RegisteredUsers               = $registeredUsers
            Source                        = 'ActiveDirectory'
        })
    }

    return $records.ToArray()
}

# -----------------------------------------------------------------------
# Public functions
# -----------------------------------------------------------------------

<#
.SYNOPSIS
    Orchestrates Active Directory collection of users, groups, memberships, and devices.
.DESCRIPTION
    Returns a skipped result immediately when Context.SkipAD is set (module missing or
    domain unreachable - see Prerequisites.psm1). Otherwise runs the four AD workloads
    against the inherited domain session and writes ADUsers.json, ADGroups.json,
    ADGroupMemberships.json, and ADDevices.json. AD collection is non-critical like every
    other collector: a failure here is recorded and returned, not rethrown, so it never
    stops the rest of the assessment.
#>
function Invoke-ADCollection {
    [CmdletBinding()]
    param([Parameter(Mandatory)][PSCustomObject]$Context)

    $start = Get-Date
    Write-SectionHeader 'Active Directory'

    if ($Context.SkipAD) {
        Write-Host ($PREFIX_SKIP + 'AD collection skipped (SkipAD = true)') -ForegroundColor Yellow
        Update-CollectorStatus -CollectorName 'AD Data' -Status 'Skipped' `
            -RawPath $Context.RawPath -StartTime $start
        Write-JsonOutput -FileName 'ADUsers.json'            -Data @() -RawPath $Context.RawPath
        Write-JsonOutput -FileName 'ADGroups.json'           -Data @() -RawPath $Context.RawPath
        Write-JsonOutput -FileName 'ADGroupMemberships.json' -Data @() -RawPath $Context.RawPath
        Write-JsonOutput -FileName 'ADDevices.json'          -Data @() -RawPath $Context.RawPath
        return New-CollectorResult -Skipped $true
    }

    try {
        Import-Module ActiveDirectory -DisableNameChecking -ErrorAction Stop

        Write-Host ($PREFIX_INFO + 'Collecting AD users...') -ForegroundColor DarkGray
        $users = Get-VBUScopedUsers -Context $Context
        Write-ProgressLine -Label 'AD Users' -Count $users.Count

        Write-Host ($PREFIX_INFO + 'Collecting AD groups...') -ForegroundColor DarkGray
        $groups = Get-VBUScopedGroups -Context $Context
        Write-ProgressLine -Label 'AD Groups' -Count $groups.Count

        Write-Host ($PREFIX_INFO + 'Collecting group memberships...') -ForegroundColor DarkGray
        $memberships = Get-ADGroupMembershipData -Groups $groups
        Write-ProgressLine -Label 'AD Group Memberships' -Count $memberships.Count

        Write-Host ($PREFIX_INFO + 'Collecting AD devices...') -ForegroundColor DarkGray
        $devices = Get-ADDeviceData -InScopeUsers $users
        Write-ProgressLine -Label 'AD Devices' -Count $devices.Count

        Write-JsonOutput -FileName 'ADUsers.json'            -Data $users        -RawPath $Context.RawPath
        Write-JsonOutput -FileName 'ADGroups.json'           -Data $groups       -RawPath $Context.RawPath
        Write-JsonOutput -FileName 'ADGroupMemberships.json' -Data $memberships  -RawPath $Context.RawPath
        Write-JsonOutput -FileName 'ADDevices.json'          -Data $devices      -RawPath $Context.RawPath

        $counts = @{
            UserCount       = $users.Count
            GroupCount      = $groups.Count
            MembershipCount = $memberships.Count
            DeviceCount     = $devices.Count
        }
        $msg = "Users: $($counts.UserCount), Groups: $($counts.GroupCount), " +
               "Memberships: $($counts.MembershipCount), Devices: $($counts.DeviceCount)"
        Update-CollectorStatus -CollectorName 'AD Data' -Status 'Complete' `
            -RawPath $Context.RawPath -StartTime $start -Message $msg

        return New-CollectorResult -Success $true -Counts $counts
    }
    catch {
        Write-Host ($PREFIX_FAIL + 'AD collection failed: ' + $_.Exception.Message) -ForegroundColor Red
        Update-CollectorStatus -CollectorName 'AD Data' -Status 'Failed' `
            -RawPath $Context.RawPath -StartTime $start -Message $_.Exception.Message
        Write-JsonOutput -FileName 'ADUsers.json'            -Data @() -RawPath $Context.RawPath
        Write-JsonOutput -FileName 'ADGroups.json'           -Data @() -RawPath $Context.RawPath
        Write-JsonOutput -FileName 'ADGroupMemberships.json' -Data @() -RawPath $Context.RawPath
        Write-JsonOutput -FileName 'ADDevices.json'          -Data @() -RawPath $Context.RawPath
        return New-CollectorResult -Success $false -ErrorMessage $_.Exception.Message
    }
}

# -----------------------------------------------------------------------
# Exports
# -----------------------------------------------------------------------

Export-ModuleMember -Function 'Invoke-ADCollection'

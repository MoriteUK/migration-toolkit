#Requires -Version 7.0

Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -DisableNameChecking -Force -Global
Import-Module ActiveDirectory -DisableNameChecking -ErrorAction Stop

$script:ManagerCache = @{}

$script:ADUserProperties = @(
    'DisplayName', 'GivenName', 'Surname', 'UserPrincipalName', 'Mail',
    'Department', 'Title', 'Company', 'Manager', 'EmployeeID',
    'proxyAddresses', 'ObjectGUID', 'SID', 'Enabled', 'DistinguishedName', 'SamAccountName',
    'Initials', 'EmailAddress', 'mailNickname', 'EmployeeNumber', 'EmployeeType', 'directReports',
    'physicalDeliveryOfficeName', 'telephoneNumber', 'mobile', 'homePhone', 'facsimileTelephoneNumber',
    'StreetAddress', 'l', 'st', 'PostalCode', 'c', 'postOfficeBox', 'Description', 'LockedOut',
    'PasswordExpired', 'PasswordLastSet', 'PasswordNeverExpires', 'AccountExpirationDate', 'LastLogonDate',
    'UserAccountControl', 'WhenCreated', 'WhenChanged', 'CanonicalName', 'targetAddress',
    'msExchHideFromAddressLists', 'msExchRecipientTypeDetails', 'msExchRemoteRecipientType',
    'extensionAttribute1', 'extensionAttribute2', 'extensionAttribute3', 'extensionAttribute4',
    'extensionAttribute5', 'extensionAttribute6', 'extensionAttribute7', 'extensionAttribute8',
    'extensionAttribute9', 'extensionAttribute10', 'extensionAttribute11', 'extensionAttribute12',
    'extensionAttribute13', 'extensionAttribute14', 'extensionAttribute15',
    'memberOf', 'servicePrincipalName'
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
    Retrieves all AD users and filters client-side to the VBU scope on tagging attributes and mail domain.
.DESCRIPTION
    Users are scoped by how they are tagged and by their mail domain - never by display name.
    A user is in scope when ANY of these holds: extensionAttribute7 exactly matches the VBU ID,
    extensionAttribute6 contains the VBU search term (it holds a name such as 'Bravura Security',
    not the domain), any proxyAddress contains the VBU domain, or the UPN contains the VBU domain.
    Name/search-term matching is a group, SharePoint, and Teams concern only.
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
            ($_.extensionAttribute7 -eq   $i)       -or
            (Test-WholeWordMatch -Text $_.extensionAttribute6 -Term $t) -or
            ($_.UserPrincipalName   -like "*$d*")   -or
            ($_.proxyAddresses -and ($_.proxyAddresses | Where-Object { $_ -like "*$d*" }))
        }

    return @($raw | ForEach-Object {
        $u = $_

        $imAddress = ''
        if ($u.proxyAddresses) {
            $sip = @($u.proxyAddresses | Where-Object { $_ -like 'sip:*' }) | Select-Object -First 1
            if ($sip) { $imAddress = $sip -replace '^sip:', '' }
        }

        $ou = if ($u.DistinguishedName) { $u.DistinguishedName -replace '^CN=[^,]+,', '' } else { '' }

        $primary = @($u.proxyAddresses | Where-Object { $_ -clike 'SMTP:*' }) | Select-Object -First 1

        $discoveredBy =
            if (($u.UserPrincipalName -like "*$d*") -or ($primary -like "*$d*"))       { 'PrimarySmtp/UPN' }
            elseif ($u.extensionAttribute7 -eq $i)                                     { 'VBU ID' }
            elseif (Test-WholeWordMatch -Text $u.extensionAttribute6 -Term $t)         { 'VBU Name' }
            elseif ($u.proxyAddresses -and ($u.proxyAddresses | Where-Object { $_ -like "*$d*" })) { 'Alias' }
            else                                                                       { '' }

        [PSCustomObject]@{
            DisplayName               = $_.DisplayName
            DiscoveredBy              = $discoveredBy
            GivenName                 = $_.GivenName
            Surname                   = $_.Surname
            UserPrincipalName         = $_.UserPrincipalName
            Mail                      = $_.Mail
            IMAddress                 = $imAddress
            Department                = $_.Department
            Title                     = $_.Title
            Company                   = $_.Company
            Manager                   = Resolve-ManagerName -ManagerDN $_.Manager
            EmployeeID                = $_.EmployeeID
            ProxyAddresses            = ($_.proxyAddresses | Sort-Object) -join '|'
            ObjectGUID                = $_.ObjectGUID.ToString()
            SID                       = $_.SID.ToString()
            Enabled                   = $_.Enabled
            Initials                  = $_.Initials
            EmailAddress              = $_.EmailAddress
            MailNickname              = $_.mailNickname
            EmployeeNumber            = $_.EmployeeNumber
            EmployeeType              = $_.EmployeeType
            ManagerDN                 = $_.Manager
            DirectReportCount         = if ($_.directReports) { @($_.directReports).Count } else { 0 }
            Office                    = $_.physicalDeliveryOfficeName
            OfficePhone               = $_.telephoneNumber
            MobilePhone               = $_.mobile
            HomePhone                 = $_.homePhone
            Fax                       = $_.facsimileTelephoneNumber
            StreetAddress             = $_.StreetAddress
            City                      = $_.l
            State                     = $_.st
            PostalCode                = $_.PostalCode
            Country                   = $_.c
            POBox                     = $_.postOfficeBox
            Description               = $_.Description
            LockedOut                 = $_.LockedOut
            PasswordExpired           = $_.PasswordExpired
            PasswordLastSet           = $_.PasswordLastSet
            PasswordNeverExpires      = $_.PasswordNeverExpires
            AccountExpirationDate     = $_.AccountExpirationDate
            LastLogonDate             = $_.LastLogonDate
            UserAccountControl        = $_.UserAccountControl
            WhenCreated               = $_.WhenCreated
            WhenChanged               = $_.WhenChanged
            OU                        = $ou
            CanonicalName             = $_.CanonicalName
            TargetAddress             = $_.targetAddress
            HideFromGAL               = $_.msExchHideFromAddressLists
            MsExchRecipientTypeDetails = $_.msExchRecipientTypeDetails
            MsExchRemoteRecipientType = $_.msExchRemoteRecipientType
            ExtensionAttribute1       = $_.extensionAttribute1
            ExtensionAttribute2       = $_.extensionAttribute2
            ExtensionAttribute3       = $_.extensionAttribute3
            ExtensionAttribute4       = $_.extensionAttribute4
            ExtensionAttribute5       = $_.extensionAttribute5
            ExtensionAttribute6       = $_.extensionAttribute6
            ExtensionAttribute7       = $_.extensionAttribute7
            ExtensionAttribute8       = $_.extensionAttribute8
            ExtensionAttribute9       = $_.extensionAttribute9
            ExtensionAttribute10      = $_.extensionAttribute10
            ExtensionAttribute11      = $_.extensionAttribute11
            ExtensionAttribute12      = $_.extensionAttribute12
            ExtensionAttribute13      = $_.extensionAttribute13
            ExtensionAttribute14      = $_.extensionAttribute14
            ExtensionAttribute15      = $_.extensionAttribute15
            GroupMembershipCount      = if ($_.memberOf) { @($_.memberOf).Count } else { 0 }
            ServicePrincipalNames     = if ($_.servicePrincipalName) { ($_.servicePrincipalName | Sort-Object) -join '|' } else { '' }
            # Included for device correlation in Get-ADDeviceData - not a workbook field
            DistinguishedName         = $_.DistinguishedName
            SamAccountName            = $_.SamAccountName
        }
    })
}

<#
.SYNOPSIS
    Retrieves AD groups matching the VBU search term as a whole word on Name, mail, or proxyAddresses.
.DESCRIPTION
    LDAP filters cannot express word boundaries, so the server-side -like filter deliberately
    over-includes (substring is a superset of whole-word) and a client-side Test-WholeWordMatch
    pass then drops the substring false positives - e.g. 'ceramics' for the term 'amic'.
#>
function Get-VBUScopedGroups {
    param([PSCustomObject]$Context)

    $t      = $Context.VBUSearchTerm
    $filter = "(Name -like '*$t*') -or (mail -like '*$t*') -or (proxyAddresses -like '*$t*')"

    $raw = Get-ADGroup -Filter $filter -Properties $script:ADGroupProperties -ErrorAction Stop |
        Where-Object {
            (Test-WholeWordMatch -Text $_.Name -Term $t) -or
            (Test-WholeWordMatch -Text $_.Mail -Term $t) -or
            ($_.proxyAddresses -and ($_.proxyAddresses | Where-Object { Test-WholeWordMatch -Text $_ -Term $t }))
        }

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
    Runs the four AD workloads against the inherited domain session and writes
    ADUsers.json, ADGroups.json, ADGroupMemberships.json, and ADDevices.json.
    AD collection is critical: on failure the status is recorded and the error
    is rethrown so the orchestrator stops the run.
#>
function Invoke-ADCollection {
    [CmdletBinding()]
    param([Parameter(Mandatory)][PSCustomObject]$Context)

    $start = Get-Date
    Write-SectionHeader 'Active Directory'

    try {
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
        throw  # AD failure is critical - rethrow so orchestrator stops the run
    }
}

# -----------------------------------------------------------------------
# Exports
# -----------------------------------------------------------------------

Export-ModuleMember -Function 'Invoke-ADCollection'

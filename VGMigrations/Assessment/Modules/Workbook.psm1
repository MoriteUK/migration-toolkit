#Requires -Version 7.0

Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -DisableNameChecking -Force -Global
Import-Module ImportExcel -DisableNameChecking -ErrorAction Stop

# -----------------------------------------------------------------------
# Private functions
# -----------------------------------------------------------------------

<#
.SYNOPSIS
    Appends a data worksheet to the workbook, writing a placeholder row when there is no data.
#>
function Write-WorksheetData {
    param(
        [string]$XlsxPath,
        [string]$WorksheetName,
        $Data
    )
    $hasData = $null -ne $Data -and @($Data).Count -gt 0
    $rows    = if ($hasData) { $Data } else {
        @([PSCustomObject]@{ Note = '(No data collected)' })
    }

    if ($hasData) {
        @($rows) | Export-Excel -Path $XlsxPath -WorksheetName $WorksheetName `
            -Append -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter
    }
    else {
        @($rows) | Export-Excel -Path $XlsxPath -WorksheetName $WorksheetName `
            -Append -AutoSize
    }
}

<#
.SYNOPSIS
    Applies the Team -> M365 Group -> SharePoint Site relationship logic to set MigrationObjectType and AssociatedObject on each record.
#>
function Set-MigrationObjectType {
    param(
        [object[]]$Teams,
        [object[]]$M365Groups,
        [object[]]$SPOSites
    )

    # GroupId -> Team DisplayName for all VBU-scoped teams
    $teamGroupIds = @{}
    foreach ($t in $Teams) {
        if ($t.GroupId) { $teamGroupIds[$t.GroupId] = $t.DisplayName }
    }

    # GroupId -> M365 Group DisplayName (all unified groups - used for SPO site fallback)
    $m365GroupIds = @{}
    foreach ($g in $M365Groups) {
        if ($g.GroupId) { $m365GroupIds[$g.GroupId] = $g.DisplayName }
    }

    # Teams are always "Migrated as Team"
    foreach ($t in $Teams) {
        $t.MigrationObjectType = 'Migrated as Team'
        $t.AssociatedObject    = $t.DisplayName
    }

    # M365 Groups - check against Team set first
    foreach ($g in $M365Groups) {
        if ($g.GroupId -and $teamGroupIds.ContainsKey($g.GroupId)) {
            $g.MigrationObjectType = 'Migrated as Team'
            $g.AssociatedObject    = $teamGroupIds[$g.GroupId]
        }
        else {
            $g.MigrationObjectType = 'Migrated as M365 Group'
            $g.AssociatedObject    = $null
        }
    }

    # Older SharePointSites.json predates IsOrphaned - add it so the assignments below don't throw
    foreach ($s in $SPOSites) {
        if (-not $s.PSObject.Properties['IsOrphaned']) {
            $s | Add-Member -NotePropertyName 'IsOrphaned' -NotePropertyValue $false
        }
    }

    # SPO Sites - channel-site check first (channel sites have no GroupId, only
    # RelatedGroupId pointing at the parent team), then Team, then M365 Group, then standalone
    foreach ($s in $SPOSites) {
        if ($s.IsTeamsChannelConnected -and $s.RelatedGroupId -and
            $teamGroupIds.ContainsKey($s.RelatedGroupId)) {
            $s.MigrationObjectType = 'Migrated as part of Team'
            $s.IsTeamsConnected    = $true
            $s.AssociatedObject    = $teamGroupIds[$s.RelatedGroupId]
            $s.IsOrphaned          = $false
        }
        elseif ($s.GroupId -and $teamGroupIds.ContainsKey($s.GroupId)) {
            $s.MigrationObjectType = 'Migrated as part of Team'
            $s.IsTeamsConnected    = $true
            $s.AssociatedObject    = $teamGroupIds[$s.GroupId]
            $s.IsOrphaned          = $false
        }
        elseif ($s.GroupId -and $m365GroupIds.ContainsKey($s.GroupId)) {
            $s.MigrationObjectType = 'Migrated as part of M365 Group'
            $s.IsTeamsConnected    = $false
            $s.AssociatedObject    = $m365GroupIds[$s.GroupId]
            $s.IsOrphaned          = $false
        }
        elseif ($s.GroupId) {
            # Has a GroupId but it resolved to nothing - orphaned site
            $s.MigrationObjectType = 'Migrate as SharePoint Site'
            $s.IsTeamsConnected    = $false
            $s.AssociatedObject    = $null
            $s.IsOrphaned          = $true
        }
        else {
            # Truly standalone - no group association
            $s.MigrationObjectType = 'Migrate as SharePoint Site'
            $s.IsTeamsConnected    = $false
            $s.AssociatedObject    = $null
            $s.IsOrphaned          = $false
        }
    }

    return @{
        Teams      = $Teams
        M365Groups = $M365Groups
        SPOSites   = $SPOSites
    }
}

<#
.SYNOPSIS
    Combines Intune and AD device records, deduplicating case-insensitively on DeviceName with Intune as the preferred source.
#>
function Merge-DeviceData {
    param(
        [object[]]$IntuneDevices,
        [object[]]$ADDevices
    )

    $seen    = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $records = [System.Collections.Generic.List[object]]::new()

    foreach ($d in $IntuneDevices) {
        if ($d.DeviceName) { [void]$seen.Add($d.DeviceName) }
        $records.Add($d)
    }

    foreach ($d in $ADDevices) {
        if ($d.DeviceName -and -not $seen.Contains($d.DeviceName)) {
            $records.Add($d)
        }
    }

    return $records.ToArray()
}

# -----------------------------------------------------------------------
# Public functions
# -----------------------------------------------------------------------

<#
.SYNOPSIS
    Reads all raw JSON, derives migration dispositions, and writes the assessment workbook.
.DESCRIPTION
    Loads every collector JSON file, classifies Teams, M365 Groups, and SharePoint sites
    via Set-MigrationObjectType, merges Intune and AD device data, then writes the
    Assessment Summary and all data sheets to the xlsx via ImportExcel. Rethrows on
    failure so the orchestrator can report the error.
#>
function Invoke-WorkbookGeneration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][PSCustomObject]$Context,
        [Parameter(Mandatory)][string]$XlsxPath
    )

    Write-SectionHeader 'Workbook'

    try {
        Write-Host ($PREFIX_INFO + 'Loading JSON data...') -ForegroundColor DarkGray

        $adUsers           = Import-AssessmentJson -FileName 'ADUsers.json'                   -RawPath $Context.RawPath
        $adGroups          = Import-AssessmentJson -FileName 'ADGroups.json'                  -RawPath $Context.RawPath
        $adMemberships     = Import-AssessmentJson -FileName 'ADGroupMemberships.json'        -RawPath $Context.RawPath
        $adDevices         = Import-AssessmentJson -FileName 'ADDevices.json'                 -RawPath $Context.RawPath
        $userMailboxes     = Import-AssessmentJson -FileName 'UserMailboxes.json'             -RawPath $Context.RawPath
        $sharedMailboxes   = Import-AssessmentJson -FileName 'SharedMailboxes.json'           -RawPath $Context.RawPath
        $resourceMailboxes = Import-AssessmentJson -FileName 'ResourceMailboxes.json'         -RawPath $Context.RawPath
        $mailContacts      = Import-AssessmentJson -FileName 'MailContacts.json'              -RawPath $Context.RawPath
        $dgs               = Import-AssessmentJson -FileName 'DistributionGroups.json'        -RawPath $Context.RawPath
        $mesgs             = Import-AssessmentJson -FileName 'MailEnabledSecurityGroups.json' -RawPath $Context.RawPath
        $ddgs              = Import-AssessmentJson -FileName 'DynamicDistributionGroups.json' -RawPath $Context.RawPath
        $mailboxPerms      = Import-AssessmentJson -FileName 'MailboxPermissions.json'        -RawPath $Context.RawPath
        $transportRules    = Import-AssessmentJson -FileName 'TransportRules.json'            -RawPath $Context.RawPath
        $connectors        = Import-AssessmentJson -FileName 'Connectors.json'                -RawPath $Context.RawPath
        $acceptedDomains   = Import-AssessmentJson -FileName 'AcceptedDomains.json'           -RawPath $Context.RawPath
        $remoteDomains     = Import-AssessmentJson -FileName 'RemoteDomains.json'             -RawPath $Context.RawPath
        $journalRules      = Import-AssessmentJson -FileName 'JournalRules.json'              -RawPath $Context.RawPath
        $teams             = Import-AssessmentJson -FileName 'Teams.json'                     -RawPath $Context.RawPath
        $m365Groups        = Import-AssessmentJson -FileName 'M365Groups.json'                -RawPath $Context.RawPath
        $intuneDevices     = Import-AssessmentJson -FileName 'IntuneDevices.json'             -RawPath $Context.RawPath
        $appRegistrations  = Import-AssessmentJson -FileName 'AppRegistrations.json'          -RawPath $Context.RawPath
        $spoSites          = Import-AssessmentJson -FileName 'SharePointSites.json'           -RawPath $Context.RawPath
        $oneDrives         = Import-AssessmentJson -FileName 'OneDrives.json'                 -RawPath $Context.RawPath
        $powerPlatform     = Import-AssessmentJson -FileName 'PowerPlatform.json'             -RawPath $Context.RawPath

        Write-Host ($PREFIX_INFO + 'Classifying Teams, Groups, and Sites...') -ForegroundColor DarkGray

        $classified = Set-MigrationObjectType -Teams $teams -M365Groups $m365Groups -SPOSites $spoSites
        $teams      = $classified.Teams
        $m365Groups = $classified.M365Groups
        $spoSites   = $classified.SPOSites

        Write-Host ($PREFIX_INFO + 'Merging device data...') -ForegroundColor DarkGray

        $devices = Merge-DeviceData -IntuneDevices $intuneDevices -ADDevices $adDevices

        # Build Assessment Summary - two sections, two columns
        $summaryRows = @(
            [PSCustomObject]@{ Section = 'ASSESSMENT INFORMATION'; Value = '' }
            [PSCustomObject]@{ Section = 'VBU Domain';             Value = $Context.VBUDomain }
            [PSCustomObject]@{ Section = 'VBU Name';               Value = $Context.VBUName }
            [PSCustomObject]@{ Section = 'VBU ID';             Value = $Context.VBUId }
            [PSCustomObject]@{ Section = 'VBU Search Term';  Value = $Context.VBUSearchTerm }
            [PSCustomObject]@{ Section = 'Assessment Date';   Value = $Context.AssessmentDate.ToString('yyyy-MM-dd HH:mm') }
            [PSCustomObject]@{ Section = '';                        Value = '' }
            [PSCustomObject]@{ Section = 'INVENTORY COUNTS';                Value = '' }
            [PSCustomObject]@{ Section = 'AD Users';                        Value = $adUsers.Count }
            [PSCustomObject]@{ Section = 'AD Groups';                       Value = $adGroups.Count }
            [PSCustomObject]@{ Section = 'AD Group Memberships';            Value = $adMemberships.Count }
            [PSCustomObject]@{ Section = 'SharePoint Sites';                Value = $spoSites.Count }
            [PSCustomObject]@{ Section = 'OneDrives';                       Value = $oneDrives.Count }
            [PSCustomObject]@{ Section = 'Teams';                           Value = $teams.Count }
            [PSCustomObject]@{ Section = 'M365 Groups';                     Value = $m365Groups.Count }
            [PSCustomObject]@{ Section = 'User Mailboxes';                  Value = $userMailboxes.Count }
            [PSCustomObject]@{ Section = 'Shared Mailboxes';                Value = $sharedMailboxes.Count }
            [PSCustomObject]@{ Section = 'Resource Mailboxes';              Value = $resourceMailboxes.Count }
            [PSCustomObject]@{ Section = 'Mail Contacts';                   Value = $mailContacts.Count }
            [PSCustomObject]@{ Section = 'Distribution Groups';             Value = $dgs.Count }
            [PSCustomObject]@{ Section = 'Mail-Enabled Security Groups';    Value = $mesgs.Count }
            [PSCustomObject]@{ Section = 'Dynamic Distribution Groups';     Value = $ddgs.Count }
            [PSCustomObject]@{ Section = 'Mailbox Permissions';             Value = $mailboxPerms.Count }
            [PSCustomObject]@{ Section = 'Transport Rules';                 Value = $transportRules.Count }
            [PSCustomObject]@{ Section = 'Connectors';                      Value = $connectors.Count }
            [PSCustomObject]@{ Section = 'Accepted Domains';                Value = $acceptedDomains.Count }
            [PSCustomObject]@{ Section = 'Remote Domains';                  Value = $remoteDomains.Count }
            [PSCustomObject]@{ Section = 'Journal Rules';                   Value = $journalRules.Count }
            [PSCustomObject]@{ Section = 'Devices';                         Value = $devices.Count }
            [PSCustomObject]@{ Section = 'App Registrations';              Value = $appRegistrations.Count }
            [PSCustomObject]@{ Section = 'Power Platform Objects';          Value = $powerPlatform.Count }
        )

        Write-Host ($PREFIX_INFO + 'Writing workbook...') -ForegroundColor DarkGray

        if (Test-Path $XlsxPath) { Remove-Item $XlsxPath -Force }

        # Sheet 1 - Assessment Summary (no table style)
        $summaryRows | Export-Excel -Path $XlsxPath -WorksheetName 'Assessment Summary' -AutoSize

        # Sheet 2 - AD Users: exclude internal correlation fields (DistinguishedName, SamAccountName)
        Write-WorksheetData -XlsxPath $XlsxPath -WorksheetName 'AD Users' -Data (
            $adUsers | Select-Object DisplayName, GivenName, Surname, UserPrincipalName, Mail,
                Department, Title, Company, Manager, EmployeeID,
                ExtensionAttribute6, ExtensionAttribute7, ProxyAddresses,
                ObjectGUID, SID, Enabled
        )

        # Sheet 3 - AD Groups: exclude internal DistinguishedName field
        Write-WorksheetData -XlsxPath $XlsxPath -WorksheetName 'AD Groups' -Data (
            $adGroups | Select-Object DisplayName, Name, SamAccountName, Description, Mail,
                ProxyAddresses, ExtensionAttribute6, ExtensionAttribute7,
                GroupScope, GroupCategory
        )

        Write-WorksheetData -XlsxPath $XlsxPath -WorksheetName 'AD Group Memberships'        -Data $adMemberships
        Write-WorksheetData -XlsxPath $XlsxPath -WorksheetName 'SharePoint Sites'             -Data $spoSites
        Write-WorksheetData -XlsxPath $XlsxPath -WorksheetName 'OneDrives'                    -Data $oneDrives
        Write-WorksheetData -XlsxPath $XlsxPath -WorksheetName 'Teams'                        -Data $teams
        Write-WorksheetData -XlsxPath $XlsxPath -WorksheetName 'M365 Groups'                  -Data $m365Groups
        Write-WorksheetData -XlsxPath $XlsxPath -WorksheetName 'User Mailboxes'               -Data $userMailboxes
        Write-WorksheetData -XlsxPath $XlsxPath -WorksheetName 'Shared Mailboxes'             -Data $sharedMailboxes
        Write-WorksheetData -XlsxPath $XlsxPath -WorksheetName 'Resource Mailboxes'           -Data $resourceMailboxes
        Write-WorksheetData -XlsxPath $XlsxPath -WorksheetName 'Mail Contacts'                -Data $mailContacts
        Write-WorksheetData -XlsxPath $XlsxPath -WorksheetName 'Distribution Groups'          -Data $dgs
        Write-WorksheetData -XlsxPath $XlsxPath -WorksheetName 'Mail-Enabled Security Groups' -Data $mesgs
        Write-WorksheetData -XlsxPath $XlsxPath -WorksheetName 'Dynamic Distribution Groups'  -Data $ddgs
        Write-WorksheetData -XlsxPath $XlsxPath -WorksheetName 'Mailbox Permissions'          -Data $mailboxPerms
        Write-WorksheetData -XlsxPath $XlsxPath -WorksheetName 'Transport Rules'              -Data $transportRules
        Write-WorksheetData -XlsxPath $XlsxPath -WorksheetName 'Connectors'                   -Data $connectors
        Write-WorksheetData -XlsxPath $XlsxPath -WorksheetName 'Accepted Domains'             -Data $acceptedDomains
        Write-WorksheetData -XlsxPath $XlsxPath -WorksheetName 'Remote Domains'               -Data $remoteDomains
        Write-WorksheetData -XlsxPath $XlsxPath -WorksheetName 'Journal Rules'                -Data $journalRules
        Write-WorksheetData -XlsxPath $XlsxPath -WorksheetName 'Devices'                      -Data $devices
        Write-WorksheetData -XlsxPath $XlsxPath -WorksheetName 'App Registrations'            -Data $appRegistrations
        Write-WorksheetData -XlsxPath $XlsxPath -WorksheetName 'Power Platform'               -Data $powerPlatform

        Write-Host ($PREFIX_OK + 'Workbook written: ' + (Split-Path $XlsxPath -Leaf)) -ForegroundColor Green
    }
    catch {
        Write-Host ($PREFIX_FAIL + 'Workbook generation failed: ' + $_.Exception.Message) -ForegroundColor Red
        throw
    }
}

# -----------------------------------------------------------------------
# Exports
# -----------------------------------------------------------------------

Export-ModuleMember -Function 'Invoke-WorkbookGeneration'

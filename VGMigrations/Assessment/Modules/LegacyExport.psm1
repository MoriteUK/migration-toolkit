#Requires -Version 7.0

Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -DisableNameChecking -Force -Global

<#
.SYNOPSIS
    Writes search-domain.ps1-compatible CSVs from the new engine's JSON output.
.DESCRIPTION
    search-domain.ps1 is retired, but its <domain>\Discovery\NN_<Area>.csv output is a data
    contract read directly by 7 Domain Removal scripts (Rename-DomainObjects.ps1,
    Hide-AddressBook.ps1, Remove-AliasAddresses.ps1, remove-domain.ps1, Retire-Devices.ps1,
    Remove-devices.ps1, Restore-ProxyAddresses.ps1). Rather than touch those scripts - each
    individually battle-tested against real tenant WAM/Conditional-Access failures - this
    module reproduces their expected CSV files from the new JSON collectors, so the data
    still originates entirely from the new engine but every downstream consumer stays
    byte-for-byte unchanged.

    Some legacy columns simply aren't collected by the new AD.psm1/Exchange.psm1 (all 15 AD
    extension attributes, mailbox quotas/UAC/lockout/phone/address fields) - those come back
    blank. Verified against the actual consumers that this doesn't affect correctness: every
    downstream script resolves recipients by UserPrincipalName/PrimarySmtpAddress/
    ExternalEmailAddress and re-queries EXO/AD live for anything it mutates - the CSV only
    ever supplies the identity key plus the specific address rows RecipientProxyAddresses
    needs.
#>

# -----------------------------------------------------------------------
# Private helpers
# -----------------------------------------------------------------------

<#
.SYNOPSIS
    Writes rows to a CSV, always producing a file (matches search-domain.ps1's Export-SafeCsv,
    which wrote a header-only file rather than nothing when a section found no data).
#>
function Export-LegacyCsv {
    param(
        [Parameter(Mandatory)][string]$Path,
        [AllowNull()][object[]]$Data,
        [Parameter(Mandatory)][string]$Label
    )
    $rows = @($Data)
    if ($rows.Count -gt 0) {
        $rows | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Force
    }
    else {
        # Header-only placeholder so downstream Test-Path checks still find the file
        [PSCustomObject]@{ Note = 'No data' } | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Force
    }
    Write-ProgressLine -Label $Label -Count $rows.Count
}

<#
.SYNOPSIS
    Explodes one recipient's pipe-joined, type-prefixed EmailAddresses (e.g. "SMTP:a@b.com|smtp:c@b.com")
    into individual legacy-shaped proxy-address rows, filtered to addresses containing $DomainFilter.
#>
function ConvertTo-LegacyProxyRows {
    param(
        [Parameter(Mandatory)][string]$RecipientType,
        [string]$DisplayName,
        [string]$PrimarySmtpAddress,
        [string]$UserPrincipalName,
        [AllowNull()][string]$EmailAddresses,
        [Parameter(Mandatory)][string]$DomainFilter
    )

    if ([string]::IsNullOrWhiteSpace($EmailAddresses)) { return @() }

    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($addr in ($EmailAddresses -split '\|' | Where-Object { $_ })) {
        if ($addr -notlike "*$DomainFilter*") { continue }

        if ($addr -match '^([^:]+):(.+)$') {
            $addrType    = $Matches[1].ToUpper()
            $addrAddress = $Matches[2]
        }
        else {
            $addrType    = ''
            $addrAddress = $addr
        }

        $rows.Add([PSCustomObject]@{
            RecipientType      = $RecipientType
            DisplayName        = $DisplayName
            PrimarySmtpAddress = $PrimarySmtpAddress
            UserPrincipalName  = $UserPrincipalName
            ProxyAddress       = $addrAddress
            AddressType        = $addrType
            IsPrimary          = ($addrType -ceq 'SMTP')
        })
    }
    return $rows.ToArray()
}

# -----------------------------------------------------------------------
# Public function
# -----------------------------------------------------------------------

<#
.SYNOPSIS
    Writes the 8 search-domain.ps1-compatible CSVs into DiscoveryFolder from the Raw JSON
    already written by this assessment run.
#>
function Export-LegacyDiscoveryCsvs {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][PSCustomObject]$Context,
        [Parameter(Mandatory)][string]$DiscoveryFolder
    )

    Write-SectionHeader 'Domain Removal Compatibility CSVs'
    New-Item -ItemType Directory -Path $DiscoveryFolder -Force | Out-Null

    $rawPath = $Context.RawPath

    # --- 01_AcceptedDomains.csv ---
    $acceptedDomains = Import-AssessmentJson -FileName 'AcceptedDomains.json' -RawPath $rawPath
    $rows01 = @($acceptedDomains | ForEach-Object {
        [PSCustomObject]@{
            Name       = $_.Name
            DomainName = $_.DomainName
            DomainType = $_.DomainType
            IsDefault  = [bool]$_.Default
        }
    })
    Export-LegacyCsv -Path (Join-Path $DiscoveryFolder '01_AcceptedDomains.csv') -Data $rows01 -Label 'Accepted Domains'

    # --- 02_ADUsers.csv: ADUsers.json (LEFT) join UserMailboxes.json on UPN ---
    $adUsers       = Import-AssessmentJson -FileName 'ADUsers.json'        -RawPath $rawPath
    $userMailboxes = Import-AssessmentJson -FileName 'UserMailboxes.json'  -RawPath $rawPath
    $mbByUpn = @{}
    foreach ($mb in $userMailboxes) {
        if ($mb.UserPrincipalName) { $mbByUpn[$mb.UserPrincipalName.ToLower()] = $mb }
    }
    $rows02 = @($adUsers | ForEach-Object {
        $mb = if ($_.UserPrincipalName -and $mbByUpn.ContainsKey($_.UserPrincipalName.ToLower())) { $mbByUpn[$_.UserPrincipalName.ToLower()] } else { $null }
        [PSCustomObject]@{
            DisplayName              = $_.DisplayName
            GivenName                = $_.GivenName
            Surname                  = $_.Surname
            UserPrincipalName        = $_.UserPrincipalName
            sAMAccountName           = $_.SamAccountName
            Mail                     = $_.Mail
            EmployeeID               = $_.EmployeeID
            Department               = $_.Department
            Title                    = $_.Title
            Company                  = $_.Company
            Manager                  = $_.Manager
            Enabled                  = $_.Enabled
            DistinguishedName        = $_.DistinguishedName
            ObjectGUID               = $_.ObjectGUID
            SID                      = $_.SID
            ProxyAddresses           = $_.ProxyAddresses -replace '\|', '; '
            ExtensionAttribute6      = $_.ExtensionAttribute6
            ExtensionAttribute7      = $_.ExtensionAttribute7
            MailboxFound             = ($null -ne $mb)
            RecipientType            = if ($mb) { 'UserMailbox' }         else { $null }
            PrimarySmtpAddress       = if ($mb) { $mb.PrimarySmtpAddress } else { $null }
            Alias                    = if ($mb) { $mb.Alias }             else { $null }
            Database                 = if ($mb) { $mb.Database }          else { $null }
            ItemCount                = if ($mb) { $mb.ItemCount }         else { $null }
            TotalItemSizeGB          = if ($mb) { $mb.TotalItemSizeGB }   else { $null }
            LitigationHoldEnabled    = if ($mb) { $mb.LitigationHoldEnabled } else { $null }
            ArchiveStatus            = if ($mb) { $mb.ArchiveStatus }     else { $null }
            WhenCreated              = if ($mb) { $mb.WhenCreated }       else { $null }
            WhenMailboxCreated       = if ($mb) { $mb.WhenMailboxCreated } else { $null }
        }
    })
    Export-LegacyCsv -Path (Join-Path $DiscoveryFolder '02_ADUsers.csv') -Data $rows02 -Label 'AD Users (Mailboxes)'

    # --- 03_DistributionGroups.csv: DistributionGroups.json + MailEnabledSecurityGroups.json ---
    $dgs   = Import-AssessmentJson -FileName 'DistributionGroups.json'        -RawPath $rawPath
    $mesgs = Import-AssessmentJson -FileName 'MailEnabledSecurityGroups.json' -RawPath $rawPath
    $rows03 = @(@($dgs) + @($mesgs))
    Export-LegacyCsv -Path (Join-Path $DiscoveryFolder '03_DistributionGroups.csv') -Data $rows03 -Label 'Distribution Groups'

    # --- 04_MailContacts.csv ---
    $contacts = Import-AssessmentJson -FileName 'MailContacts.json' -RawPath $rawPath
    Export-LegacyCsv -Path (Join-Path $DiscoveryFolder '04_MailContacts.csv') -Data $contacts -Label 'Mail Contacts'

    # --- 05_SharedMailboxes.csv ---
    $sharedMailboxes = Import-AssessmentJson -FileName 'SharedMailboxes.json' -RawPath $rawPath
    Export-LegacyCsv -Path (Join-Path $DiscoveryFolder '05_SharedMailboxes.csv') -Data $sharedMailboxes -Label 'Shared Mailboxes'

    # --- 06_M365Groups.csv: M365Groups.json + Teams.json (PrimarySmtpAddress backfilled via GroupId join) ---
    # Every VBU-scoped Team is always also a VBU-scoped M365 Group in this architecture (Graph.psm1's
    # Get-TeamData and Get-M365GroupData apply the identical scoping predicate to the same source set),
    # so Teams.json's GroupIds are a subset of M365Groups.json's. Only add a Team row when its GroupId
    # is genuinely missing from M365Groups.json, so a Team-backed group isn't processed twice by the
    # downstream Set-DistributionGroup/Set-UnifiedGroup loops - legacy's single Get-UnifiedGroup call
    # only ever produced one row per group regardless of Team-backing.
    $m365Groups = Import-AssessmentJson -FileName 'M365Groups.json' -RawPath $rawPath
    $teamsJson  = Import-AssessmentJson -FileName 'Teams.json'      -RawPath $rawPath
    $groupIdToEmail = @{}
    $knownGroupIds  = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($g in $m365Groups) {
        if ($g.GroupId -and $g.PrimarySmtpAddress) { $groupIdToEmail[$g.GroupId] = $g.PrimarySmtpAddress }
        if ($g.GroupId) { [void]$knownGroupIds.Add($g.GroupId) }
    }
    $teamRows = @($teamsJson | Where-Object { -not $_.GroupId -or -not $knownGroupIds.Contains($_.GroupId) } | ForEach-Object {
        $email = if ($_.GroupId -and $groupIdToEmail.ContainsKey($_.GroupId)) { $groupIdToEmail[$_.GroupId] } else { $null }
        [PSCustomObject]@{
            GroupId            = $_.GroupId
            DisplayName        = $_.DisplayName
            PrimarySmtpAddress = $email
            Visibility         = $_.Visibility
            CreatedDateTime    = $_.CreatedDateTime
            Owners             = $_.Owners
            Members            = $_.Members
            OwnerCount         = $_.OwnerCount
            MemberCount        = $_.MemberCount
            IsTeam             = $true
        }
    })
    $groupRows = @($m365Groups | ForEach-Object {
        [PSCustomObject]@{
            GroupId            = $_.GroupId
            DisplayName        = $_.DisplayName
            PrimarySmtpAddress = $_.PrimarySmtpAddress
            Visibility         = $_.Visibility
            CreatedDateTime    = $_.CreatedDateTime
            Owners             = $_.Owners
            Members            = $_.Members
            OwnerCount         = $_.OwnerCount
            MemberCount        = $_.MemberCount
            IsTeam             = $false
        }
    })
    $rows06 = @(@($groupRows) + @($teamRows))
    Export-LegacyCsv -Path (Join-Path $DiscoveryFolder '06_M365Groups.csv') -Data $rows06 -Label 'M365 Groups (+ Teams)'

    # --- 12_Devices.csv: IntuneDevices.json only (matches legacy's Graph-only device scan) ---
    $intuneDevices = Import-AssessmentJson -FileName 'IntuneDevices.json' -RawPath $rawPath
    $rows12 = @($intuneDevices | ForEach-Object {
        [PSCustomObject]@{
            OwnerUPN       = $_.PrimaryUserUPN
            OwnerName      = ''
            AccountEnabled = $_.Enabled
            DeviceName     = $_.DeviceName
            DeviceObjectId = $_.EntraObjectId
            EntraDeviceId  = $_.DeviceId
            OS             = $_.OperatingSystem
            OSVersion      = $_.OperatingSystemVersion
            TrustType      = $_.TrustType
            LastSignIn     = $_.ApproximateLastSignInDateTime
            DiscoveredAt   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        }
    })
    Export-LegacyCsv -Path (Join-Path $DiscoveryFolder '12_Devices.csv') -Data $rows12 -Label 'Entra Registered Devices'

    # --- 13_ProxyAddresses.csv: derived by exploding EmailAddresses across every recipient type ---
    $domainFilter    = $Context.VBUDomain
    $resourceMbx     = Import-AssessmentJson -FileName 'ResourceMailboxes.json' -RawPath $rawPath
    $proxyRows       = [System.Collections.Generic.List[object]]::new()

    foreach ($mb in $userMailboxes) {
        $proxyRows.AddRange(@(ConvertTo-LegacyProxyRows -RecipientType 'UserMailbox' -DisplayName $mb.DisplayName `
            -PrimarySmtpAddress $mb.PrimarySmtpAddress -UserPrincipalName $mb.UserPrincipalName `
            -EmailAddresses $mb.EmailAddresses -DomainFilter $domainFilter))
    }
    foreach ($mb in $sharedMailboxes) {
        $proxyRows.AddRange(@(ConvertTo-LegacyProxyRows -RecipientType 'SharedMailbox' -DisplayName $mb.DisplayName `
            -PrimarySmtpAddress $mb.PrimarySmtpAddress -UserPrincipalName $null `
            -EmailAddresses $mb.EmailAddresses -DomainFilter $domainFilter))
    }
    foreach ($mb in $resourceMbx) {
        $proxyRows.AddRange(@(ConvertTo-LegacyProxyRows -RecipientType $mb.ResourceType -DisplayName $mb.DisplayName `
            -PrimarySmtpAddress $mb.PrimarySmtpAddress -UserPrincipalName $null `
            -EmailAddresses $mb.EmailAddresses -DomainFilter $domainFilter))
    }
    foreach ($dg in $rows03) {
        $proxyRows.AddRange(@(ConvertTo-LegacyProxyRows -RecipientType $dg.GroupType -DisplayName $dg.DisplayName `
            -PrimarySmtpAddress $dg.PrimarySmtpAddress -UserPrincipalName $null `
            -EmailAddresses $dg.EmailAddresses -DomainFilter $domainFilter))
    }
    foreach ($c in $contacts) {
        # Mail contacts expose EmailAddresses (proxy list) plus a distinct ExternalEmailAddress
        $proxyRows.AddRange(@(ConvertTo-LegacyProxyRows -RecipientType 'MailContact' -DisplayName $c.DisplayName `
            -PrimarySmtpAddress $c.ExternalEmailAddress -UserPrincipalName $null `
            -EmailAddresses $c.EmailAddresses -DomainFilter $domainFilter))
    }
    foreach ($g in $m365Groups) {
        # ProxyAddresses (raw, semicolon/pipe format from Graph) - only present when Graph.psm1
        # collected it; older assessment data without it just yields no rows here.
        $proxyRows.AddRange(@(ConvertTo-LegacyProxyRows -RecipientType 'GroupMailbox' -DisplayName $g.DisplayName `
            -PrimarySmtpAddress $g.PrimarySmtpAddress -UserPrincipalName $null `
            -EmailAddresses $g.ProxyAddresses -DomainFilter $domainFilter))
    }
    Export-LegacyCsv -Path (Join-Path $DiscoveryFolder '13_ProxyAddresses.csv') -Data $proxyRows.ToArray() -Label 'Proxy Addresses to Remove'

    Write-Host ''
    Write-Host ($PREFIX_OK + "Domain Removal compatibility CSVs written: $DiscoveryFolder") -ForegroundColor Green
}

# -----------------------------------------------------------------------
# Exports
# -----------------------------------------------------------------------

Export-ModuleMember -Function 'Export-LegacyDiscoveryCsvs'

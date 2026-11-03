#Requires -Version 7.0

Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -DisableNameChecking -Force -Global
Import-Module ExchangeOnlineManagement -DisableNameChecking -ErrorAction Stop

$script:EXOMailboxProperties = @(
    'DisplayName', 'UserPrincipalName', 'PrimarySmtpAddress', 'EmailAddresses',
    'Alias', 'ExchangeGuid', 'RecipientTypeDetails', 'Database',
    'LitigationHoldEnabled', 'ArchiveStatus', 'WhenCreated', 'WhenMailboxCreated',
    'ProhibitSendQuota', 'ProhibitSendReceiveQuota', 'ForwardingSmtpAddress',
    'GrantSendOnBehalfTo', 'CustomAttribute6', 'CustomAttribute7'
)

# -----------------------------------------------------------------------
# Private functions
# -----------------------------------------------------------------------

<#
.SYNOPSIS
    Retrieves all VBU-scoped mailboxes, preferring cache over live query.
.DESCRIPTION
    Checks for Mailboxes-All.json cache first (< 7 days old). If found, loads from cache and
    filters in memory. Otherwise, queries Exchange Online live with server-side filtering.

    The live query's server-side -like filter cannot express a word boundary, so it deliberately
    over-includes and a client-side pass then trims the CustomAttribute6 substring false
    positives - e.g. a mailbox tagged 'Ceramics' surfacing for the search term 'amic'.
    The trim only ever removes CustomAttribute6-only hits: a mailbox that also qualifies
    via CustomAttribute7 (exact VBU ID) or whose UPN/addresses carry the VBU domain is
    always kept, as is one with a blank CustomAttribute6.

    Cache includes mailbox statistics and permissions (FullAccess, SendAs) which are returned
    directly without additional queries.
#>
function Get-AllMailboxesRaw {
    param([PSCustomObject]$Context)

    $t      = $Context.VBUSearchTerm
    $d      = $Context.VBUDomain
    $i      = $Context.VBUId
    $mbCache = if ($Context.CacheFolder) { Join-Path $Context.CacheFolder 'Mailboxes-All.json' } else { '' }

    # Try cache first
    if ($mbCache -and (Test-Path $mbCache)) {
        $cacheAge = (Get-Date) - (Get-Item $mbCache).LastWriteTime
        if ($cacheAge.TotalDays -lt 7) {
            Write-Host (Get-TimestampPrefix ($PREFIX_INFO + "Loading mailboxes from cache ($([math]::Round($cacheAge.TotalDays, 1)) days old)...")) -ForegroundColor DarkGray

            $allMailboxes = @(Get-Content $mbCache -Raw | ConvertFrom-Json)
            $kept = @($allMailboxes | Where-Object {
                $ca6 = "$($_.CustomAttribute6)"
                $addresses = @($_.EmailAddresses -split '\|') + @($_.UserPrincipalName) + @($_.PrimarySmtpAddress)

                # Same filtering logic as live query
                $ca6Ok = [string]::IsNullOrWhiteSpace($ca6) -or (Test-WholeWordMatch -Text $ca6 -Term $t)
                $byId  = (-not [string]::IsNullOrWhiteSpace($i)) -and ("$($_.CustomAttribute7)" -eq $i)
                $byDomain = @($addresses | Where-Object { $_ -and $_ -like "*$d*" }).Count -gt 0

                $ca6Ok -or $byId -or $byDomain
            })

            Write-Host (Get-TimestampPrefix ($PREFIX_OK + "Filtered to $($kept.Count) mailboxes from cache")) -ForegroundColor Green
            return $kept
        }
    }

    # No cache or cache stale - query live
    Write-Host (Get-TimestampPrefix ($PREFIX_INFO + 'Querying mailboxes live (no cache)...')) -ForegroundColor DarkGray
    Write-Host (Get-TimestampPrefix ($PREFIX_WARN + 'This can take 10-30 minutes on large tenants - please wait...')) -ForegroundColor Yellow
    Write-Host (Get-TimestampPrefix ($PREFIX_INFO + 'TIP: Run Update-MailboxesCache.ps1 to create a cache for faster future assessments')) -ForegroundColor DarkGray
    $filter = "(CustomAttribute6 -like '*$t*') -or (CustomAttribute7 -eq '$i') -or " +
              "(UserPrincipalName -like '*$t*')"

    $queryStart = Get-Date
    $raw = @(Get-EXOMailbox -Filter $filter `
        -RecipientTypeDetails UserMailbox, SharedMailbox, RoomMailbox, EquipmentMailbox `
        -ResultSize Unlimited `
        -Properties $script:EXOMailboxProperties `
        -ErrorAction Stop)
    $queryDuration = (Get-Date) - $queryStart

    Write-Host (Get-TimestampPrefix ($PREFIX_OK + "Retrieved $($raw.Count) mailboxes from Exchange Online (took $([math]::Round($queryDuration.TotalMinutes, 1)) min)")) -ForegroundColor Green

    $kept = @($raw | Where-Object {
        $ca6 = "$($_.CustomAttribute6)"

        # Blank CustomAttribute6 was never a search-term hit; a whole-word hit is genuine.
        $ca6Ok = [string]::IsNullOrWhiteSpace($ca6) -or (Test-WholeWordMatch -Text $ca6 -Term $t)

        # Independent signals - either one keeps the mailbox regardless of CustomAttribute6.
        $addresses = @($_.EmailAddresses) + @($_.UserPrincipalName) + @($_.PrimarySmtpAddress)
        $byId      = (-not [string]::IsNullOrWhiteSpace($i)) -and ("$($_.CustomAttribute7)" -eq $i)
        $byDomain  = @($addresses | Where-Object { $_ -and $_ -like "*$d*" }).Count -gt 0

        $ca6Ok -or $byId -or $byDomain
    })

    $dropped = $raw.Count - $kept.Count
    if ($dropped -gt 0) {
        Write-Host ($PREFIX_INFO + "Trimmed $dropped mailbox(es) matching '$t' only as a substring") -ForegroundColor DarkGray
    }

    return $kept
}

<#
.SYNOPSIS
    Builds output records for one recipient type from raw mailbox data, fetching per-mailbox statistics for User and Shared types.
#>
function Build-MailboxRecords {
    param(
        [object[]]$RawMailboxes,
        [string]$RecipientType  # UserMailbox | SharedMailbox | RoomMailbox | EquipmentMailbox
    )

    $needStats = $RecipientType -in @('UserMailbox', 'SharedMailbox')
    $records   = [System.Collections.Generic.List[PSCustomObject]]::new()
    $mailboxes = @($RawMailboxes | Where-Object { $_.RecipientTypeDetails -eq $RecipientType })

    if ($needStats -and $mailboxes.Count -gt 0) {
        Write-Host (Get-TimestampPrefix ($PREFIX_INFO + "Getting statistics for $($mailboxes.Count) $RecipientType mailboxes...")) -ForegroundColor DarkGray
    }

    $processed = 0
    foreach ($mb in $mailboxes) {
        $processed++
        if ($needStats -and $processed % 100 -eq 0) {
            $pct = [math]::Round(($processed / $mailboxes.Count) * 100)
            Write-Host (Get-TimestampPrefix ($PREFIX_INFO + "  $processed / $($mailboxes.Count) ($pct%)")) -ForegroundColor DarkGray
        }
        $itemCount    = $null
        $sizeGB       = [double]0
        $sizeBytes    = [long]0
        $lastLogon    = $null

        # Check if statistics are already present (from cache)
        if ($mb.PSObject.Properties['ItemCount'] -and $null -ne $mb.ItemCount) {
            # Using cached statistics
            $itemCount = $mb.ItemCount
            $sizeGB    = $mb.SizeGB
            $sizeBytes = $mb.SizeBytes
            $lastLogon = $mb.LastLogon
        }
        elseif ($needStats) {
            # No cached stats - query live
            try {
                $stats     = Get-EXOMailboxStatistics -Identity $mb.ExchangeGuid.ToString() -ErrorAction Stop
                $itemCount = $stats.ItemCount
                $sizeGB    = Convert-SizeToGB -Value $stats.TotalItemSize
                $sizeBytes = Convert-SizeToBytes -Value $stats.TotalItemSize
                $lastLogon = if ($stats.PSObject.Properties['LastLogonTime']) {
                    $stats.LastLogonTime
                } elseif ($stats.PSObject.Properties['LastUserActionTime']) {
                    $stats.LastUserActionTime
                } else { $null }
            }
            catch { }
        }

        # Handle EmailAddresses - could be pipe-delimited string (from cache) or array (from live query)
        $emailAddresses = if ($mb.EmailAddresses -is [string]) {
            $mb.EmailAddresses
        } else {
            ($mb.EmailAddresses | Sort-Object) -join '|'
        }

        if ($RecipientType -eq 'UserMailbox') {
            $records.Add([PSCustomObject]@{
                DisplayName                  = $mb.DisplayName
                UserPrincipalName            = $mb.UserPrincipalName
                PrimarySmtpAddress           = $mb.PrimarySmtpAddress
                EmailAddresses               = $emailAddresses
                Alias                        = $mb.Alias
                ExchangeGuid                 = $mb.ExchangeGuid.ToString()
                Database                     = $mb.Database
                ItemCount                    = $itemCount
                TotalItemSizeGB              = $sizeGB
                TotalItemSizeBytes           = $sizeBytes
                ProhibitSendQuotaGB          = Convert-SizeToGB -Value $mb.ProhibitSendQuota
                ProhibitSendReceiveQuotaGB   = Convert-SizeToGB -Value $mb.ProhibitSendReceiveQuota
                LitigationHoldEnabled        = $mb.LitigationHoldEnabled
                ArchiveStatus                = $mb.ArchiveStatus
                LastLogonTime                = $lastLogon
                WhenCreated                  = $mb.WhenCreated
                WhenMailboxCreated           = $mb.WhenMailboxCreated
                ForwardingSmtpAddress        = if ($mb.ForwardingSmtpAddress) { $mb.ForwardingSmtpAddress.ToString() } else { '' }
            })
        }
        elseif ($RecipientType -eq 'SharedMailbox') {
            $records.Add([PSCustomObject]@{
                DisplayName                  = $mb.DisplayName
                PrimarySmtpAddress           = $mb.PrimarySmtpAddress
                EmailAddresses               = $emailAddresses
                Alias                        = $mb.Alias
                ExchangeGuid                 = $mb.ExchangeGuid.ToString()
                ItemCount                    = $itemCount
                TotalItemSizeGB              = $sizeGB
                TotalItemSizeBytes           = $sizeBytes
                ProhibitSendQuotaGB          = Convert-SizeToGB -Value $mb.ProhibitSendQuota
                ProhibitSendReceiveQuotaGB   = Convert-SizeToGB -Value $mb.ProhibitSendReceiveQuota
                LitigationHoldEnabled        = $mb.LitigationHoldEnabled
                ArchiveStatus                = $mb.ArchiveStatus
                WhenCreated                  = $mb.WhenCreated
                WhenMailboxCreated           = $mb.WhenMailboxCreated
                ForwardingSmtpAddress        = if ($mb.ForwardingSmtpAddress) { $mb.ForwardingSmtpAddress.ToString() } else { '' }
            })
        }
        else {
            # Room or Equipment -> Resource Mailboxes
            $resourceType = if ($RecipientType -eq 'RoomMailbox') { 'Room' } else { 'Equipment' }
            $records.Add([PSCustomObject]@{
                DisplayName            = $mb.DisplayName
                PrimarySmtpAddress     = $mb.PrimarySmtpAddress
                EmailAddresses         = $emailAddresses
                Alias                  = $mb.Alias
                ResourceType           = $resourceType
                ExchangeGuid           = $mb.ExchangeGuid.ToString()
                WhenCreated            = $mb.WhenCreated
                WhenMailboxCreated     = $mb.WhenMailboxCreated
                ForwardingSmtpAddress  = if ($mb.ForwardingSmtpAddress) { $mb.ForwardingSmtpAddress.ToString() } else { '' }
            })
        }
    }

    return $records.ToArray()
}

<#
.SYNOPSIS
    Retrieves mail contacts matching the VBU search term via server-side EmailAddresses filter.
#>
function Get-MailContactData {
    param([PSCustomObject]$Context)

    $t = $Context.VBUSearchTerm
    return @(Get-MailContact -Filter "EmailAddresses -like '*$t*'" -ResultSize Unlimited -ErrorAction Stop |
        ForEach-Object {
            [PSCustomObject]@{
                DisplayName          = $_.DisplayName
                ExternalEmailAddress = if ($_.ExternalEmailAddress) { $_.ExternalEmailAddress.ToString() } else { '' }
                EmailAddresses       = if ($_.EmailAddresses) { ($_.EmailAddresses | Sort-Object) -join '|' } else { '' }
                Alias                = $_.Alias
                WhenCreated          = $_.WhenCreated
            }
        })
}

<#
.SYNOPSIS
    Retrieves VBU-scoped distribution groups - from cache if available, otherwise via live query.
.DESCRIPTION
    Checks for cached DistributionGroups-All.json and MailEnabledSecurityGroups-All.json (< 7 days old).
    If found, loads from cache and filters by VBU search term. Otherwise falls back to the slower
    Get-DistributionGroup -Filter query with live member counts.
#>
function Get-DistributionGroupData {
    param([PSCustomObject]$Context)

    $t         = $Context.VBUSearchTerm
    $dgCache   = if ($Context.CacheFolder) { Join-Path $Context.CacheFolder 'DistributionGroups-All.json' } else { '' }
    $mesgCache = if ($Context.CacheFolder) { Join-Path $Context.CacheFolder 'MailEnabledSecurityGroups-All.json' } else { '' }
    $dgs       = [System.Collections.Generic.List[PSCustomObject]]::new()
    $mesg      = [System.Collections.Generic.List[PSCustomObject]]::new()

    # Try cache first
    if ($dgCache -and (Test-Path $dgCache) -and (Test-Path $mesgCache)) {
        $dgAge = (Get-Date) - (Get-Item $dgCache).LastWriteTime
        if ($dgAge.TotalDays -lt 7) {
            Write-Host ($PREFIX_INFO + "Loading distribution groups from cache ($([math]::Round($dgAge.TotalDays, 1)) days old)...") -ForegroundColor DarkGray

            # Load DGs from cache and filter
            $allDGs = @(Get-Content $dgCache -Raw | ConvertFrom-Json)
            foreach ($dg in $allDGs) {
                if ($dg.EmailAddresses -like "*$t*") {
                    $dgs.Add($dg)
                }
            }

            # Load MESGs from cache and filter
            $allMESG = @(Get-Content $mesgCache -Raw | ConvertFrom-Json)
            foreach ($m in $allMESG) {
                if ($m.EmailAddresses -like "*$t*") {
                    $mesg.Add($m)
                }
            }

            return @{ DGs = $dgs.ToArray(); MailEnabledSecurityGroups = $mesg.ToArray() }
        }
    }

    # No cache or cache stale - query live
    Write-Host ($PREFIX_INFO + 'Querying distribution groups live (no cache)...') -ForegroundColor DarkGray
    $raw  = @(Get-DistributionGroup -Filter "EmailAddresses -like '*$t*'" -ResultSize Unlimited -ErrorAction Stop)

    foreach ($dg in $raw) {
        try {
            $memberCount = (Get-DistributionGroupMember -Identity $dg.Guid.ToString() -ResultSize Unlimited -ErrorAction Stop).Count
        }
        catch { $memberCount = $null }

        $record = [PSCustomObject]@{
            DisplayName                  = $dg.DisplayName
            PrimarySmtpAddress           = if ($dg.PrimarySmtpAddress) { $dg.PrimarySmtpAddress.ToString() } else { '' }
            EmailAddresses               = if ($dg.EmailAddresses) { ($dg.EmailAddresses | Sort-Object) -join '|' } else { '' }
            Alias                        = $dg.Alias
            GroupType                    = $dg.RecipientTypeDetails.ToString()
            MemberCount                  = $memberCount
            ManagedBy                    = if ($dg.ManagedBy) { ($dg.ManagedBy | ForEach-Object { $_.ToString() }) -join '|' } else { '' }
            HiddenFromAddressListsEnabled = $dg.HiddenFromAddressListsEnabled
            WhenCreated                  = $dg.WhenCreated
        }

        if ($dg.RecipientTypeDetails -eq 'MailUniversalSecurityGroup') {
            $mesg.Add($record)
        }
        else {
            $dgs.Add($record)
        }
    }

    return @{ DGs = $dgs.ToArray(); MailEnabledSecurityGroups = $mesg.ToArray() }
}

<#
.SYNOPSIS
    Retrieves dynamic distribution groups matching the VBU search term via server-side EmailAddresses filter.
#>
function Get-DynamicDistributionGroupData {
    param([PSCustomObject]$Context)

    $t = $Context.VBUSearchTerm
    return @(Get-DynamicDistributionGroup -Filter "EmailAddresses -like '*$t*'" -ResultSize Unlimited -ErrorAction Stop |
        ForEach-Object {
            [PSCustomObject]@{
                DisplayName        = $_.DisplayName
                PrimarySmtpAddress = if ($_.PrimarySmtpAddress) { $_.PrimarySmtpAddress.ToString() } else { '' }
                EmailAddresses     = if ($_.EmailAddresses) { ($_.EmailAddresses | Sort-Object) -join '|' } else { '' }
                Alias              = $_.Alias
                RecipientFilter    = $_.RecipientFilter
                ManagedBy          = if ($_.ManagedBy) { ($_.ManagedBy | ForEach-Object { $_.ToString() }) -join '|' } else { '' }
                WhenCreated        = $_.WhenCreated
            }
        })
}

<#
.SYNOPSIS
    Collects FullAccess, SendAs, and SendOnBehalf delegation across user, shared, and resource mailboxes, one row per grant.
#>
function Resolve-DelegateIdentity {
    param([string]$Value, [hashtable]$Cache)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $Value }
    if ($Cache.ContainsKey($Value)) { return $Cache[$Value] }
    $resolved = $Value
    try {
        $rcpt = Get-EXORecipient -Identity $Value -ErrorAction Stop
        if ($rcpt.PrimarySmtpAddress) { $resolved = $rcpt.PrimarySmtpAddress.ToString() }
        elseif ($rcpt.DisplayName)    { $resolved = $rcpt.DisplayName }
    }
    catch { }
    $Cache[$Value] = $resolved
    return $resolved
}

function Get-MailboxPermissionData {
    param([object[]]$Mailboxes)

    $records = [System.Collections.Generic.List[PSCustomObject]]::new()
    $delegateCache = @{}

    foreach ($mb in $Mailboxes) {
        $mailboxType = $mb.RecipientTypeDetails.ToString()
        $mailboxSmtp = $mb.PrimarySmtpAddress.ToString()
        $mailboxUpn  = if ($mailboxType -eq 'UserMailbox') { $mb.UserPrincipalName } else { '' }

        try {
            $fullAccess = Get-EXOMailboxPermission -Identity $mb.ExchangeGuid.ToString() -ErrorAction Stop |
                Where-Object { -not $_.IsInherited -and $_.User -ne 'NT AUTHORITY\SELF' -and $_.User -notlike 'S-1-5-*' }
            foreach ($p in $fullAccess) {
                $records.Add([PSCustomObject]@{
                    Mailbox        = $mailboxSmtp
                    MailboxUPN     = $mailboxUpn
                    MailboxType    = $mailboxType
                    Delegate       = Resolve-DelegateIdentity -Value "$($p.User)" -Cache $delegateCache
                    PermissionType = 'FullAccess'
                    AccessRights   = ($p.AccessRights | ForEach-Object { $_.ToString() }) -join '|'
                })
            }
        }
        catch {
            Write-Host ($PREFIX_WARN + "Could not get FullAccess permissions for '$($mb.DisplayName)': " + $_.Exception.Message) -ForegroundColor Yellow
        }

        try {
            $sendAs = Get-EXORecipientPermission -Identity $mb.ExchangeGuid.ToString() -ErrorAction Stop |
                Where-Object { -not $_.IsInherited -and $_.AccessRights -contains 'SendAs' -and $_.Trustee -ne 'NT AUTHORITY\SELF' }
            foreach ($p in $sendAs) {
                $records.Add([PSCustomObject]@{
                    Mailbox        = $mailboxSmtp
                    MailboxUPN     = $mailboxUpn
                    MailboxType    = $mailboxType
                    Delegate       = Resolve-DelegateIdentity -Value "$($p.Trustee)" -Cache $delegateCache
                    PermissionType = 'SendAs'
                    AccessRights   = ($p.AccessRights | ForEach-Object { $_.ToString() }) -join '|'
                })
            }
        }
        catch {
            Write-Host ($PREFIX_WARN + "Could not get SendAs permissions for '$($mb.DisplayName)': " + $_.Exception.Message) -ForegroundColor Yellow
        }

        foreach ($delegate in @($mb.GrantSendOnBehalfTo)) {
            if ($null -eq $delegate -or "$delegate" -eq '') { continue }
            $records.Add([PSCustomObject]@{
                Mailbox        = $mailboxSmtp
                MailboxUPN     = $mailboxUpn
                MailboxType    = $mailboxType
                Delegate       = Resolve-DelegateIdentity -Value "$delegate" -Cache $delegateCache
                PermissionType = 'SendOnBehalf'
                AccessRights   = ''
            })
        }
    }

    return $records.ToArray()
}

<#
.SYNOPSIS
    Retrieves tenant-wide transport rules.
#>
function Get-TransportRuleData {
    return @(Get-TransportRule -ErrorAction Stop |
        ForEach-Object {
            [PSCustomObject]@{
                Name        = $_.Name
                State       = $_.State.ToString()
                Priority    = $_.Priority
                Description = $_.Description
                WhenChanged = $_.WhenChanged
            }
        })
}

<#
.SYNOPSIS
    Retrieves tenant-wide inbound and outbound connectors as a combined list.
#>
function Get-ConnectorData {
    $records = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($c in @(Get-InboundConnector -ErrorAction Stop)) {
        $records.Add([PSCustomObject]@{
            Name                     = $c.Name
            ConnectorType            = 'Inbound'
            Enabled                  = $c.Enabled
            TlsSenderCertificateName = $c.TlsSenderCertificateName
            ConnectorSource          = $c.ConnectorSource.ToString()
            WhenChanged              = $c.WhenChanged
        })
    }

    foreach ($c in @(Get-OutboundConnector -ErrorAction Stop)) {
        $records.Add([PSCustomObject]@{
            Name                     = $c.Name
            ConnectorType            = 'Outbound'
            Enabled                  = $c.Enabled
            TlsSenderCertificateName = $null
            ConnectorSource          = $c.ConnectorSource.ToString()
            WhenChanged              = $c.WhenChanged
        })
    }

    return $records.ToArray()
}

<#
.SYNOPSIS
    Retrieves tenant-wide accepted domains.
#>
function Get-AcceptedDomainData {
    return @(Get-AcceptedDomain -ErrorAction Stop |
        ForEach-Object {
            [PSCustomObject]@{
                Name               = $_.Name
                DomainName         = $_.DomainName
                DomainType         = $_.DomainType.ToString()
                Default            = $_.Default
                AuthenticationType = $_.AuthenticationType
            }
        })
}

<#
.SYNOPSIS
    Retrieves tenant-wide remote domains.
#>
function Get-RemoteDomainData {
    return @(Get-RemoteDomain -ErrorAction Stop |
        ForEach-Object {
            [PSCustomObject]@{
                Name                = $_.Name
                DomainName          = $_.DomainName
                AllowedOOFType      = $_.AllowedOOFType.ToString()
                AutoReplyEnabled    = $_.AutoReplyEnabled
                AutoForwardEnabled  = $_.AutoForwardEnabled
                WhenChanged         = $_.WhenChanged
            }
        })
}

<#
.SYNOPSIS
    Retrieves tenant-wide journal rules.
#>
function Get-JournalRuleData {
    return @(Get-JournalRule -ErrorAction Stop |
        ForEach-Object {
            [PSCustomObject]@{
                Name               = $_.Name
                JournalEmailAddress = $_.JournalEmailAddress
                Scope              = $_.Scope.ToString()
                Enabled            = $_.Enabled
                Recipient          = $_.Recipient
            }
        })
}

# -----------------------------------------------------------------------
# Public functions
# -----------------------------------------------------------------------

<#
.SYNOPSIS
    Orchestrates all Exchange Online collection workloads.
.DESCRIPTION
    Uses the session connected by the orchestrator to collect mailboxes, contacts,
    distribution groups, permissions, and tenant-wide config objects, writing thirteen
    JSON files to the Raw folder. Failures are non-critical: the error is recorded in
    CollectorStatus.json and a failed collector result is returned.
#>
function Invoke-ExchangeCollection {
    [CmdletBinding()]
    param([Parameter(Mandatory)][PSCustomObject]$Context)

    $start = Get-Date
    Write-SectionHeader 'Exchange Online'

    try {
        Write-Host ($PREFIX_INFO + 'Querying mailboxes...') -ForegroundColor DarkGray
        $allRaw = Get-AllMailboxesRaw -Context $Context

        $userMailboxes     = Build-MailboxRecords -RawMailboxes $allRaw -RecipientType 'UserMailbox'
        $sharedMailboxes   = Build-MailboxRecords -RawMailboxes $allRaw -RecipientType 'SharedMailbox'
        $roomMailboxes     = Build-MailboxRecords -RawMailboxes $allRaw -RecipientType 'RoomMailbox'
        $equipMailboxes    = Build-MailboxRecords -RawMailboxes $allRaw -RecipientType 'EquipmentMailbox'
        $resourceMailboxes = @($roomMailboxes) + @($equipMailboxes)

        Write-ProgressLine -Label 'User Mailboxes'     -Count $userMailboxes.Count
        Write-ProgressLine -Label 'Shared Mailboxes'   -Count $sharedMailboxes.Count
        Write-ProgressLine -Label 'Resource Mailboxes' -Count $resourceMailboxes.Count

        Write-Host ($PREFIX_INFO + 'Collecting mail contacts...') -ForegroundColor DarkGray
        $contacts = Get-MailContactData -Context $Context
        Write-ProgressLine -Label 'Mail Contacts' -Count $contacts.Count

        Write-Host ($PREFIX_INFO + 'Collecting distribution groups...') -ForegroundColor DarkGray
        $dgResult = Get-DistributionGroupData -Context $Context
        Write-ProgressLine -Label 'Distribution Groups'          -Count $dgResult.DGs.Count
        Write-ProgressLine -Label 'Mail-Enabled Security Groups' -Count $dgResult.MailEnabledSecurityGroups.Count

        Write-Host ($PREFIX_INFO + 'Collecting dynamic distribution groups...') -ForegroundColor DarkGray
        $ddgs = Get-DynamicDistributionGroupData -Context $Context
        Write-ProgressLine -Label 'Dynamic Distribution Groups' -Count $ddgs.Count

        Write-Host ($PREFIX_INFO + 'Collecting mailbox permissions...') -ForegroundColor DarkGray
        $permissions = Get-MailboxPermissionData -Mailboxes $allRaw
        Write-ProgressLine -Label 'Mailbox Permissions' -Count $permissions.Count

        Write-Host ($PREFIX_INFO + 'Collecting transport rules...') -ForegroundColor DarkGray
        $transportRules = Get-TransportRuleData
        Write-ProgressLine -Label 'Transport Rules' -Count $transportRules.Count

        Write-Host ($PREFIX_INFO + 'Collecting connectors...') -ForegroundColor DarkGray
        $connectors = Get-ConnectorData
        Write-ProgressLine -Label 'Connectors' -Count $connectors.Count

        Write-Host ($PREFIX_INFO + 'Collecting accepted domains...') -ForegroundColor DarkGray
        $acceptedDomains = Get-AcceptedDomainData
        Write-ProgressLine -Label 'Accepted Domains' -Count $acceptedDomains.Count

        Write-Host ($PREFIX_INFO + 'Collecting remote domains...') -ForegroundColor DarkGray
        $remoteDomains = Get-RemoteDomainData
        Write-ProgressLine -Label 'Remote Domains' -Count $remoteDomains.Count

        Write-Host ($PREFIX_INFO + 'Collecting journal rules...') -ForegroundColor DarkGray
        $journalRules = Get-JournalRuleData
        Write-ProgressLine -Label 'Journal Rules' -Count $journalRules.Count

        Write-JsonOutput -FileName 'UserMailboxes.json'              -Data $userMailboxes                   -RawPath $Context.RawPath
        Write-JsonOutput -FileName 'SharedMailboxes.json'            -Data $sharedMailboxes                 -RawPath $Context.RawPath
        Write-JsonOutput -FileName 'ResourceMailboxes.json'          -Data $resourceMailboxes               -RawPath $Context.RawPath
        Write-JsonOutput -FileName 'MailContacts.json'               -Data $contacts                        -RawPath $Context.RawPath
        Write-JsonOutput -FileName 'DistributionGroups.json'         -Data $dgResult.DGs                   -RawPath $Context.RawPath
        Write-JsonOutput -FileName 'MailEnabledSecurityGroups.json'  -Data $dgResult.MailEnabledSecurityGroups -RawPath $Context.RawPath
        Write-JsonOutput -FileName 'DynamicDistributionGroups.json'  -Data $ddgs                            -RawPath $Context.RawPath
        Write-JsonOutput -FileName 'MailboxPermissions.json'         -Data $permissions                     -RawPath $Context.RawPath
        Write-JsonOutput -FileName 'TransportRules.json'             -Data $transportRules                  -RawPath $Context.RawPath
        Write-JsonOutput -FileName 'Connectors.json'                 -Data $connectors                      -RawPath $Context.RawPath
        Write-JsonOutput -FileName 'AcceptedDomains.json'            -Data $acceptedDomains                 -RawPath $Context.RawPath
        Write-JsonOutput -FileName 'RemoteDomains.json'              -Data $remoteDomains                   -RawPath $Context.RawPath
        Write-JsonOutput -FileName 'JournalRules.json'               -Data $journalRules                    -RawPath $Context.RawPath

        $counts = @{
            UserMailboxCount               = $userMailboxes.Count
            SharedMailboxCount             = $sharedMailboxes.Count
            ResourceMailboxCount           = $resourceMailboxes.Count
            MailContactCount               = $contacts.Count
            DGCount                        = $dgResult.DGs.Count
            MailEnabledSecurityGroupCount  = $dgResult.MailEnabledSecurityGroups.Count
            DDGCount                       = $ddgs.Count
            MailboxPermissionCount         = $permissions.Count
            TransportRuleCount             = $transportRules.Count
            ConnectorCount                 = $connectors.Count
            AcceptedDomainCount            = $acceptedDomains.Count
            RemoteDomainCount              = $remoteDomains.Count
            JournalRuleCount               = $journalRules.Count
        }
        $msg = "User: $($counts.UserMailboxCount), Shared: $($counts.SharedMailboxCount), " +
               "Resource: $($counts.ResourceMailboxCount), DG: $($counts.DGCount), " +
               "MESG: $($counts.MailEnabledSecurityGroupCount), DDG: $($counts.DDGCount)"
        Update-CollectorStatus -CollectorName 'Exchange Data' -Status 'Complete' `
            -RawPath $Context.RawPath -StartTime $start -Message $msg

        return New-CollectorResult -Success $true -Counts $counts
    }
    catch {
        Write-Host ($PREFIX_FAIL + 'Exchange collection failed: ' + $_.Exception.Message) -ForegroundColor Red
        Update-CollectorStatus -CollectorName 'Exchange Data' -Status 'Failed' `
            -RawPath $Context.RawPath -StartTime $start -Message $_.Exception.Message
        return New-CollectorResult -Success $false -ErrorMessage $_.Exception.Message
    }
}

<#
.SYNOPSIS
    Updates the tenant-wide distribution groups cache file.
.DESCRIPTION
    Runs Get-DistributionGroup -ResultSize Unlimited to collect all distribution groups and
    mail-enabled security groups, then saves them to a cache file that individual VBU assessments
    can read and filter locally instead of running the slow server-side filter every time.
#>
function Update-DistributionGroupsCacheFile {
    param(
        [Parameter(Mandatory)]
        [string]$CacheFolder
    )

    Write-Host ($PREFIX_INFO + 'Collecting all distribution groups...') -ForegroundColor DarkGray
    $allGroups = @(Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop)

    $dgs  = [System.Collections.Generic.List[PSCustomObject]]::new()
    $mesg = [System.Collections.Generic.List[PSCustomObject]]::new()

    $processed = 0
    foreach ($dg in $allGroups) {
        $processed++
        if ($processed % 100 -eq 0) {
            Write-Host ($PREFIX_INFO + "  ...$processed / $($allGroups.Count) groups") -ForegroundColor DarkGray
        }

        try {
            $memberCount = (Get-DistributionGroupMember -Identity $dg.Guid.ToString() -ResultSize Unlimited -ErrorAction Stop).Count
        }
        catch { $memberCount = $null }

        $record = [PSCustomObject]@{
            DisplayName                  = $dg.DisplayName
            PrimarySmtpAddress           = if ($dg.PrimarySmtpAddress) { $dg.PrimarySmtpAddress.ToString() } else { '' }
            EmailAddresses               = if ($dg.EmailAddresses) { ($dg.EmailAddresses | Sort-Object) -join '|' } else { '' }
            Alias                        = $dg.Alias
            GroupType                    = $dg.RecipientTypeDetails.ToString()
            MemberCount                  = $memberCount
            ManagedBy                    = if ($dg.ManagedBy) { ($dg.ManagedBy | ForEach-Object { $_.ToString() }) -join '|' } else { '' }
            HiddenFromAddressListsEnabled = $dg.HiddenFromAddressListsEnabled
            WhenCreated                  = $dg.WhenCreated
            Guid                         = $dg.Guid.ToString()
        }

        if ($dg.RecipientTypeDetails -eq 'MailUniversalSecurityGroup') {
            $mesg.Add($record)
        }
        else {
            $dgs.Add($record)
        }
    }

    Write-Host ($PREFIX_INFO + 'Collecting all dynamic distribution groups...') -ForegroundColor DarkGray
    $ddgs = @(Get-DynamicDistributionGroup -ResultSize Unlimited -ErrorAction Stop |
        ForEach-Object {
            [PSCustomObject]@{
                DisplayName        = $_.DisplayName
                PrimarySmtpAddress = if ($_.PrimarySmtpAddress) { $_.PrimarySmtpAddress.ToString() } else { '' }
                EmailAddresses     = if ($_.EmailAddresses) { ($_.EmailAddresses | Sort-Object) -join '|' } else { '' }
                Alias              = $_.Alias
                RecipientFilter    = $_.RecipientFilter
                ManagedBy          = if ($_.ManagedBy) { ($_.ManagedBy | ForEach-Object { $_.ToString() }) -join '|' } else { '' }
                WhenCreated        = $_.WhenCreated
                Guid               = $_.Guid.ToString()
            }
        })

    # Save each type to its own cache file
    $dgFile = Join-Path $CacheFolder 'DistributionGroups-All.json'
    $dgs.ToArray() | ConvertTo-Json -Depth 10 | Set-Content -Path $dgFile -Encoding UTF8
    Write-Host ($PREFIX_OK + "Cached $($dgs.Count) distribution groups") -ForegroundColor Green

    $mesgFile = Join-Path $CacheFolder 'MailEnabledSecurityGroups-All.json'
    $mesg.ToArray() | ConvertTo-Json -Depth 10 | Set-Content -Path $mesgFile -Encoding UTF8
    Write-Host ($PREFIX_OK + "Cached $($mesg.Count) mail-enabled security groups") -ForegroundColor Green

    $ddgFile = Join-Path $CacheFolder 'DynamicDistributionGroups-All.json'
    $ddgs | ConvertTo-Json -Depth 10 | Set-Content -Path $ddgFile -Encoding UTF8
    Write-Host ($PREFIX_OK + "Cached $($ddgs.Length) dynamic distribution groups") -ForegroundColor Green
}

<#
.SYNOPSIS
    Refreshes the tenant-wide mailbox cache including statistics and permissions.
.DESCRIPTION
    Caches:
    - All mailboxes (User, Shared, Room, Equipment) with properties
    - Statistics for User and Shared mailboxes (item count, size, last logon)
    - FullAccess permissions (who can open each mailbox)
    - SendAs permissions (who can send as each mailbox)

    Creates:
    - Mailboxes-All.json - array of mailbox objects with embedded statistics and permissions
#>
function Update-MailboxesCacheFile {
    param(
        [Parameter(Mandatory)]
        [string]$CacheFolder
    )

    # Step 1: Get all mailboxes
    Write-Host ($PREFIX_INFO + 'Collecting all mailboxes...') -ForegroundColor DarkGray
    $allMailboxes = @(Get-EXOMailbox `
        -RecipientTypeDetails UserMailbox, SharedMailbox, RoomMailbox, EquipmentMailbox `
        -ResultSize Unlimited `
        -Properties $script:EXOMailboxProperties `
        -ErrorAction Stop)
    Write-Host ($PREFIX_OK + "Found $($allMailboxes.Count) mailboxes") -ForegroundColor Green

    # Step 2: Get statistics for User and Shared mailboxes
    Write-Host ($PREFIX_INFO + 'Collecting mailbox statistics (User + Shared only)...') -ForegroundColor DarkGray
    $statsHash = @{}
    $needsStats = @($allMailboxes | Where-Object { $_.RecipientTypeDetails -in @('UserMailbox', 'SharedMailbox') })

    $processed = 0
    foreach ($mb in $needsStats) {
        $processed++
        if ($processed % 100 -eq 0) {
            Write-Host ($PREFIX_INFO + "  ...$processed / $($needsStats.Count) mailboxes") -ForegroundColor DarkGray
        }

        try {
            $stats = Get-EXOMailboxStatistics -Identity $mb.ExchangeGuid.ToString() -ErrorAction Stop
            $statsHash[$mb.ExchangeGuid.ToString()] = @{
                ItemCount   = $stats.ItemCount
                SizeGB      = Convert-SizeToGB -Value $stats.TotalItemSize
                SizeBytes   = Convert-SizeToBytes -Value $stats.TotalItemSize
                LastLogon   = if ($stats.PSObject.Properties['LastLogonTime']) {
                    $stats.LastLogonTime
                } elseif ($stats.PSObject.Properties['LastUserActionTime']) {
                    $stats.LastUserActionTime
                } else { $null }
            }
        }
        catch {
            # Stats unavailable - leave empty
        }
    }
    Write-Host ($PREFIX_OK + "Retrieved statistics for $($statsHash.Count) mailboxes") -ForegroundColor Green

    # Step 3: Get FullAccess permissions for all mailboxes
    Write-Host ($PREFIX_INFO + 'Collecting FullAccess permissions...') -ForegroundColor DarkGray
    $fullAccessHash = @{}
    $processed = 0
    foreach ($mb in $allMailboxes) {
        $processed++
        if ($processed % 100 -eq 0) {
            Write-Host ($PREFIX_INFO + "  ...$processed / $($allMailboxes.Count) mailboxes") -ForegroundColor DarkGray
        }

        try {
            $perms = @(Get-MailboxPermission -Identity $mb.ExchangeGuid.ToString() -ErrorAction Stop |
                Where-Object {
                    $_.User -notlike 'NT AUTHORITY\SELF' -and
                    $_.User -notlike 'S-1-5-*' -and
                    $_.IsInherited -eq $false -and
                    $_.AccessRights -contains 'FullAccess'
                } | ForEach-Object {
                    [PSCustomObject]@{
                        User         = $_.User.ToString()
                        AccessRights = ($_.AccessRights | Sort-Object) -join ','
                    }
                })
            if ($perms.Count -gt 0) {
                $fullAccessHash[$mb.ExchangeGuid.ToString()] = $perms
            }
        }
        catch {
            # Permissions unavailable - leave empty
        }
    }
    Write-Host ($PREFIX_OK + "Found FullAccess permissions on $($fullAccessHash.Count) mailboxes") -ForegroundColor Green

    # Step 4: Get SendAs permissions for all mailboxes
    Write-Host ($PREFIX_INFO + 'Collecting SendAs permissions...') -ForegroundColor DarkGray
    $sendAsHash = @{}
    $processed = 0
    foreach ($mb in $allMailboxes) {
        $processed++
        if ($processed % 100 -eq 0) {
            Write-Host ($PREFIX_INFO + "  ...$processed / $($allMailboxes.Count) mailboxes") -ForegroundColor DarkGray
        }

        try {
            $perms = @(Get-RecipientPermission -Identity $mb.ExchangeGuid.ToString() -ErrorAction Stop |
                Where-Object {
                    $_.Trustee -notlike 'NT AUTHORITY\SELF' -and
                    $_.Trustee -notlike 'S-1-5-*' -and
                    $_.AccessRights -contains 'SendAs'
                } | ForEach-Object {
                    [PSCustomObject]@{
                        Trustee      = $_.Trustee.ToString()
                        AccessRights = ($_.AccessRights | Sort-Object) -join ','
                    }
                })
            if ($perms.Count -gt 0) {
                $sendAsHash[$mb.ExchangeGuid.ToString()] = $perms
            }
        }
        catch {
            # Permissions unavailable - leave empty
        }
    }
    Write-Host ($PREFIX_OK + "Found SendAs permissions on $($sendAsHash.Count) mailboxes") -ForegroundColor Green

    # Step 5: Build combined records
    Write-Host ($PREFIX_INFO + 'Building cache records...') -ForegroundColor DarkGray
    $records = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($mb in $allMailboxes) {
        $guid = $mb.ExchangeGuid.ToString()
        $stats = $statsHash[$guid]
        $fullAccess = $fullAccessHash[$guid]
        $sendAs = $sendAsHash[$guid]

        $record = [PSCustomObject]@{
            DisplayName           = $mb.DisplayName
            UserPrincipalName     = $mb.UserPrincipalName
            PrimarySmtpAddress    = if ($mb.PrimarySmtpAddress) { $mb.PrimarySmtpAddress.ToString() } else { '' }
            EmailAddresses        = if ($mb.EmailAddresses) { ($mb.EmailAddresses | Sort-Object) -join '|' } else { '' }
            Alias                 = $mb.Alias
            RecipientTypeDetails  = $mb.RecipientTypeDetails.ToString()
            ExchangeGuid          = $guid
            Database              = $mb.Database
            LitigationHoldEnabled = $mb.LitigationHoldEnabled
            ArchiveStatus         = $mb.ArchiveStatus.ToString()
            WhenCreated           = $mb.WhenCreated
            WhenMailboxCreated    = $mb.WhenMailboxCreated
            ProhibitSendQuota     = if ($mb.ProhibitSendQuota) { $mb.ProhibitSendQuota.ToString() } else { '' }
            ProhibitSendReceiveQuota = if ($mb.ProhibitSendReceiveQuota) { $mb.ProhibitSendReceiveQuota.ToString() } else { '' }
            ForwardingSmtpAddress = $mb.ForwardingSmtpAddress
            GrantSendOnBehalfTo   = if ($mb.GrantSendOnBehalfTo) { ($mb.GrantSendOnBehalfTo | ForEach-Object { $_.ToString() }) -join '|' } else { '' }
            CustomAttribute6      = $mb.CustomAttribute6
            CustomAttribute7      = $mb.CustomAttribute7
            # Statistics (null if not retrieved)
            ItemCount             = if ($stats) { $stats.ItemCount } else { $null }
            SizeGB                = if ($stats) { $stats.SizeGB } else { 0 }
            SizeBytes             = if ($stats) { $stats.SizeBytes } else { 0 }
            LastLogon             = if ($stats) { $stats.LastLogon } else { $null }
            # Permissions
            FullAccessPermissions = if ($fullAccess) { $fullAccess } else { @() }
            SendAsPermissions     = if ($sendAs) { $sendAs } else { @() }
        }
        $records.Add($record)
    }

    # Step 6: Save to cache file
    $cacheFile = Join-Path $CacheFolder 'Mailboxes-All.json'
    $records.ToArray() | ConvertTo-Json -Depth 10 | Set-Content -Path $cacheFile -Encoding UTF8
    Write-Host ($PREFIX_OK + "Cached $($records.Count) mailboxes with statistics and permissions") -ForegroundColor Green
    Write-Host ($PREFIX_INFO + "Cache file: $cacheFile") -ForegroundColor DarkGray
}

# -----------------------------------------------------------------------
# Exports
# -----------------------------------------------------------------------

Export-ModuleMember -Function 'Invoke-ExchangeCollection', 'Update-DistributionGroupsCacheFile', 'Update-MailboxesCacheFile'

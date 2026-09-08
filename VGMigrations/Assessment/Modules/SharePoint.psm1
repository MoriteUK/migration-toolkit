#Requires -Version 7.0

Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -DisableNameChecking -Force -Global
#Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -ErrorAction Stop
$prev = $WarningPreference
$WarningPreference = 'SilentlyContinue'
Import-Module Microsoft.Online.SharePoint.PowerShell `
    -UseWindowsPowerShell -DisableNameChecking -ErrorAction Stop
$WarningPreference = $prev

# -----------------------------------------------------------------------
# Private functions
# -----------------------------------------------------------------------

<#
.SYNOPSIS
    Builds a site output record with null guards on every property due to -UseWindowsPowerShell deserialization.
#>
function Build-SPOSiteRecord {
    param([object]$Site)

    $groupId = $null
    if ($Site.GroupId) {
        $gStr = $Site.GroupId.ToString()
        if ($gStr -ne '00000000-0000-0000-0000-000000000000') { $groupId = $gStr }
    }

    $siteId    = if ($Site.Id)                  { $Site.Id.ToString() }                                                    else { '' }
    $usedGB    = if ($Site.StorageUsageCurrent)  { Convert-SizeToGB -Value $Site.StorageUsageCurrent -FromMB }              else { [double]0 }
    $usedBytes = if ($Site.StorageUsageCurrent)  { Convert-SizeToBytes -Value $Site.StorageUsageCurrent -FromMB }           else { [long]0 }
    $quotaGB   = if ($Site.StorageQuota)         { Convert-SizeToGB -Value $Site.StorageQuota -FromMB }                     else { [double]0 }
    $usedPct = if ($Site.StorageQuota -gt 0)   { [math]::Round($Site.StorageUsageCurrent / $Site.StorageQuota * 100, 1) } else { [double]0 }
    $owner   = if ($Site.Owner)                { $Site.Owner }                                                            else { '' }
    $lastMod = if ($Site.LastContentModifiedDate) { $Site.LastContentModifiedDate }                                       else { $null }

    $isChannelSite = if ($Site.IsTeamsChannelConnected) { [bool]$Site.IsTeamsChannelConnected } else { $false }

    # Channel sites carry the parent team's group ID in RelatedGroupId, not GroupId
    $relatedGroupId = $null
    if ($Site.RelatedGroupId) {
        $rStr = $Site.RelatedGroupId.ToString()
        if ($rStr -ne '00000000-0000-0000-0000-000000000000') { $relatedGroupId = $rStr }
    }

    [PSCustomObject]@{
        Title                   = $Site.Title
        Url                     = $Site.Url
        Template                = $Site.Template
        SiteId                  = $siteId
        GroupId                 = $groupId
        RelatedGroupId          = $relatedGroupId
        IsTeamsChannelConnected = $isChannelSite
        StorageUsedGB           = $usedGB
        StorageUsedBytes        = $usedBytes
        StorageQuotaGB          = $quotaGB
        StorageUsedPercent      = $usedPct
        SiteOwner               = $owner
        LastModified            = $lastMod
        IsTeamsConnected        = $null
        AssociatedObject        = $null
        MigrationObjectType     = $null
        # Initialised here so the property exists before Workbook.psm1 assigns it
        IsOrphaned              = $false
    }
}

<#
.SYNOPSIS
    Retrieves all SPO sites via Get-SPOSite -Limit All and filters client-side to the VBU scope on URL and title.
#>
function Get-SPOSiteData {
    param([PSCustomObject]$Context)

    $sites      = [System.Collections.Generic.List[PSCustomObject]]::new()
    $searchTerm = $Context.VBUSearchTerm
    $domain     = $Context.VBUDomain

    $allSites = @(Get-SPOSite -Limit All -ErrorAction Stop)
    foreach ($s in $allSites) {
        if (($s.Url   -like "*$searchTerm*") -or
            ($s.Url   -like "*$domain*")     -or
            (Test-WholeWordMatch -Text $s.Title -Term $searchTerm)) {
            $sites.Add((Build-SPOSiteRecord -Site $s))
        }
    }

    return $sites.ToArray()
}

<#
.SYNOPSIS
    Retrieves personal sites and keeps those whose owner login (claims prefix stripped) matches an in-scope user mailbox.
.DESCRIPTION
    OneDrives are a migration target, so they follow the same scope as mailboxes rather
    than the broader AD user scope (ADUsers.json is intentionally broad - it also catches
    sibling-VBU users who carry the VBU domain as a proxy address, kept for domain-blocker
    cleanup, not migration). Scoped to UserMailboxes.json only - not SharedMailboxes.json -
    since a shared mailbox is not a migrating user and its OneDrive (if any) should not
    appear on the OneDrives tab. The owner lookup is keyed on both PrimarySmtpAddress and
    UserPrincipalName (lowercased) so either address form on the OneDrive owner login resolves.
#>
function Get-OneDriveData {
    param([PSCustomObject]$Context)

    $userMailboxesPath = Join-Path $Context.RawPath 'UserMailboxes.json'
    if (-not (Test-Path $userMailboxesPath)) {
        Write-Host ($PREFIX_WARN + 'UserMailboxes.json not found - OneDrive cross-reference skipped') -ForegroundColor Yellow
        return @()
    }
    $userMailboxes = @(Get-Content $userMailboxesPath -Raw | ConvertFrom-Json)

    $loginToMailbox = @{}
    foreach ($mb in $userMailboxes) {
        if ($mb.PrimarySmtpAddress) { $loginToMailbox["$($mb.PrimarySmtpAddress)".ToLower()] = $mb }
        if ($mb.UserPrincipalName)  { $loginToMailbox["$($mb.UserPrincipalName)".ToLower()]  = $mb }
    }

    $rawSites = @(Get-SPOSite -IncludePersonalSite $true -Limit All -ErrorAction Stop)
    $records  = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($od in ($rawSites | Where-Object { $_.Url -like '*/personal/*' })) {
        # Strip claims provider prefix: i:0#.f|membership|user@domain.com -> user@domain.com
        $ownerLogin = $od.Owner
        if ($ownerLogin -like '*|*') { $ownerLogin = $ownerLogin.Split('|')[-1] }

        if (-not $loginToMailbox.ContainsKey($ownerLogin.ToLower())) { continue }

        $mailbox  = $loginToMailbox[$ownerLogin.ToLower()]
        $usedGB   = Convert-SizeToGB -Value $od.StorageUsageCurrent -FromMB
        $quotaGB  = Convert-SizeToGB -Value $od.StorageQuota -FromMB
        $remainGB = [math]::Round($quotaGB - $usedGB, 2)
        $usedPct  = if ($od.StorageQuota -gt 0) {
            [math]::Round($od.StorageUsageCurrent / $od.StorageQuota * 100, 1)
        } else { [double]0 }

        $records.Add([PSCustomObject]@{
            OwnerUPN           = $ownerLogin
            OwnerDisplayName   = $mailbox.DisplayName
            OneDriveUrl        = $od.Url
            StorageUsedGB      = $usedGB
            StorageQuotaGB     = $quotaGB
            StorageRemainingGB = $remainGB
            StorageUsedPercent = $usedPct
            LastModified       = $od.LastContentModifiedDate
            IsProvisioned      = ($od.Status -eq 'Active')
        })
    }

    return $records.ToArray()
}

# -----------------------------------------------------------------------
# Public functions
# -----------------------------------------------------------------------

<#
.SYNOPSIS
    Orchestrates SharePoint site and OneDrive collection.
.DESCRIPTION
    Returns a skipped result immediately when Context.SkipSharePoint is set. Otherwise
    collects sites and OneDrive accounts over the session connected by the orchestrator
    and writes SharePointSites.json and OneDrives.json. Failures are non-critical and
    are returned as a failed collector result.
#>
function Invoke-SharePointCollection {
    [CmdletBinding()]
    param([Parameter(Mandatory)][PSCustomObject]$Context)

    $start = Get-Date
    Write-SectionHeader 'SharePoint Online'

    if ($Context.SkipSharePoint) {
        Write-Host ($PREFIX_SKIP + 'SharePoint collection skipped (SkipSharePoint = true)') -ForegroundColor Yellow
        Update-CollectorStatus -CollectorName 'SharePoint Data' -Status 'Skipped' `
            -RawPath $Context.RawPath -StartTime $start
        return New-CollectorResult -Skipped $true
    }

    try {
        Write-Host ($PREFIX_INFO + 'Collecting SharePoint sites...') -ForegroundColor DarkGray
        $sites = Get-SPOSiteData -Context $Context
        Write-ProgressLine -Label 'SharePoint Sites' -Count $sites.Count

        Write-Host ($PREFIX_INFO + 'Collecting OneDrive accounts...') -ForegroundColor DarkGray
        $oneDrives = Get-OneDriveData -Context $Context
        Write-ProgressLine -Label 'OneDrive Accounts' -Count $oneDrives.Count

        Write-JsonOutput -FileName 'SharePointSites.json' -Data $sites     -RawPath $Context.RawPath
        Write-JsonOutput -FileName 'OneDrives.json'       -Data $oneDrives -RawPath $Context.RawPath

        $counts = @{
            SPOSiteCount  = $sites.Count
            OneDriveCount = $oneDrives.Count
        }
        $msg = "Sites: $($counts.SPOSiteCount), OneDrives: $($counts.OneDriveCount)"
        Update-CollectorStatus -CollectorName 'SharePoint Data' -Status 'Complete' `
            -RawPath $Context.RawPath -StartTime $start -Message $msg

        return New-CollectorResult -Success $true -Counts $counts
    }
    catch {
        Write-Host ($PREFIX_FAIL + 'SharePoint collection failed: ' + $_.Exception.Message) -ForegroundColor Red
        Update-CollectorStatus -CollectorName 'SharePoint Data' -Status 'Failed' `
            -RawPath $Context.RawPath -StartTime $start -Message $_.Exception.Message
        return New-CollectorResult -Success $false -ErrorMessage $_.Exception.Message
    }
}

# -----------------------------------------------------------------------
# Exports
# -----------------------------------------------------------------------

Export-ModuleMember -Function 'Invoke-SharePointCollection'

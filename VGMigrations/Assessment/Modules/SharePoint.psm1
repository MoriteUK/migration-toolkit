#Requires -Version 7.0

Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -DisableNameChecking -Force -Global

# SharePoint module is optional - if not available, the functions in this module will not work
# but the module itself should load successfully (caller checks SkipSharePoint flag)
try {
    $prev = $WarningPreference
    $WarningPreference = 'SilentlyContinue'
    Import-Module Microsoft.Online.SharePoint.PowerShell `
        -UseWindowsPowerShell -DisableNameChecking -ErrorAction Stop
    $WarningPreference = $prev
}
catch {
    $WarningPreference = $prev
    # Silently skip - caller already warned the user
}

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
    Retrieves all SPO sites - from cache if available and fresh, otherwise via Get-SPOSite -Limit All.
.DESCRIPTION
    Checks for cached SharePointSites-All.json (< 7 days old). If found, loads from cache and filters
    to VBU scope. Otherwise falls back to Get-SPOSite -Limit All (slow tenant-wide query).
#>
function Get-SPOSiteData {
    param([PSCustomObject]$Context)

    $sites      = [System.Collections.Generic.List[PSCustomObject]]::new()
    $searchTerm = $Context.VBUSearchTerm
    $domain     = $Context.VBUDomain
    $cacheFile  = if ($Context.CacheFolder) { Join-Path $Context.CacheFolder 'SharePointSites-All.json' } else { '' }

    # Try cache first
    if ($cacheFile -and (Test-Path $cacheFile)) {
        $cacheAge = (Get-Date) - (Get-Item $cacheFile).LastWriteTime
        if ($cacheAge.TotalDays -lt 7) {
            Write-Host ($PREFIX_INFO + "Loading SharePoint sites from cache ($([math]::Round($cacheAge.TotalDays, 1)) days old)...") -ForegroundColor DarkGray
            $allSites = @(Get-Content $cacheFile -Raw | ConvertFrom-Json)

            foreach ($s in $allSites) {
                if (($s.Url   -like "*$searchTerm*") -or
                    ($s.Url   -like "*$domain*")     -or
                    (Test-WholeWordMatch -Text $s.Title -Term $searchTerm)) {
                    $sites.Add($s)
                }
            }
            return $sites.ToArray()
        }
    }

    # No cache or cache stale - query live
    Write-Host ($PREFIX_INFO + 'Querying SharePoint sites live (no cache)...') -ForegroundColor DarkGray
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
    Retrieves personal sites - from cache if available, otherwise via Get-SPOSite -IncludePersonalSite.
.DESCRIPTION
    OneDrives are a migration target, so they follow the same scope as mailboxes rather
    than the broader AD user scope (ADUsers.json is intentionally broad - it also catches
    sibling-VBU users who carry the VBU domain as a proxy address, kept for domain-blocker
    cleanup, not migration). Scoped to UserMailboxes.json only - not SharedMailboxes.json -
    since a shared mailbox is not a migrating user and its OneDrive (if any) should not
    appear on the OneDrives tab. The owner lookup is keyed on both PrimarySmtpAddress and
    UserPrincipalName (lowercased) so either address form on the OneDrive owner login resolves.

    Checks for cached OneDrives-All.json (< 7 days old). If found, loads from cache and filters
    to in-scope mailboxes. Otherwise falls back to Get-SPOSite -IncludePersonalSite.
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

    $records   = [System.Collections.Generic.List[PSCustomObject]]::new()
    $cacheFile = if ($Context.CacheFolder) { Join-Path $Context.CacheFolder 'OneDrives-All.json' } else { '' }

    # Try cache first
    if ($cacheFile -and (Test-Path $cacheFile)) {
        $cacheAge = (Get-Date) - (Get-Item $cacheFile).LastWriteTime
        if ($cacheAge.TotalDays -lt 7) {
            Write-Host ($PREFIX_INFO + "Loading OneDrives from cache ($([math]::Round($cacheAge.TotalDays, 1)) days old)...") -ForegroundColor DarkGray
            $allOneDrives = @(Get-Content $cacheFile -Raw | ConvertFrom-Json)

            foreach ($od in $allOneDrives) {
                $ownerLogin = $od.OwnerUPN
                if (-not $loginToMailbox.ContainsKey($ownerLogin.ToLower())) { continue }

                $mailbox = $loginToMailbox[$ownerLogin.ToLower()]
                $records.Add([PSCustomObject]@{
                    OwnerUPN           = $ownerLogin
                    OwnerDisplayName   = $mailbox.DisplayName
                    OneDriveUrl        = $od.OneDriveUrl
                    StorageUsedGB      = $od.StorageUsedGB
                    StorageQuotaGB     = $od.StorageQuotaGB
                    StorageRemainingGB = [math]::Round($od.StorageQuotaGB - $od.StorageUsedGB, 2)
                    StorageUsedPercent = $od.StorageUsedPercent
                    LastModified       = $od.LastModified
                    IsProvisioned      = $od.IsProvisioned
                })
            }
            return $records.ToArray()
        }
    }

    # No cache or cache stale - query live
    Write-Host ($PREFIX_INFO + 'Querying OneDrive accounts live (no cache)...') -ForegroundColor DarkGray
    $rawSites = @(Get-SPOSite -IncludePersonalSite $true -Limit All -ErrorAction Stop)

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

<#
.SYNOPSIS
    Updates the tenant-wide SharePoint sites cache file.
.DESCRIPTION
    Runs Get-SPOSite -Limit All and Get-SPOSite -IncludePersonalSite $true -Limit All
    to collect all sites and OneDrive accounts, then saves them to cache files that
    individual VBU assessments can read and filter locally instead of re-scanning the
    entire tenant every time.
#>
function Update-SharePointSitesCacheFile {
    param(
        [Parameter(Mandatory)]
        [string]$CacheFolder
    )

    Write-Host ($PREFIX_INFO + 'Collecting all SharePoint sites...') -ForegroundColor DarkGray
    $allSites = @(Get-SPOSite -Limit All -ErrorAction Stop)

    # Build records for all sites
    $siteRecords = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($site in $allSites) {
        $siteRecords.Add((Build-SPOSiteRecord -Site $site))
    }

    $sitesFile = Join-Path $CacheFolder 'SharePointSites-All.json'
    $siteRecords.ToArray() | ConvertTo-Json -Depth 10 | Set-Content -Path $sitesFile -Encoding UTF8
    Write-Host ($PREFIX_OK + "Cached $($siteRecords.Count) SharePoint sites") -ForegroundColor Green

    Write-Host ($PREFIX_INFO + 'Collecting all OneDrive accounts...') -ForegroundColor DarkGray
    $allPersonalSites = @(Get-SPOSite -IncludePersonalSite $true -Limit All -ErrorAction Stop |
        Where-Object { $_.Url -like '*/personal/*' })

    # Build OneDrive records (simplified - no mailbox cross-reference for cache)
    $oneDriveRecords = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($od in $allPersonalSites) {
        $ownerLogin = $od.Owner
        if ($ownerLogin -like '*|*') { $ownerLogin = $ownerLogin.Split('|')[-1] }

        $usedGB = Convert-SizeToGB -Value $od.StorageUsageCurrent -FromMB
        $quotaGB = Convert-SizeToGB -Value $od.StorageQuota -FromMB
        $usedPct = if ($od.StorageQuota -gt 0) {
            [math]::Round($od.StorageUsageCurrent / $od.StorageQuota * 100, 1)
        } else { [double]0 }

        $oneDriveRecords.Add([PSCustomObject]@{
            OwnerUPN = $ownerLogin
            OneDriveUrl = $od.Url
            StorageUsedGB = $usedGB
            StorageQuotaGB = $quotaGB
            StorageUsedPercent = $usedPct
            LastModified = $od.LastContentModifiedDate
            IsProvisioned = ($od.Status -eq 'Active')
        })
    }

    $oneDrivesFile = Join-Path $CacheFolder 'OneDrives-All.json'
    $oneDriveRecords.ToArray() | ConvertTo-Json -Depth 10 | Set-Content -Path $oneDrivesFile -Encoding UTF8
    Write-Host ($PREFIX_OK + "Cached $($oneDriveRecords.Count) OneDrive accounts") -ForegroundColor Green
}

# -----------------------------------------------------------------------
# Exports
# -----------------------------------------------------------------------

Export-ModuleMember -Function 'Invoke-SharePointCollection', 'Update-SharePointSitesCacheFile'

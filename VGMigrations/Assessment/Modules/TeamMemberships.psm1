#Requires -Version 7.0

Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -DisableNameChecking -Force -Global

<#
.SYNOPSIS
    Assessment-engine port of the standalone Get-DomainTeamMemberships.ps1.
.DESCRIPTION
    Reports every Team (and private/shared channel) that VBU-domain users belong to, so they
    can be re-added afterwards to whichever ones don't migrate with them (see
    Restore-DomainTeamMemberships.ps1). Runs on the Graph session Run-Assessment.ps1 already
    established - no separate sign-in, unlike the standalone script this replicates. Writes the
    same CSV schema that script always has, so Restore-DomainTeamMemberships.ps1 works
    unchanged regardless of which one produced the CSV.

    The tenant-wide walk (every Team, every channel, every member) used to happen live on every
    Discovery run - that's the expensive part on a large tenant. It now lives only in
    Update-TeamsChannelsCacheFile, run standalone via Update-TeamsChannelsCache.ps1 (or on demand
    from the app's Discovery Cache screen). Invoke-TeamMembershipCollection (the Discovery-time
    path) only reads the cached TeamsChannelsMembers.json and matches it against the VBU's own
    (live-queried, cheap) user list - Run-Assessment.ps1 checks the cache is present and fresh
    (Test-DiscoveryCachePrereqs) before this ever runs. -SkipTeamMemberships still exists as a
    plain opt-out, though it's no longer needed for speed now that this is cache-backed.
#>

# -----------------------------------------------------------------------
# Private functions
# -----------------------------------------------------------------------

<#
.SYNOPSIS
    Retries a Graph call with exponential backoff on throttling (429/503/504).
#>
function Invoke-ThrottledGraph {
    param([scriptblock]$Script, [string]$What = 'Graph call')
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        try { return & $Script }
        catch {
            $status = $null
            try { $status = $_.Exception.Response.StatusCode.value__ } catch {}
            if ($status -in 429, 503, 504 -and $attempt -lt 5) {
                $wait = [math]::Min(60, [math]::Pow(2, $attempt))
                Write-Host ($PREFIX_WARN + "throttled on $What (HTTP $status) - retry $attempt in ${wait}s") -ForegroundColor Yellow
                Start-Sleep -Seconds $wait
                continue
            }
            throw
        }
    }
}

<#
.SYNOPSIS
    Builds the id/UPN/mail lookup of every user in the VBU domain.
#>
function Get-VBUDomainUserIndex {
    param([string]$Domain)

    $domainUsers = @{}
    $props = 'Id', 'DisplayName', 'UserPrincipalName', 'Mail', 'UserType'

    function Add-DomainUser {
        param($u)
        if (-not $u -or -not $u.Id) { return }
        $domainUsers[$u.Id] = [pscustomobject]@{
            DisplayName = $u.DisplayName
            UPN         = $u.UserPrincipalName
            Mail        = $u.Mail
            UserType    = $u.UserType
        }
    }

    try {
        Invoke-ThrottledGraph -What 'Get-MgUser endsWith(UPN)' -Script {
            Get-MgUser -All -Property $props -ConsistencyLevel eventual -CountVariable _c `
                -Filter "endsWith(userPrincipalName,'@$Domain')"
        } | ForEach-Object { Add-DomainUser $_ }

        Invoke-ThrottledGraph -What 'Get-MgUser endsWith(mail)' -Script {
            Get-MgUser -All -Property $props -ConsistencyLevel eventual -CountVariable _c `
                -Filter "endsWith(mail,'@$Domain')"
        } | ForEach-Object { Add-DomainUser $_ }
    }
    catch {
        Write-Host ($PREFIX_WARN + "Advanced user filter failed ($($_.Exception.Message.Split([Environment]::NewLine)[0])) - falling back to a full user scan") -ForegroundColor Yellow
        Invoke-ThrottledGraph -What 'Get-MgUser -All' -Script { Get-MgUser -All -Property $props } |
            Where-Object {
                ("" + $_.UserPrincipalName).ToLowerInvariant().EndsWith("@$Domain") -or
                ("" + $_.Mail).ToLowerInvariant().EndsWith("@$Domain")
            } | ForEach-Object { Add-DomainUser $_ }
    }

    return $domainUsers
}

<#
.SYNOPSIS
    Resolves a Teams/channel member (aadUserConversationMember) to a matched VBU-domain user, or $null.
#>
function Resolve-DomainMember {
    param($Member, [hashtable]$DomainUsers, [string]$Domain)

    $uid   = "" + $Member.AdditionalProperties['userId']
    $email = ("" + $Member.AdditionalProperties['email']).ToLowerInvariant()
    $hit   = $null
    if ($uid -and $DomainUsers.ContainsKey($uid)) { $hit = $DomainUsers[$uid] }
    elseif ($email.EndsWith("@$Domain")) {
        $hit = [pscustomobject]@{ DisplayName = $Member.DisplayName; UPN = ''; Mail = $Member.AdditionalProperties['email']; UserType = '' }
    }
    if (-not $hit) { return $null }

    $roles = @($Member.Roles)
    $role  = if ($roles -contains 'owner') { 'Owner' }
             elseif ($roles -contains 'guest') { 'Guest' }
             else { 'Member' }
    [pscustomobject]@{
        DisplayName = if ($hit.DisplayName) { $hit.DisplayName } else { $Member.DisplayName }
        UPN         = $hit.UPN
        Mail        = $hit.Mail
        UserType    = $hit.UserType
        UserId      = $uid
        Role        = $role
    }
}

<#
.SYNOPSIS
    Flattens a live Graph conversation-member object to the plain schema stored in the cache file.
#>
function ConvertTo-CachedMemberRecord {
    param($Member)
    [pscustomobject]@{
        UserId      = "" + $Member.AdditionalProperties['userId']
        DisplayName = $Member.DisplayName
        Email       = $Member.AdditionalProperties['email']
        Roles       = @($Member.Roles)
    }
}

<#
.SYNOPSIS
    Resolves a cached member record (plain properties, post-JSON-round-trip) to a matched VBU-domain user, or $null.
.DESCRIPTION
    Same matching logic as the live-Graph Resolve-DomainMember above, adapted for cached
    TeamsChannelsMembers.json records - those no longer carry a Graph SDK AdditionalProperties
    hashtable once they've been through ConvertTo-Json/ConvertFrom-Json, just plain UserId/Email/Roles.
#>
function Resolve-CachedDomainMember {
    param($Member, [hashtable]$DomainUsers, [string]$Domain)

    $uid   = "" + $Member.UserId
    $email = ("" + $Member.Email).ToLowerInvariant()
    $hit   = $null
    if ($uid -and $DomainUsers.ContainsKey($uid)) { $hit = $DomainUsers[$uid] }
    elseif ($email.EndsWith("@$Domain")) {
        $hit = [pscustomobject]@{ DisplayName = $Member.DisplayName; UPN = ''; Mail = $Member.Email; UserType = '' }
    }
    if (-not $hit) { return $null }

    $roles = @($Member.Roles)
    $role  = if ($roles -contains 'owner') { 'Owner' }
             elseif ($roles -contains 'guest') { 'Guest' }
             else { 'Member' }
    [pscustomobject]@{
        DisplayName = if ($hit.DisplayName) { $hit.DisplayName } else { $Member.DisplayName }
        UPN         = $hit.UPN
        Mail        = $hit.Mail
        UserType    = $hit.UserType
        UserId      = $uid
        Role        = $role
    }
}

# -----------------------------------------------------------------------
# Public functions
# -----------------------------------------------------------------------

<#
.SYNOPSIS
    Walks every Team, every channel, and every member in the connected tenant and writes the cache file.
.DESCRIPTION
    Called by Update-TeamsChannelsCache.ps1, which owns the Graph sign-in first. Unfiltered -
    every Team tenant-wide, not scoped to any one VBU domain, so the same cache serves every VBU
    split from this source tenant. Standard channels record their metadata only (Members left
    empty) - their membership always equals the parent Team's, which is already captured.
#>
function Update-TeamsChannelsCacheFile {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$CacheFolder)

    Write-Host ($PREFIX_INFO + 'Enumerating all Teams tenant-wide...') -ForegroundColor DarkGray
    $teams = @(Invoke-ThrottledGraph -What 'Get-MgTeam -All' -Script { Get-MgTeam -All -Property 'Id,DisplayName,IsArchived' })
    Write-ProgressLine -Label 'Teams (tenant-wide)' -Count $teams.Count

    $result    = [System.Collections.Generic.List[object]]::new()
    $teamIndex = 0

    foreach ($team in $teams) {
        $teamIndex++
        if ($teamIndex % 25 -eq 0) {
            Write-Host ($PREFIX_INFO + "  ...$teamIndex / $($teams.Count) teams") -ForegroundColor DarkGray
        }

        $grp = $null
        try {
            $grp = Invoke-ThrottledGraph -What "Get-MgGroup $($team.Id)" -Script {
                Get-MgGroup -GroupId $team.Id -Property 'Mail,Visibility,DisplayName'
            }
        }
        catch { }

        $memberRecords = @()
        try {
            $members = @(Invoke-ThrottledGraph -What "Get-MgTeamMember $($team.DisplayName)" -Script {
                Get-MgTeamMember -TeamId $team.Id -All
            })
            $memberRecords = @($members | ForEach-Object { ConvertTo-CachedMemberRecord -Member $_ })
        }
        catch {
            Write-Host ($PREFIX_WARN + "$($team.DisplayName): could not read members - $($_.Exception.Message.Split([Environment]::NewLine)[0])") -ForegroundColor Yellow
        }

        $channelRecords = [System.Collections.Generic.List[object]]::new()
        try {
            $channels = @(Invoke-ThrottledGraph -What "Get-MgTeamChannel $($team.DisplayName)" -Script {
                Get-MgTeamChannel -TeamId $team.Id -All -Property 'Id,DisplayName,MembershipType'
            })
        }
        catch {
            Write-Host ($PREFIX_WARN + "$($team.DisplayName): could not read channels - $($_.Exception.Message.Split([Environment]::NewLine)[0])") -ForegroundColor Yellow
            $channels = @()
        }

        foreach ($ch in $channels) {
            $type     = "" + $ch.MembershipType
            $chRecord = [pscustomobject]@{
                ChannelId      = $ch.Id
                ChannelName    = $ch.DisplayName
                MembershipType = $type
                Members        = @()
            }

            if ($type -in 'private', 'shared') {
                try {
                    $chMembers = @(Invoke-ThrottledGraph -What "Get-MgTeamChannelMember $($ch.DisplayName)" -Script {
                        Get-MgTeamChannelMember -TeamId $team.Id -ChannelId $ch.Id -All
                    })
                    $chRecord.Members = @($chMembers | ForEach-Object { ConvertTo-CachedMemberRecord -Member $_ })
                }
                catch {
                    Write-Host ($PREFIX_WARN + "$($team.DisplayName) / $($ch.DisplayName): could not read channel members - $($_.Exception.Message.Split([Environment]::NewLine)[0])") -ForegroundColor Yellow
                }
            }

            $channelRecords.Add($chRecord)
        }

        $result.Add([pscustomobject]@{
            TeamId          = $team.Id
            TeamDisplayName = $team.DisplayName
            TeamMail        = $grp.Mail
            TeamVisibility  = $grp.Visibility
            TeamArchived    = [bool]$team.IsArchived
            Members         = $memberRecords
            Channels        = $channelRecords.ToArray()
        })
    }

    New-Item -ItemType Directory -Path $CacheFolder -Force | Out-Null
    $cachePath = Join-Path $CacheFolder 'TeamsChannelsMembers.json'
    $result.ToArray() | ConvertTo-Json -Depth 10 | Set-Content -Path $cachePath -Encoding UTF8
    Write-Host ($PREFIX_OK + "TeamsChannelsMembers.json cache written: $cachePath") -ForegroundColor Green

    return $result.Count
}

<#
.SYNOPSIS
    Orchestrates the tenant-wide Team/channel membership scan for the VBU domain.
.DESCRIPTION
    Writes <OutputCsvPath> directly (same schema as Get-DomainTeamMemberships.ps1) and a
    TeamMemberships.json copy to the Raw folder for the assessment summary. Non-critical like
    every other collector - a failure is recorded and returned, never thrown.
#>
function Invoke-TeamMembershipCollection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][PSCustomObject]$Context,
        [Parameter(Mandatory)][string]$OutputCsvPath,
        [Parameter(Mandatory)][string]$CacheFolder,
        [switch]$IncludeStandardChannels
    )

    $start = Get-Date
    Write-SectionHeader 'Domain Team Memberships'

    $domain = $Context.VBUDomain.ToLowerInvariant()

    try {
        Write-Host ($PREFIX_INFO + "Enumerating users in @$domain...") -ForegroundColor DarkGray
        $domainUsers = Get-VBUDomainUserIndex -Domain $domain
        Write-ProgressLine -Label 'VBU domain users' -Count $domainUsers.Count

        if ($domainUsers.Count -eq 0) {
            Write-Host ($PREFIX_WARN + 'No users found in this domain - nothing to scan') -ForegroundColor Yellow
            Write-JsonOutput -FileName 'TeamMemberships.json' -Data @() -RawPath $Context.RawPath
            [pscustomobject]@{ Note = 'No data' } | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding UTF8
            Update-CollectorStatus -CollectorName 'Team Memberships' -Status 'Complete' `
                -RawPath $Context.RawPath -StartTime $start -Message 'No domain users found'
            return New-CollectorResult -Success $true -Counts @{ TeamMembershipRowCount = 0 }
        }

        Write-Host ($PREFIX_INFO + 'Reading Teams/channels/members from cache...') -ForegroundColor DarkGray
        $cachedTeams = Import-AssessmentJson -FileName 'TeamsChannelsMembers.json' -RawPath $CacheFolder
        Write-ProgressLine -Label 'Teams (from cache)' -Count $cachedTeams.Count

        $rows = [System.Collections.Generic.List[object]]::new()

        foreach ($team in $cachedTeams) {
            $matched = @(@(foreach ($m in $team.Members) { Resolve-CachedDomainMember -Member $m -DomainUsers $domainUsers -Domain $domain }) | Where-Object { $_ })

            foreach ($u in $matched) {
                $rows.Add([pscustomobject]@{
                    TeamDisplayName   = $team.TeamDisplayName
                    TeamMail          = $team.TeamMail
                    TeamVisibility    = $team.TeamVisibility
                    TeamArchived      = [bool]$team.TeamArchived
                    Scope             = 'Team'
                    ChannelName       = ''
                    ChannelType       = ''
                    UserDisplayName   = $u.DisplayName
                    UserPrincipalName = $u.UPN
                    UserMail          = $u.Mail
                    UserType          = $u.UserType
                    Role              = $u.Role
                    TeamId            = $team.TeamId
                    ChannelId         = ''
                    UserId            = $u.UserId
                })
            }

            foreach ($ch in $team.Channels) {
                $type = "" + $ch.MembershipType
                if ($type -in 'private', 'shared') {
                    $chMatched = @(@(foreach ($m in $ch.Members) { Resolve-CachedDomainMember -Member $m -DomainUsers $domainUsers -Domain $domain }) | Where-Object { $_ })
                }
                elseif ($IncludeStandardChannels) {
                    $chMatched = $matched
                }
                else {
                    continue
                }

                foreach ($u in $chMatched) {
                    $rows.Add([pscustomobject]@{
                        TeamDisplayName   = $team.TeamDisplayName
                        TeamMail          = $team.TeamMail
                        TeamVisibility    = $team.TeamVisibility
                        TeamArchived      = [bool]$team.TeamArchived
                        Scope             = 'Channel'
                        ChannelName       = $ch.ChannelName
                        ChannelType       = $type
                        UserDisplayName   = $u.DisplayName
                        UserPrincipalName = $u.UPN
                        UserMail          = $u.Mail
                        UserType          = $u.UserType
                        Role              = $u.Role
                        TeamId            = $team.TeamId
                        ChannelId         = $ch.ChannelId
                        UserId            = $u.UserId
                    })
                }
            }
        }

        $rows = @($rows | Sort-Object TeamDisplayName, Scope, ChannelName, UserDisplayName)

        if ($rows.Count -gt 0) {
            $rows | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding UTF8
        } else {
            [pscustomobject]@{ Note = 'No data' } | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding UTF8
        }
        Write-JsonOutput -FileName 'TeamMemberships.json' -Data $rows -RawPath $Context.RawPath

        $teamCount = (@($rows | Select-Object -ExpandProperty TeamId -Unique)).Count
        $userCount = (@($rows | Where-Object { $_.UserId } | Select-Object -ExpandProperty UserId -Unique)).Count
        Write-ProgressLine -Label 'Team Membership Rows' -Count $rows.Count
        Write-ProgressLine -Label 'Distinct Teams Matched' -Count $teamCount

        $counts = @{
            TeamMembershipRowCount  = $rows.Count
            TeamMembershipTeamCount = $teamCount
            TeamMembershipUserCount = $userCount
        }
        Update-CollectorStatus -CollectorName 'Team Memberships' -Status 'Complete' `
            -RawPath $Context.RawPath -StartTime $start -Message "Rows: $($rows.Count), Teams: $teamCount, Users: $userCount"

        return New-CollectorResult -Success $true -Counts $counts
    }
    catch {
        Write-Host ($PREFIX_FAIL + 'Team Membership collection failed: ' + $_.Exception.Message) -ForegroundColor Red
        Update-CollectorStatus -CollectorName 'Team Memberships' -Status 'Failed' `
            -RawPath $Context.RawPath -StartTime $start -Message $_.Exception.Message
        return New-CollectorResult -Success $false -ErrorMessage $_.Exception.Message
    }
}

# -----------------------------------------------------------------------
# Exports
# -----------------------------------------------------------------------

Export-ModuleMember -Function 'Invoke-TeamMembershipCollection', 'Update-TeamsChannelsCacheFile'

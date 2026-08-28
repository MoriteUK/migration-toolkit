#Requires -Version 7.0
<#
.SYNOPSIS
    Get-DomainTeamMemberships.ps1 — Reports every Team (and private/shared channel) that
    users from a given email domain belong to, written to a CSV.

.DESCRIPTION
    Before a domain is split out into its own tenant you lose the membership links between
    those users and any Teams that are NOT being migrated with them. This script captures
    that list from the CURRENT (source) tenant so the users can be re-invited/re-added
    afterwards.

    For every Team in the tenant it records which members have a UPN or mail address in the
    target domain, their role (Owner / Member / Guest), and does the same for every PRIVATE
    and SHARED channel (whose membership is tracked separately from the parent team).

    Standard channels are NOT listed per-user: their membership always equals the parent
    team's membership, so re-adding a user to the team restores every standard channel
    automatically. Pass -IncludeStandardChannels to list them anyway.

    Connects with Microsoft Graph (delegated). The signed-in account needs to be able to
    read Teams, channels and members org-wide (Global Reader / Teams Administrator, or an
    admin who can consent to the scopes below).

.PARAMETER Domain
    The email domain to report on, e.g. contoso.com. A user matches if their
    userPrincipalName OR mail ends with @<Domain> (so guest accounts whose UPN is
    user_contoso.com#EXT#@... still match on their mail address).

.PARAMETER User
    Optional. Restrict the report to a single person — their UPN or primary mail address.
    Still limited to the -Domain given.

.PARAMETER OutputCsv
    Path for the CSV. Defaults to .\DomainTeamMemberships-<domain>-<timestamp>.csv in the
    current directory.

.PARAMETER IncludeStandardChannels
    Also emit a row per standard channel for each matched user (normally redundant — see
    DESCRIPTION).

.PARAMETER TenantId
    Optional tenant id / domain to sign into, passed straight to Connect-MgGraph. Use this
    if your admin account can see more than one tenant.

.EXAMPLE
    .\Get-DomainTeamMemberships.ps1 -Domain contoso.com

.EXAMPLE
    .\Get-DomainTeamMemberships.ps1 -Domain contoso.com -User jane.doe@contoso.com -OutputCsv C:\Temp\jane-teams.csv
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$Domain,

    [string]$User = '',

    [string]$OutputCsv = '',

    [switch]$IncludeStandardChannels,

    [string]$TenantId = ''
)

$ErrorActionPreference = 'Stop'
$Domain = $Domain.Trim().TrimStart('@').ToLowerInvariant()
$User   = $User.Trim().ToLowerInvariant()

# ── Logging ────────────────────────────────────────────────────────────────────
$_logDir = 'C:\Users\andyw\OneDrive - Andy White\Contracts\Jolera\Migrations\Logs'
if (-not (Test-Path $_logDir)) { New-Item -ItemType Directory -Path $_logDir -Force | Out-Null }
$logFile = Join-Path $_logDir "domain-team-memberships-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
function Log { param([string]$m) $ts = Get-Date -Format 'HH:mm:ss'; "$ts $m" | Tee-Object -FilePath $logFile -Append | Write-Host }

if (-not $OutputCsv) {
    $OutputCsv = Join-Path (Get-Location) "DomainTeamMemberships-$Domain-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
}
$outDir = Split-Path -Parent $OutputCsv
if ($outDir -and -not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

Log "=== Get Domain Team Memberships ==="
Log "Domain      : $Domain"
if ($User) { Log "User filter : $User" }
Log "Output CSV  : $OutputCsv"
Log "Std chans   : $([bool]$IncludeStandardChannels)"

# ── Graph modules + connection ─────────────────────────────────────────────────
. (Join-Path $PSScriptRoot 'Ensure-GraphModules.ps1') -GraphModules @(
    'Microsoft.Graph.Users', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Teams'
)

$scopes = @(
    'Team.ReadBasic.All', 'TeamMember.Read.All',
    'Channel.ReadBasic.All', 'ChannelMember.Read.All',
    'User.Read.All', 'Group.Read.All'
)

$ctx = Get-MgContext -ErrorAction SilentlyContinue
$haveScopes = $ctx -and -not ($scopes | Where-Object { $_ -notin $ctx.Scopes })
if ($haveScopes -and (-not $TenantId -or $ctx.TenantId -eq $TenantId)) {
    Log "Reusing existing Graph session: $($ctx.Account)  (tenant $($ctx.TenantId))"
} else {
    Log "Connecting to Microsoft Graph — sign in with an admin of the SOURCE tenant when the browser opens..."
    $connect = @{ Scopes = $scopes; NoWelcome = $true; ErrorAction = 'Stop' }
    if ($TenantId) { $connect.TenantId = $TenantId }
    Connect-MgGraph @connect
    $ctx = Get-MgContext
    Log "Connected as $($ctx.Account)  (tenant $($ctx.TenantId))"
}

# ── Small throttling-aware retry wrapper ───────────────────────────────────────
function Invoke-Graph {
    param([scriptblock]$Script, [string]$What = 'Graph call')
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        try { return & $Script }
        catch {
            $status = $_.Exception.Response.StatusCode.value__ 2>$null
            if ($status -in 429, 503, 504 -and $attempt -lt 5) {
                $wait = [math]::Min(60, [math]::Pow(2, $attempt))
                Log "  throttled on $What (HTTP $status) — retry $attempt in ${wait}s"
                Start-Sleep -Seconds $wait
                continue
            }
            throw
        }
    }
}

# ── Build the set of in-domain users ───────────────────────────────────────────
Log "Enumerating users in @$Domain ..."
$domainUsers = @{}   # id -> [pscustomobject] UPN, Mail, DisplayName, UserType

function Add-DomainUser {
    param($u)
    if (-not $u -or -not $u.Id) { return }
    $upn  = ("" + $u.UserPrincipalName).ToLowerInvariant()
    $mail = ("" + $u.Mail).ToLowerInvariant()
    if ($User -and $User -ne $upn -and $User -ne $mail) { return }
    $domainUsers[$u.Id] = [pscustomobject]@{
        DisplayName = $u.DisplayName
        UPN         = $u.UserPrincipalName
        Mail        = $u.Mail
        UserType    = $u.UserType
    }
}

$props = 'Id', 'DisplayName', 'UserPrincipalName', 'Mail', 'UserType'
try {
    Invoke-Graph -What 'Get-MgUser endsWith(UPN)' -Script {
        Get-MgUser -All -Property $props -ConsistencyLevel eventual -CountVariable _c `
            -Filter "endsWith(userPrincipalName,'@$Domain')"
    } | ForEach-Object { Add-DomainUser $_ }

    Invoke-Graph -What 'Get-MgUser endsWith(mail)' -Script {
        Get-MgUser -All -Property $props -ConsistencyLevel eventual -CountVariable _c `
            -Filter "endsWith(mail,'@$Domain')"
    } | ForEach-Object { Add-DomainUser $_ }
}
catch {
    Log "Advanced user filter failed ($($_.Exception.Message.Split([Environment]::NewLine)[0])) — falling back to a full user scan (slower)."
    Invoke-Graph -What 'Get-MgUser -All' -Script { Get-MgUser -All -Property $props } |
        Where-Object {
            ("" + $_.UserPrincipalName).ToLowerInvariant().EndsWith("@$Domain") -or
            ("" + $_.Mail).ToLowerInvariant().EndsWith("@$Domain")
        } | ForEach-Object { Add-DomainUser $_ }
}

Log "Matched $($domainUsers.Count) user(s) in @$Domain."
if ($domainUsers.Count -eq 0) {
    Log "Nothing to report — no users found. Check the domain spelling and that you signed into the correct tenant."
    exit 0
}

# helper: does this Teams member object belong to a target-domain user?
function Resolve-DomainMember {
    param($m)   # aadUserConversationMember from Get-MgTeamMember / Get-MgTeamChannelMember
    $uid   = "" + $m.AdditionalProperties['userId']
    $email = ("" + $m.AdditionalProperties['email']).ToLowerInvariant()
    $hit   = $null
    if ($uid -and $domainUsers.ContainsKey($uid)) { $hit = $domainUsers[$uid] }
    elseif ($email.EndsWith("@$Domain") -and (-not $User -or $User -eq $email)) {
        $hit = [pscustomobject]@{ DisplayName = $m.DisplayName; UPN = ''; Mail = $m.AdditionalProperties['email']; UserType = '' }
    }
    if (-not $hit) { return $null }

    $roles = @($m.Roles)
    $role  = if ($roles -contains 'owner') { 'Owner' }
             elseif ($roles -contains 'guest') { 'Guest' }
             else { 'Member' }
    [pscustomobject]@{
        DisplayName = if ($hit.DisplayName) { $hit.DisplayName } else { $m.DisplayName }
        UPN         = $hit.UPN
        Mail        = $hit.Mail
        UserType    = $hit.UserType
        UserId      = $uid
        Role        = $role
    }
}

# ── Walk every Team ────────────────────────────────────────────────────────────
Log "Enumerating Teams ..."
$teams = @(Invoke-Graph -What 'Get-MgTeam -All' -Script { Get-MgTeam -All -Property 'Id,DisplayName,IsArchived' })
Log "Found $($teams.Count) team(s). Scanning members..."

$rows      = [System.Collections.Generic.List[object]]::new()
$teamIndex = 0

foreach ($team in $teams) {
    $teamIndex++
    if ($teamIndex % 25 -eq 0) { Log "  ...$teamIndex / $($teams.Count) teams" }

    try {
        $members = @(Invoke-Graph -What "Get-MgTeamMember $($team.DisplayName)" -Script {
            Get-MgTeamMember -TeamId $team.Id -All
        })
    }
    catch {
        Log "  ! $($team.DisplayName): could not read members — $($_.Exception.Message.Split([Environment]::NewLine)[0])"
        continue
    }

    $matched = foreach ($m in $members) { Resolve-DomainMember $m }
    $matched = @($matched | Where-Object { $_ })
    if ($matched.Count -eq 0) { continue }

    # only now fetch the group props for mail / visibility
    $grp = $null
    try { $grp = Invoke-Graph -What "Get-MgGroup $($team.Id)" -Script {
        Get-MgGroup -GroupId $team.Id -Property 'Mail,Visibility,DisplayName'
    } } catch { }

    foreach ($u in $matched) {
        $rows.Add([pscustomobject]@{
            TeamDisplayName = $team.DisplayName
            TeamMail        = $grp.Mail
            TeamVisibility  = $grp.Visibility
            TeamArchived    = [bool]$team.IsArchived
            Scope           = 'Team'
            ChannelName     = ''
            ChannelType     = ''
            UserDisplayName = $u.DisplayName
            UserPrincipalName = $u.UPN
            UserMail        = $u.Mail
            UserType        = $u.UserType
            Role            = $u.Role
            TeamId          = $team.Id
            ChannelId       = ''
            UserId          = $u.UserId
        })
    }

    # ── channels ──────────────────────────────────────────────────────────────
    try {
        $channels = @(Invoke-Graph -What "Get-MgTeamChannel $($team.DisplayName)" -Script {
            Get-MgTeamChannel -TeamId $team.Id -All -Property 'Id,DisplayName,MembershipType'
        })
    }
    catch {
        Log "  ! $($team.DisplayName): could not read channels — $($_.Exception.Message.Split([Environment]::NewLine)[0])"
        $channels = @()
    }

    foreach ($ch in $channels) {
        $type = "" + $ch.MembershipType
        if ($type -in 'private', 'shared') {
            try {
                $chMembers = @(Invoke-Graph -What "Get-MgTeamChannelMember $($ch.DisplayName)" -Script {
                    Get-MgTeamChannelMember -TeamId $team.Id -ChannelId $ch.Id -All
                })
            }
            catch {
                Log "  ! $($team.DisplayName) / $($ch.DisplayName): could not read channel members — $($_.Exception.Message.Split([Environment]::NewLine)[0])"
                continue
            }
            $chMatched = @(@(foreach ($m in $chMembers) { Resolve-DomainMember $m }) | Where-Object { $_ })
        }
        elseif ($IncludeStandardChannels) {
            # standard channel membership == team membership
            $chMatched = $matched
        }
        else {
            continue
        }

        foreach ($u in $chMatched) {
            $rows.Add([pscustomobject]@{
                TeamDisplayName = $team.DisplayName
                TeamMail        = $grp.Mail
                TeamVisibility  = $grp.Visibility
                TeamArchived    = [bool]$team.IsArchived
                Scope           = 'Channel'
                ChannelName     = $ch.DisplayName
                ChannelType     = $type
                UserDisplayName = $u.DisplayName
                UserPrincipalName = $u.UPN
                UserMail        = $u.Mail
                UserType        = $u.UserType
                Role            = $u.Role
                TeamId          = $team.Id
                ChannelId       = $ch.Id
                UserId          = $u.UserId
            })
        }
    }
}

# ── Write CSV ─────────────────────────────────────────────────────────────────
$rows = $rows | Sort-Object TeamDisplayName, Scope, ChannelName, UserDisplayName
$rows | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8

$teamCount = ($rows | Select-Object -ExpandProperty TeamId -Unique).Count
$userCount = ($rows | Where-Object { $_.UserId } | Select-Object -ExpandProperty UserId -Unique).Count
Log ""
Log "=== Complete ==="
Log "Rows written : $($rows.Count)"
Log "Distinct teams : $teamCount"
Log "Distinct in-domain users : $userCount"
Log "CSV : $OutputCsv"
Log "Log : $logFile"

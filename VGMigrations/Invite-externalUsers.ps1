#Requires -Version 7.0
<#
.SYNOPSIS
    Invite-externalUsers.ps1 — Invites external people as B2B guests and adds each one as a
    member of one or more Teams / Microsoft 365 Groups, in a single streaming pass.

.DESCRIPTION
    Post-migration helper. When a tenant is migrated, external collaborators who used to have
    access to a Team, its SharePoint site, or an M365 Group lose that access because their
    guest account did not come across. This script re-establishes it:

      1. Invites each person in the CSV as an Entra ID B2B guest (New-MgInvitation). If a guest
         with that email already exists in the tenant it is reused, not re-invited.
      2. Adds the resulting guest directory object as a MEMBER of every target group
         (New-MgGroupMember). A Team's membership IS its underlying M365 group's membership, and
         a group-connected SharePoint site inherits access from that same group — so this one
         action restores Teams, Groups and their SharePoint sites together. "Already a member"
         is treated as a no-op success.

    Targets are resolved once and cached. Each target may be given as the group's mail address,
    its mailNickname, or its object id (GUID).

    Connects to Microsoft Graph (delegated). Sign in with an account that can invite guests and
    manage group membership — Teams Administrator + User Administrator, or Global Administrator.
    An existing Graph session with the required scopes is reused.

.PARAMETER CsvFile
    Path to the input CSV. Columns:
      Email   (required) — the external person's email address
      Name    (optional) — display name for the guest invitation
      Groups  (optional) — per-row target group(s); ';'-separated. Overrides -TargetGroups for
                           that row. Use this when different people go to different groups.
      Message (optional) — per-row custom invitation message body

.PARAMETER TargetGroups
    Default target group(s) applied to every CSV row that has no own 'Groups' value. One or
    more group mail addresses / mailNicknames / object ids, separated by ';' or ','.
    Required unless every row carries its own 'Groups' value.

.PARAMETER TenantId
    Optional tenant id or verified domain to sign into (passed to Connect-MgGraph) when your
    admin account can see more than one tenant.

.PARAMETER SendInvitationEmail
    Actually send Microsoft's "you've been invited" email to each newly-invited guest. Off by
    default — for a batch restore of people who already know they are being re-added, the
    notification is usually just noise.

.PARAMETER InviteRedirectUrl
    Landing page for the guest invitation. Defaults to https://teams.microsoft.com.

.PARAMETER DefaultMessage
    Invitation message body used when a row has no per-row 'Message' value.

.PARAMETER ResultsCsv
    Path for the per-row outcome CSV. Defaults to
    %APPDATA%\FlyMigration\Logs\InviteExternalUsers-<timestamp>.csv

.PARAMETER WhatIf
    Preview every action (guest invite, group add) without making any changes.

.EXAMPLE
    .\Invite-externalUsers.ps1 -CsvFile .\guests.csv -TargetGroups "project-falcon@contoso.com" -WhatIf

.EXAMPLE
    .\Invite-externalUsers.ps1 -CsvFile .\guests.csv `
        -TargetGroups "project-falcon@contoso.com; sales-team@contoso.com" -SendInvitationEmail

.NOTES
    Requires the Microsoft.Graph SDK. Ensure-GraphModules.ps1 pins the submodules to 2.33.0 so
    interactive sign-in uses a plain browser popup.
    Delegated scopes: User.Invite.All, User.Read.All, Group.ReadWrite.All, GroupMember.ReadWrite.All
    Guests are added with the default Member role. Teams does not allow guests to be owners.
    Standalone (non-group-connected) SharePoint sites are not group-managed and are out of
    scope here — grant those directly in the site.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$CsvFile,

    [string]$TargetGroups = '',

    [string]$TenantId = '',

    [switch]$SendInvitationEmail,

    [string]$InviteRedirectUrl = 'https://teams.microsoft.com',

    [string]$DefaultMessage = "You've been invited to keep collaborating with our organisation in Microsoft 365.",

    [string]$ResultsCsv = '',

    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'

function Split-Targets {
    param([string]$Value)
    if (-not $Value) { return @() }
    return @($Value -split '[;,]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

# ── Results CSV location ─────────────────────────────────────────────────────────
if (-not $ResultsCsv) {
    $logDir = Join-Path $env:APPDATA 'FlyMigration\Logs'
    if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    $ResultsCsv = Join-Path $logDir "InviteExternalUsers-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
} else {
    $resultsDir = Split-Path -Parent $ResultsCsv
    if ($resultsDir -and -not (Test-Path $resultsDir)) { New-Item -ItemType Directory -Path $resultsDir -Force | Out-Null }
}

Write-Host "=== Invite External Users$(if ($WhatIf) { ' [WhatIf]' }) ===" -ForegroundColor Cyan
Write-Host "CSV         : $CsvFile"
if ($TargetGroups) { Write-Host "Target group(s): $($(Split-Targets $TargetGroups) -join ', ')" }
Write-Host "Invite email : $([bool]$SendInvitationEmail)"
Write-Host "Results CSV : $ResultsCsv"

# ── Load & validate the CSV ─────────────────────────────────────────────────────
if (-not (Test-Path $CsvFile)) { Write-Error "CSV not found: $CsvFile"; exit 1 }
$rows = @(Import-Csv -Path $CsvFile -Encoding UTF8 -ErrorAction Stop)
if ($rows.Count -eq 0) { Write-Error "CSV is empty."; exit 1 }

$cols = @($rows[0].PSObject.Properties.Name)
Write-Host "Columns     : $($cols -join ', ')"
if ($cols -notcontains 'Email') {
    Write-Error "CSV must have an 'Email' column. Optional: Name, Groups, Message. Found: $($cols -join ', ')"
    exit 1
}

$defaultTargets = Split-Targets $TargetGroups
$rowsWithoutTargets = @($rows | Where-Object {
    -not ($_.PSObject.Properties['Groups'] -and ("" + $_.Groups).Trim())
})
if ($defaultTargets.Count -eq 0 -and $rowsWithoutTargets.Count -gt 0) {
    Write-Error "No target group given. Supply -TargetGroups, or a 'Groups' value on every CSV row."
    exit 1
}

# ── Graph modules + connection ─────────────────────────────────────────────────
. (Join-Path $PSScriptRoot 'Ensure-GraphModules.ps1') -GraphModules @(
    'Microsoft.Graph.Users', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Identity.SignIns'
)

$scopes = @('User.Invite.All', 'User.Read.All', 'Group.ReadWrite.All', 'GroupMember.ReadWrite.All')

$ctx = Get-MgContext -ErrorAction SilentlyContinue
$haveScopes = $ctx -and -not ($scopes | Where-Object { $_ -notin $ctx.Scopes })
if ($haveScopes -and (-not $TenantId -or $ctx.TenantId -eq $TenantId)) {
    Write-Host "Reusing existing Graph session: $($ctx.Account)  (tenant $($ctx.TenantId))" -ForegroundColor Green
} else {
    Write-Host "Connecting to Microsoft Graph — sign in when the browser opens..." -ForegroundColor Yellow
    $connect = @{ Scopes = $scopes; NoWelcome = $true; ErrorAction = 'Stop' }
    if ($TenantId) { $connect.TenantId = $TenantId }
    Connect-MgGraph @connect
    $ctx = Get-MgContext
    Write-Host "Connected as $($ctx.Account)  (tenant $($ctx.TenantId))" -ForegroundColor Green
}

# ── Target group resolution (mail / mailNickname / GUID), cached ───────────────
$groupCache = @{}   # raw target string (lower) -> @{ Id; Name; Label } or $null
$guidRegex  = '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$'
$selectProps = 'id,displayName,mail,mailNickname,groupTypes,resourceProvisioningOptions'

function Resolve-Group {
    param([string]$Target)

    $key = $Target.ToLowerInvariant()
    if ($groupCache.ContainsKey($key)) { return $groupCache[$key] }

    $group = $null
    try {
        if ($Target -match $guidRegex) {
            $group = Get-MgGroup -GroupId $Target -Property $selectProps -ErrorAction Stop
        } else {
            $group = Get-MgGroup -Filter "mail eq '$Target'" -Property $selectProps -ConsistencyLevel eventual -CountVariable c -ErrorAction Stop |
                     Select-Object -First 1
            if (-not $group) {
                $nick = $Target -replace '@.*', ''
                $group = Get-MgGroup -Filter "mailNickname eq '$nick'" -Property $selectProps -ConsistencyLevel eventual -CountVariable c -ErrorAction Stop |
                         Select-Object -First 1
            }
        }
    } catch {
        Write-Warning "  Group lookup failed for '$Target' — $($_.Exception.Message.Split([Environment]::NewLine)[0])"
    }

    if (-not $group) {
        $groupCache[$key] = $null
        return $null
    }

    $isTeam    = $group.AdditionalProperties['resourceProvisioningOptions'] -contains 'Team'
    $isUnified = $group.GroupTypes -contains 'Unified'
    $label = if ($isTeam)         { 'Team' }
             elseif ($isUnified)  { 'M365 Group' }
             else                 { 'Group' }

    $entry = @{ Id = $group.Id; Name = $group.DisplayName; Label = $label }
    $groupCache[$key] = $entry
    return $entry
}

# ── Guest resolution (invite-or-reuse), cached ────────────────────────────────
$guestCache = @{}   # email (lower) -> Entra object id, or $null

function Resolve-Guest {
    param([string]$Email, [string]$Name, [string]$Message)

    $key = $Email.ToLowerInvariant()
    if ($guestCache.ContainsKey($key)) { return $guestCache[$key] }

    try {
        $existing = Get-MgUser -Filter "mail eq '$Email'" -Property Id,Mail -ErrorAction Stop | Select-Object -First 1
        if (-not $existing) {
            $existing = Get-MgUser -Filter "otherMails/any(x:x eq '$Email')" -Property Id,Mail -ConsistencyLevel eventual -CountVariable c -ErrorAction Stop |
                        Select-Object -First 1
        }
        if ($existing) {
            Write-Host "  Guest already exists: $Email" -ForegroundColor DarkGray
            $guestCache[$key] = $existing.Id
            return $existing.Id
        }
    } catch { }

    if ($WhatIf) {
        Write-Host "  WhatIf — would invite guest: $Email" -ForegroundColor Yellow
        $guestCache[$key] = $null
        return $null
    }

    try {
        $params = @{
            InvitedUserEmailAddress = $Email
            InviteRedirectUrl       = $InviteRedirectUrl
            SendInvitationMessage   = [bool]$SendInvitationEmail
            ErrorAction             = 'Stop'
        }
        if ($Name)    { $params.InvitedUserDisplayName = $Name }
        if ($Message) { $params.InvitedUserMessageInfo = @{ CustomizedMessageBody = $Message } }

        $inv = New-MgInvitation @params
        Write-Host "  Invited guest: $Email" -ForegroundColor Green
        $guestCache[$key] = $inv.InvitedUser.Id
        return $inv.InvitedUser.Id
    } catch {
        Write-Warning "  FAILED to invite ${Email}: $($_.Exception.Message.Split([Environment]::NewLine)[0])"
        $guestCache[$key] = $null
        return $null
    }
}

# ── Add-member helper — "already a member" is a no-op success ──────────────────
function Add-GroupMemberSafe {
    param([string]$GroupId, [string]$GuestUserId)

    $ref = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$GuestUserId" }
    try {
        New-MgGroupMember -GroupId $GroupId -BodyParameter $ref -ErrorAction Stop
        return @{ Success = $true; Message = 'Added' }
    } catch {
        $msg = $_.Exception.Message.Split([Environment]::NewLine)[0]
        if ($msg -match 'already exist|One or more added object references|Conflict|duplicate') {
            return @{ Success = $true; Message = 'Already a member' }
        }
        return @{ Success = $false; Message = $msg }
    }
}

# ── Process ───────────────────────────────────────────────────────────────────
$results = [System.Collections.Generic.List[object]]::new()
$ok = 0; $fail = 0; $skip = 0
$total = $rows.Count
$i = 0

foreach ($row in $rows) {
    $i++
    $email = ("" + $row.Email).Trim()
    $name  = if ($row.PSObject.Properties['Name'])  { ("" + $row.Name).Trim() }  else { '' }
    $msg   = if ($row.PSObject.Properties['Message'] -and ("" + $row.Message).Trim()) { $row.Message } else { $DefaultMessage }

    if (-not $email) {
        Write-Warning "[$i/$total] Skipped — row has no Email"
        $results.Add([pscustomobject]@{ Email=''; Name=$name; Group=''; Result='Skipped'; Message='Missing email' })
        $skip++
        continue
    }

    $rowTargets = if ($row.PSObject.Properties['Groups'] -and ("" + $row.Groups).Trim()) {
        Split-Targets $row.Groups
    } else {
        $defaultTargets
    }
    if ($rowTargets.Count -eq 0) {
        Write-Warning "[$i/$total] $email — Skipped: no target group for this row"
        $results.Add([pscustomobject]@{ Email=$email; Name=$name; Group=''; Result='Skipped'; Message='No target group' })
        $skip++
        continue
    }

    Write-Host "[$i/$total] $email" -ForegroundColor Cyan
    $guestId = Resolve-Guest -Email $email -Name $name -Message $msg

    if (-not $guestId -and -not $WhatIf) {
        foreach ($t in $rowTargets) {
            $results.Add([pscustomobject]@{ Email=$email; Name=$name; Group=$t; Result='Failed'; Message='Guest invite failed' })
        }
        $fail++
        continue
    }

    foreach ($t in $rowTargets) {
        $g = Resolve-Group -Target $t
        if (-not $g) {
            Write-Warning "    Target not found: $t"
            $results.Add([pscustomobject]@{ Email=$email; Name=$name; Group=$t; Result='Failed'; Message='Group not found' })
            $fail++
            continue
        }

        if ($WhatIf) {
            Write-Host "    WhatIf — would add to [$($g.Label)] $($g.Name)" -ForegroundColor Yellow
            $results.Add([pscustomobject]@{ Email=$email; Name=$name; Group=$g.Name; Result='WhatIf'; Message="Would add to $($g.Label)" })
            $ok++
            continue
        }

        $add = Add-GroupMemberSafe -GroupId $g.Id -GuestUserId $guestId
        if ($add.Success) {
            Write-Host "    $($add.Message): [$($g.Label)] $($g.Name)" -ForegroundColor Green
            $results.Add([pscustomobject]@{ Email=$email; Name=$name; Group=$g.Name; Result='OK'; Message=$add.Message })
            $ok++
        } else {
            Write-Warning "    FAILED [$($g.Label)] $($g.Name) — $($add.Message)"
            $results.Add([pscustomobject]@{ Email=$email; Name=$name; Group=$g.Name; Result='Failed'; Message=$add.Message })
            $fail++
        }
    }
}

$results | Export-Csv -Path $ResultsCsv -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "=== Complete: ok $ok  |  failed $fail  |  skipped $skip ===" -ForegroundColor $(if ($fail -gt 0) { 'Yellow' } else { 'Green' })
Write-Host "Results CSV : $ResultsCsv"
if ($fail -gt 0) { Write-Warning "$fail add(s) did not succeed — see the results CSV."; exit 1 }

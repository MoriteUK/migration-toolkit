#Requires -Version 7.0
<#
.SYNOPSIS
    Invite-externalUsers.ps1 — Invites external people as B2B guests and adds each one as a
    member of one or more Teams / Microsoft 365 Groups, in a single streaming pass.

.DESCRIPTION
    Post-migration helper. When a tenant is migrated, external collaborators who used to have
    access to a Team, its SharePoint site, or an M365 Group lose that access because their
    guest account did not come across. This script re-establishes it:

      1. Invites each person in the CSV as an Entra ID B2B guest (POST /invitations). If a guest
         with that email already exists in the tenant it is reused, not re-invited.
      2. Adds the resulting guest directory object as a MEMBER of every target group
         (POST /groups/{id}/members/$ref). A Team's membership IS its underlying M365 group's
         membership, and a group-connected SharePoint site inherits access from that same group
         — so this one action restores Teams, Groups and their SharePoint sites together.
         "Already a member" is treated as a no-op success.

    Every Graph call goes through Invoke-MgGraphRequest (raw REST) rather than the typed
    Microsoft.Graph.Users / .Groups / .Identity.SignIns cmdlets. Those submodules each carry
    their own copy of Microsoft.Graph.Authentication; loading more than one version into the
    same process leaves Connect-MgGraph's token in one copy and the cmdlet calls reading an
    empty context in another, which shows up as "InteractiveBrowserCredential authentication
    failed" on the first real call. Invoke-MgGraphRequest lives in Authentication itself and
    always sees the live context, so only that one module has to load.

    Target groups are resolved once, up front. If the resolve can't authenticate or is denied,
    the run stops immediately with the real error — it does not fall through to a misleading
    "group not found" for every row.

.PARAMETER CsvFile
    Path to the input CSV. Two formats are accepted and auto-detected:

    1. Simple list:
      Email   (required) — the external person's email address
      Name    (optional) — display name for the guest invitation
      Groups  (optional) — per-row target group(s); ';'-separated. Overrides -TargetGroups for
                           that row. Use this when different people go to different groups.
      Message (optional) — per-row custom invitation message body

    2. An AvePoint Fly "M365 Group objects" migration report (columns include Title, Type,
      Source, Status, Error code). The failed 'CO-UserOrGroupNotFound' member rows are the
      external people who need a guest account created and adding to that group in the
      destination. 'Title' is the member's email; the group is the part of 'Source' before the
      first '/'. Rows on an @*.onmicrosoft.com / destination-tenant address (internal accounts)
      and non-membership rows are skipped. -TargetGroups is not needed with this format — each
      row already carries its own group.

.PARAMETER IncludeAllErrorCodes
    (Fly report format only) Also process rows whose Error code is something other than
    CO-UserOrGroupNotFound (e.g. CO-MatchMultipleUser). Off by default — those usually need a
    user mapping in the migration policy, not a guest invite.

.PARAMETER TargetGroups
    Default target group(s) applied to every CSV row that has no own 'Groups' value. One or
    more group mail addresses / mailNicknames / display names / object ids, separated by ';'
    or ','. Required unless every row carries its own 'Groups' value.

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
    Preview every action (guest invite, group add) without making any changes. Target groups
    are still resolved live so the preview is real.

.EXAMPLE
    .\Invite-externalUsers.ps1 -CsvFile .\guests.csv -TargetGroups "project-falcon@contoso.com" -WhatIf

.EXAMPLE
    .\Invite-externalUsers.ps1 -CsvFile .\guests.csv `
        -TargetGroups "project-falcon@contoso.com; sales-team@contoso.com" -SendInvitationEmail

.NOTES
    Ensure-GraphModules.ps1 pins Microsoft.Graph.Authentication to 2.33.0 so interactive
    sign-in uses a plain browser popup.
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

    [switch]$IncludeAllErrorCodes,

    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'
$GraphBase = 'https://graph.microsoft.com/v1.0'

function Split-Targets {
    param([string]$Value)
    if (-not $Value) { return @() }
    return @($Value -split '[;,]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

# OData string literal - single quotes doubled, spaces %20-encoded so a raw URI is safe.
function ConvertTo-ODataLiteral {
    param([string]$Value)
    return ($Value -replace "'", "''") -replace ' ', '%20'
}

function Test-IsAuthError {
    param([string]$Message)
    return $Message -match 'InteractiveBrowserCredential|authentication failed|AADSTS|InvalidAuthenticationToken|token has expired|Lifetime validation failed|Authorization_RequestDenied|Insufficient privileges|Access is denied|Forbidden|401|403'
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

# Rows the Fly-report parser decides upfront it will not attempt (internal accounts, wrong
# error code) - folded into the results CSV at the end so nothing silently disappears.
$flyPrescreen = [System.Collections.Generic.List[object]]::new()

# ── AvePoint Fly "M365 Group objects" migration report? ───────────────────────
# Export columns: Migration start time, Sub job ID, Title, Type, Source, Destination, Size,
# Status, Migration action, Comment, Error code.
#   Title  = the member's email address
#   Type   = Member | Owner
#   Source = "<group mail-or-name>/<member>"   (group id = everything before the first '/')
#   Status / Error code = why the member failed to migrate
# The failed CO-UserOrGroupNotFound rows are exactly the external people who need a guest
# account created and adding to that group in the destination. Normalise them into the same
# per-row { Email; Groups } shape the rest of the script already handles.
$isFlyReport = ($cols -contains 'Title') -and ($cols -contains 'Source') -and ($cols -contains 'Type')
if ($isFlyReport) {
    Write-Host "Detected an AvePoint Fly M365 Group migration report — extracting members to re-add..." -ForegroundColor Cyan

    # An @*.onmicrosoft.com address, or one on the tenant given in -TenantId, is an internal /
    # unmigrated account, never an external guest.
    $skipDomainRx = '(?i)\.onmicrosoft\.com$'
    if ($TenantId -and $TenantId -notmatch '^[0-9a-fA-F-]{36}$') {
        $skipDomainRx = "(?i)(\.onmicrosoft\.com|@$([regex]::Escape($TenantId)))$"
    }

    $seen = [System.Collections.Generic.HashSet[string]]::new()
    $norm = [System.Collections.Generic.List[object]]::new()
    foreach ($r in $rows) {
        $src = ("" + $r.Source).Trim()
        if ($src -notmatch '/') { continue }                       # group-level row, not a membership row
        $slash = $src.IndexOf('/')
        $group = $src.Substring(0, $slash).Trim()
        $email = ("" + $r.Title).Trim()
        if (-not $email) { $email = $src.Substring($slash + 1).Trim() }
        if (-not $email -or -not $group) { continue }

        $stat = ("" + $r.Status).Trim()
        if ($stat -and $stat -notmatch '(?i)error|fail') { continue }   # a member that migrated fine needs nothing

        $key = ($email + '|' + $group).ToLowerInvariant()
        if (-not $seen.Add($key)) { continue }                     # de-dupe retried sub-jobs

        $role = ("" + $r.Type).Trim()
        $code = ("" + $r.'Error code').Trim()

        if ($email -match $skipDomainRx) {
            $flyPrescreen.Add([pscustomobject]@{ Email=$email; Name=''; Group=$group; Result='Skipped'; Message='Internal / tenant account — not an external guest' })
            continue
        }
        if ($code -and $code -ne 'CO-UserOrGroupNotFound' -and -not $IncludeAllErrorCodes) {
            $flyPrescreen.Add([pscustomobject]@{ Email=$email; Name=''; Group=$group; Result='Skipped'; Message="Fly error '$code' — needs a user mapping in the migration policy, not a guest invite (use -IncludeAllErrorCodes to override)" })
            continue
        }

        $norm.Add([pscustomobject]@{ Email=$email; Name=''; Groups=$group; Message=''; SourceRole=$role })
    }

    if ($norm.Count -eq 0) {
        Write-Error "No re-addable external members found in the Fly report (after removing internal accounts, non-membership rows and non-'not found' errors)."
        if ($flyPrescreen.Count) { $flyPrescreen | Export-Csv -Path $ResultsCsv -NoTypeInformation -Encoding UTF8; Write-Host "Skipped rows written to $ResultsCsv" -ForegroundColor Yellow }
        exit 1
    }
    Write-Host ("  {0} unique (member, group) pair(s) to process  |  {1} row(s) pre-skipped" -f $norm.Count, $flyPrescreen.Count) -ForegroundColor Cyan
    $rows = $norm.ToArray()
    $cols = @('Email', 'Name', 'Groups', 'Message')
}

if ($cols -notcontains 'Email') {
    Write-Error "CSV must have an 'Email' column, or be an AvePoint Fly M365 Group migration report. Optional: Name, Groups, Message. Found: $($cols -join ', ')"
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

# ── Graph module (Authentication only) + connection ───────────────────────────
# Only Microsoft.Graph.Authentication is loaded. Invoke-MgGraphRequest below covers every call,
# so the .Users / .Groups / .Identity.SignIns submodules (and their duplicate copies of the
# Authentication assembly) are never brought into the process.
. (Join-Path $PSScriptRoot 'Ensure-GraphModules.ps1') -GraphModules @()

$scopes = @('User.Invite.All', 'User.Read.All', 'Group.ReadWrite.All', 'GroupMember.ReadWrite.All')

# Cheap authenticated probe - proves the session actually has a usable access token, not just a
# recorded account. Connect-MgGraph can report "Connected as ..." while the token acquisition
# silently failed; the failure then only surfaces on the first real request as
# "InteractiveBrowserCredential authentication failed".
function Test-GraphToken {
    try {
        Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization?$select=id' `
            -OutputType PSObject -ErrorAction Stop | Out-Null
        return $true
    } catch {
        $m = $_.Exception.Message.Split([Environment]::NewLine)[0]
        if (Test-IsAuthError $m) { return $false }
        return $true   # non-auth error from the probe endpoint - the token itself is fine
    }
}

$ctx = Get-MgContext -ErrorAction SilentlyContinue
$haveScopes = $ctx -and -not ($scopes | Where-Object { $_ -notin $ctx.Scopes })
if ($haveScopes -and (Test-GraphToken)) {
    Write-Host "Reusing existing Graph session: $($ctx.Account)  (tenant $($ctx.TenantId))" -ForegroundColor Green
} else {
    # Drop any recorded account / cached context first. Without this, Connect-MgGraph finds a
    # stale account in the shared MSAL token store (left by an earlier run or another tool),
    # tries a SILENT token acquisition against it, that fails quietly - and because it thinks it
    # already has an account it never falls back to interactive, so no browser ever opens and
    # the token probe then fails. -ContextScope Process keeps this run off the shared store
    # entirely so every launch is a clean, interactive sign-in. This is the same pattern
    # Check-OneDriveStatus.ps1 / Provision-OneDrives.ps1 use (v2.9.85).
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}

    $connect = @{ Scopes = $scopes; NoWelcome = $true; ErrorAction = 'Stop'; ContextScope = 'Process' }
    # With a process-scoped context there is no shared cache for -TenantId to collide with, so
    # honour it when given - it targets the right tenant up front instead of only being checked
    # afterwards.
    if ($TenantId) { $connect.TenantId = $TenantId }

    Write-Host "Connecting to Microsoft Graph — a browser sign-in window will open, complete the sign-in there..." -ForegroundColor Yellow
    Connect-MgGraph @connect

    if (-not (Test-GraphToken)) {
        # Device-code sign-in is deliberately NOT attempted - it's blocked tenant-wide by a
        # Conditional Access 'Authentication flows' policy on the tenants this is used against.
        Write-Error @"
Signed-in account was recorded but no usable Microsoft Graph token was obtained.
  - If a browser sign-in window opened: complete it (pick the admin account, approve any
    consent prompt), then run this again.
  - If NO browser opened: a stale cached sign-in is being reused. Run  Disconnect-MgGraph
    (or just close this window and click 'Invite & Add' again), then complete the browser
    sign-in when it appears.
Nothing was changed.
"@
        exit 1
    }
    $ctx = Get-MgContext
    Write-Host "Connected as $($ctx.Account)  (tenant $($ctx.TenantId))" -ForegroundColor Green
}

if ($TenantId -and $ctx.TenantId -and
    $ctx.TenantId -ne $TenantId -and ("$($ctx.Account)" -notlike "*@$TenantId")) {
    Write-Warning "Signed into tenant '$($ctx.TenantId)' as $($ctx.Account), but -TenantId was '$TenantId'. Continuing against the signed-in tenant."
}

$grantedScopes = @($ctx.Scopes)
Write-Host "Granted scopes: $($grantedScopes -join ', ')" -ForegroundColor DarkGray
if ('Group.Read.All' -notin $grantedScopes -and 'Group.ReadWrite.All' -notin $grantedScopes) {
    Write-Warning "The sign-in did not grant Group.Read.All / Group.ReadWrite.All — group lookups will be denied. Re-run and complete the admin-consent prompt, or have a Global Admin consent for the app."
}

# ── Thin Graph REST wrapper ───────────────────────────────────────────────────
function Invoke-Graph {
    param(
        [string]$Method = 'GET',
        [Parameter(Mandatory)][string]$Uri,
        $Body,
        [switch]$AdvancedQuery   # adds ConsistencyLevel: eventual for $filter/$count/$search
    )
    $p = @{ Method = $Method; Uri = $Uri; OutputType = 'PSObject'; ErrorAction = 'Stop' }
    if ($PSBoundParameters.ContainsKey('Body') -and $null -ne $Body) {
        $p.Body        = ($Body | ConvertTo-Json -Depth 6 -Compress)
        $p.ContentType = 'application/json'
    }
    if ($AdvancedQuery) { $p.Headers = @{ ConsistencyLevel = 'eventual' } }
    return Invoke-MgGraphRequest @p
}

# ── Target group resolution (mail / mailNickname / displayName / GUID), cached ──
$groupCache = @{}   # raw target string (lower) -> @{ Id; Name; Label } or $null
$guidRegex  = '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$'
$grpSelect  = 'id,displayName,mail,mailNickname,groupTypes,resourceProvisioningOptions'

# Throws on auth/permission errors (caught once at the top-level resolve pass so the whole run
# stops with a real message); returns $null only for a genuine clean no-match.
function Resolve-Group {
    param([string]$Target)

    $key = $Target.ToLowerInvariant()
    if ($groupCache.ContainsKey($key)) { return $groupCache[$key] }

    $group = $null
    try {
        if ($Target -match $guidRegex) {
            $group = Invoke-Graph -Uri "$GraphBase/groups/$Target`?`$select=$grpSelect"
        } else {
            $bare = $Target -replace '@.*', ''          # strip any @domain the caller appended
            $lit  = ConvertTo-ODataLiteral $Target
            $nick = ConvertTo-ODataLiteral $bare
            # Try the exact-match filters, both with and without the @domain, since a plain
            # security group (no mail) is commonly referenced by the address form of its name
            # but actually only matches on displayName / mailNickname of the bare label.
            $clauses = @("mail eq '$lit'", "mailNickname eq '$nick'",
                         "displayName eq '$nick'", "displayName eq '$lit'") | Select-Object -Unique
            foreach ($clause in $clauses) {
                $enc = $clause -replace ' ', '%20'
                $r = Invoke-Graph -AdvancedQuery -Uri "$GraphBase/groups?`$filter=$enc&`$select=$grpSelect&`$count=true&`$top=2"
                if ($r.value -and $r.value.Count -gt 0) { $group = $r.value[0]; break }
            }
            # Last resort: a substring search on displayName. Take it only if it's unambiguous;
            # otherwise list the candidates so the caller can pass an exact name or the id.
            if (-not $group) {
                $sTerm = ConvertTo-ODataLiteral $bare
                $sr = Invoke-Graph -AdvancedQuery -Uri "$GraphBase/groups?`$search=%22displayName:$sTerm%22&`$select=$grpSelect&`$count=true&`$top=5"
                $cand = @($sr.value)
                if ($cand.Count -eq 1) {
                    $group = $cand[0]
                    Write-Host "  (matched '$Target' by displayName search -> '$($group.displayName)')" -ForegroundColor DarkGray
                } elseif ($cand.Count -gt 1) {
                    Write-Warning "  '$Target' is ambiguous - $($cand.Count) groups match. Pass an exact displayName or the group id:"
                    foreach ($c in $cand) { Write-Warning "      $($c.displayName)   $($c.id)" }
                }
            }
        }
    } catch {
        $m = $_.Exception.Message.Split([Environment]::NewLine)[0]
        if (Test-IsAuthError $m) {
            throw "Graph rejected the group lookup for '$Target': $m"
        }
        Write-Warning "  Group lookup error for '$Target' — $m"
    }

    if (-not $group) {
        $groupCache[$key] = $null
        return $null
    }

    $rpo       = @($group.resourceProvisioningOptions)
    $gt        = @($group.groupTypes)
    $isTeam    = $rpo -contains 'Team'
    $isUnified = $gt  -contains 'Unified'
    $label = if ($isTeam)         { 'Team' }
             elseif ($isUnified)  { 'M365 Group' }
             else                 { 'Group' }

    $entry = @{ Id = $group.id; Name = $group.displayName; Label = $label }
    $groupCache[$key] = $entry
    return $entry
}

# ── Resolve every target up front - fail fast on auth/permission problems ──────
$allTargets = [System.Collections.Generic.List[string]]::new()
foreach ($t in $defaultTargets) { if ($t -notin $allTargets) { $allTargets.Add($t) } }
foreach ($row in $rows) {
    if ($row.PSObject.Properties['Groups'] -and ("" + $row.Groups).Trim()) {
        foreach ($t in (Split-Targets $row.Groups)) { if ($t -notin $allTargets) { $allTargets.Add($t) } }
    }
}

Write-Host ""
Write-Host "Resolving $($allTargets.Count) target group(s)..." -ForegroundColor Cyan
$unresolved = [System.Collections.Generic.List[string]]::new()
try {
    foreach ($t in $allTargets) {
        $g = Resolve-Group -Target $t
        if ($g) {
            Write-Host "  OK  $t  ->  [$($g.Label)] $($g.Name)  ($($g.Id))" -ForegroundColor Green
        } else {
            Write-Warning "  NOT FOUND  $t  — no group in this tenant matched on mail / mailNickname / displayName (with or without @domain) / id. Try the group's object id from Entra."
            $unresolved.Add($t)
        }
    }
} catch {
    Write-Host ""
    Write-Error $_.Exception.Message
    Write-Host "Nothing was changed. Re-run once the sign-in / permissions are sorted." -ForegroundColor Yellow
    exit 1
}
if ($unresolved.Count -eq $allTargets.Count) {
    Write-Error "None of the target group(s) could be resolved in tenant '$($ctx.TenantId)'. Check the address/name is exactly as it appears in this tenant."
    exit 1
}

# ── Guest resolution (invite-or-reuse), cached ────────────────────────────────
$guestCache = @{}   # email (lower) -> Entra object id, or $null
$script:LastGuestError = $null   # real reason the last Resolve-Guest returned $null (-> results CSV)

function Resolve-Guest {
    param([string]$Email, [string]$Name, [string]$Message)

    $script:LastGuestError = $null
    $key = $Email.ToLowerInvariant()
    if ($guestCache.ContainsKey($key)) { return $guestCache[$key] }

    # Reject an obviously malformed address before Graph does. The 'Error_*' CSV exports seen in
    # the field have the Email column mangled with '_' where '@' should be
    # (e.g. luca.prada_alten.it), which fails every invite with a generic error.
    if ($Email -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
        $hint = if ($Email -notmatch '@' -and $Email -match '_') {
            " — looks like '@' was replaced with '_'; did you mean '$($Email -replace '_', '@')'?"
        } else { '' }
        $script:LastGuestError = "'$Email' is not a valid email address$hint"
        Write-Warning "  $($script:LastGuestError)"
        $guestCache[$key] = $null
        return $null
    }

    $lit = ConvertTo-ODataLiteral $Email
    try {
        $r = Invoke-Graph -AdvancedQuery -Uri "$GraphBase/users?`$filter=mail%20eq%20'$lit'%20or%20otherMails/any(x:x%20eq%20'$lit')&`$select=id,mail&`$count=true&`$top=2"
        if ($r.value -and $r.value.Count -gt 0) {
            Write-Host "  Guest already exists: $Email" -ForegroundColor DarkGray
            $guestCache[$key] = $r.value[0].id
            return $r.value[0].id
        }
    } catch {
        $m = $_.Exception.Message.Split([Environment]::NewLine)[0]
        if (Test-IsAuthError $m) { throw "Graph rejected the user lookup for '$Email': $m" }
        Write-Warning "  User lookup error for ${Email}: $m"
    }

    if ($WhatIf) {
        Write-Host "  WhatIf — would invite guest: $Email" -ForegroundColor Yellow
        $guestCache[$key] = $null
        return $null
    }

    try {
        $body = @{
            invitedUserEmailAddress = $Email
            inviteRedirectUrl       = $InviteRedirectUrl
            sendInvitationMessage   = [bool]$SendInvitationEmail
        }
        if ($Name)    { $body.invitedUserDisplayName = $Name }
        if ($Message) { $body.invitedUserMessageInfo = @{ customizedMessageBody = $Message } }

        $inv = Invoke-Graph -Method POST -Uri "$GraphBase/invitations" -Body $body
        Write-Host "  Invited guest: $Email" -ForegroundColor Green
        $guestCache[$key] = $inv.invitedUser.id
        return $inv.invitedUser.id
    } catch {
        $script:LastGuestError = $_.Exception.Message.Split([Environment]::NewLine)[0]
        Write-Warning "  FAILED to invite ${Email}: $($script:LastGuestError)"
        $guestCache[$key] = $null
        return $null
    }
}

# ── Add-member helper — "already a member" is a no-op success ──────────────────
function Add-GroupMemberSafe {
    param([string]$GroupId, [string]$GuestUserId)

    $body = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$GuestUserId" }
    try {
        Invoke-Graph -Method POST -Uri "$GraphBase/groups/$GroupId/members/`$ref" -Body $body | Out-Null
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

# Carry through the rows the Fly-report parser already decided to skip.
foreach ($p in $flyPrescreen) { $results.Add($p); $skip++ }
if ($flyPrescreen.Count) { Write-Host "$($flyPrescreen.Count) row(s) pre-skipped from the Fly report (see results CSV)" -ForegroundColor DarkGray }

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

    $srcRole = if ($row.PSObject.Properties['SourceRole']) { ("" + $row.SourceRole).Trim() } else { '' }
    Write-Host "[$i/$total] $email$(if ($srcRole) { "  (was $srcRole in source)" })" -ForegroundColor Cyan
    if ($srcRole -eq 'Owner') {
        Write-Host "    note: guests can't be group owners — re-adding as a Member" -ForegroundColor DarkYellow
    }
    try {
        $guestId = Resolve-Guest -Email $email -Name $name -Message $msg
    } catch {
        Write-Host ""
        Write-Error $_.Exception.Message
        Write-Host "Stopped after $ok add(s). Re-run once the sign-in / permissions are sorted." -ForegroundColor Yellow
        $results | Export-Csv -Path $ResultsCsv -NoTypeInformation -Encoding UTF8
        exit 1
    }

    # A null guest id is a hard fail EXCEPT under WhatIf, where it's expected (no invite is made)
    # — unless the address itself is invalid, which WhatIf should still flag rather than pretend
    # it would work.
    $guestBlocked = (-not $guestId) -and ((-not $WhatIf) -or $script:LastGuestError)
    if ($guestBlocked) {
        $why = if ($script:LastGuestError) { $script:LastGuestError } else { 'Guest invite failed' }
        foreach ($t in $rowTargets) {
            $results.Add([pscustomobject]@{ Email=$email; Name=$name; Group=$t; Result='Failed'; Message=$why })
        }
        $fail++
        continue
    }

    foreach ($t in $rowTargets) {
        $g = Resolve-Group -Target $t
        if (-not $g) {
            Write-Warning "    Target not resolvable: $t"
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

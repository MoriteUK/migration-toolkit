#Requires -Version 7.0
<#
.SYNOPSIS
    Restore-DomainTeamMemberships.ps1 — Re-adds migrated users, as guests, to the Teams (and
    private/shared channels) captured in the -MembershipCsv that are NOT moving with them —
    i.e. Teams that will still live in the SOURCE tenant after cutover.

.DESCRIPTION
    -MembershipCsv captures, from the source tenant BEFORE migration, every Team/channel a
    domain's users belonged to — either the Domain Team Memberships step that's now part of
    every Run-Assessment.ps1 assessment (DomainTeamMemberships-<domain>-<timestamp>.csv at the
    top of the assessment folder), or the older standalone Get-DomainTeamMemberships.ps1. Same
    CSV schema either way. Once that domain's accounts move to the destination tenant, they
    silently lose membership in every Team that stays behind in the source tenant. This script
    re-adds them — as B2B guests, since their account no longer lives in the source tenant —
    to just those Teams.

    "Still in the source tenant" is inferred from the CSV's TeamMail column: rows for a Team
    whose own mailbox address is on the domain being migrated are skipped (that Team is
    presumably migrating too, via the normal Fly workflow, and needs no manual restore here).
    Rows for a Team on any OTHER domain are treated as staying and are restored. Archived Teams
    and Standard-channel rows (their membership always mirrors the parent team - no separate
    action needed) are always skipped.

    Each user's NEW (destination-tenant) address is required to invite them back as a guest.
    Resolved in this order:
      1. -MappingCsv — the Fly "mapping exchange" file (xlsx or csv, Source/Destination
         columns) this toolkit's Assessment engine already produces. The authoritative source —
         reflects what the real mailbox migration actually used, aliases and all.
      2. -CustomerPrefix — falls back to <local-part>@<destination tenant domain>, the same
         Settings > Customers AccountName lookup Restore-ProxyAddresses.ps1 uses. Only reliable
         before any post-migration renames; prefer -MappingCsv when you have it.
    A user with no resolvable new address is skipped and logged, not guessed at.

    Role handling: Teams does not allow guests to be Team owners. A row with Role=Owner is
    added as a regular member instead, with a WARN logged — ownership itself can't be
    preserved for a guest account, only membership/access.

    Guest resolution is cached per run: if a guest with the target email already exists in the
    source tenant (a prior run, or invited some other way), it's reused rather than re-invited.
    Membership calls that fail because the person is already a member of that Team/channel are
    treated as a no-op success, not a failure.

    Connects with Microsoft Graph (delegated) to the SOURCE tenant — sign in with an account
    that can invite guests and manage Team/channel membership (Teams Administrator + User
    Administrator, or Global Administrator).

.PARAMETER MembershipCsv
    Path to the DomainTeamMemberships-<domain>-<timestamp>.csv written by Run-Assessment.ps1
    (unless -SkipTeamMemberships was set) or by the standalone Get-DomainTeamMemberships.ps1.

.PARAMETER Domain
    The domain that migrated away — same value passed to Get-DomainTeamMemberships.ps1. Used
    to exclude rows for Teams whose own TeamMail is on this domain (those are presumed to be
    migrating too).

.PARAMETER MappingCsv
    Optional. The Fly exchange mapping file (Source/Destination columns) used to resolve each
    user's new destination-tenant address. Accepts .xlsx (via ImportExcel) or .csv.

.PARAMETER CustomerPrefix
    Optional. Settings > Customers Prefix (e.g. "AEP") — used as a fallback when a user isn't
    found in -MappingCsv (or it wasn't supplied). Derives <local-part>@<tenant domain> from
    that customer's AccountName in %APPDATA%\FlyMigration\config.json.

.PARAMETER TenantId
    Optional tenant id / domain to sign into (passed to Connect-MgGraph) if your admin account
    can see more than one tenant.

.PARAMETER SendInvitationEmail
    Actually emails each newly-invited guest Microsoft's "you've been invited" notice. Off by
    default — for a batch restore of colleagues who already know they're being re-added, the
    notification is usually just noise; turn it on if you want them to get it.

.PARAMETER InviteRedirectUrl
    Landing page for the guest invitation. Defaults to https://teams.microsoft.com.

.PARAMETER ResultsCsv
    Path for the per-row outcome CSV. Defaults alongside the shared toolkit reports folder.

.PARAMETER WhatIf
    Preview every action (guest invite, team/channel add) without making any changes.

.EXAMPLE
    .\Restore-DomainTeamMemberships.ps1 -MembershipCsv .\DomainTeamMemberships-aep-italia.it-*.csv `
        -Domain aep-italia.it -MappingCsv ".\AEP mapping exchange.xlsx" -WhatIf

.EXAMPLE
    .\Restore-DomainTeamMemberships.ps1 -MembershipCsv .\DomainTeamMemberships-aep-italia.it-*.csv `
        -Domain aep-italia.it -CustomerPrefix AEP

.NOTES
    Not yet exercised against a live tenant — no Graph connection was available while writing
    this. Run with -WhatIf first and read the results CSV closely before trusting a live run,
    especially the Owner-downgrade and guest-resolution logic.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$MembershipCsv,

    [Parameter(Mandatory = $true)]
    [string]$Domain,

    [string]$MappingCsv = '',
    [string]$CustomerPrefix = '',
    [string]$TenantId = '',
    [switch]$SendInvitationEmail,
    [string]$InviteRedirectUrl = 'https://teams.microsoft.com',
    [string]$ResultsCsv = '',
    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'
$Domain = $Domain.Trim().TrimStart('@').ToLowerInvariant()

# ── Logging ──────────────────────────────────────────────────────────────────────
$_logDir = 'C:\Users\andyw\OneDrive - Andy White\Contracts\Jolera\Migrations\Logs'
if (-not (Test-Path $_logDir)) { New-Item -ItemType Directory -Path $_logDir -Force | Out-Null }
$logFile = Join-Path $_logDir "restore-team-memberships-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
function Log { param([string]$m) $ts = Get-Date -Format 'HH:mm:ss'; "$ts $m" | Tee-Object -FilePath $logFile -Append | Write-Host }

if (-not $ResultsCsv) {
    $ResultsCsv = Join-Path 'C:\Users\andyw\OneDrive - Andy White\Contracts\Jolera\Migrations\Reports' `
        "RestoreTeamMemberships-$Domain-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
}
$resultsDir = Split-Path -Parent $ResultsCsv
if ($resultsDir -and -not (Test-Path $resultsDir)) { New-Item -ItemType Directory -Path $resultsDir -Force | Out-Null }

Log "=== Restore Domain Team Memberships$(if ($WhatIf) { ' [WhatIf]' }) ==="
Log "Membership CSV : $MembershipCsv"
Log "Domain (moved) : $Domain"
if ($MappingCsv)     { Log "Mapping file   : $MappingCsv" }
if ($CustomerPrefix) { Log "Customer       : $CustomerPrefix" }
Log "Results CSV    : $ResultsCsv"

# ── Load the membership CSV ────────────────────────────────────────────────────────
if (-not (Test-Path $MembershipCsv)) {
    Log "ERROR: Membership CSV not found: $MembershipCsv"
    exit 1
}
$allRows = @(Import-Csv -Path $MembershipCsv -Encoding UTF8)
if ($allRows.Count -eq 0) {
    Log "Membership CSV is empty — nothing to restore."
    exit 0
}

# Teams staying in the source tenant: TeamMail is NOT on the migrated domain. Also drop
# archived Teams (can't add members) and Standard channel rows (membership auto-follows the
# team — no separate action, and re-adding a Standard row here would be a duplicate no-op at
# best against the parent Team row already in the CSV).
$rows = @($allRows | Where-Object {
    $_.TeamArchived -ne 'True' -and
    $_.ChannelType  -ne 'standard' -and
    -not ("" + $_.TeamMail).ToLowerInvariant().EndsWith("@$Domain")
})

$skippedMigrating = $allRows.Count - $rows.Count
Log "Loaded $($allRows.Count) row(s); $skippedMigrating skipped (migrating team / archived / standard channel); $($rows.Count) to process"
if ($rows.Count -eq 0) {
    Log "Nothing to restore."
    exit 0
}

# ── Resolve each old identity's new (destination-tenant) address ──────────────────────
$mappingTable = @{}   # old address (lower) -> new address
if ($MappingCsv) {
    if (-not (Test-Path $MappingCsv)) {
        Log "WARNING: -MappingCsv not found: $MappingCsv — continuing without it"
    } else {
        try {
            $mapRows = if ($MappingCsv -match '\.xlsx$') {
                Import-Module ImportExcel -DisableNameChecking -ErrorAction Stop
                @(Import-Excel -Path $MappingCsv -ErrorAction Stop)
            } else {
                @(Import-Csv -Path $MappingCsv -Encoding UTF8 -ErrorAction Stop)
            }
            foreach ($m in $mapRows) {
                $src = "" + $m.Source
                $dst = "" + $m.Destination
                if ($src -and $dst) { $mappingTable[$src.ToLowerInvariant()] = $dst }
            }
            Log "Loaded $($mappingTable.Count) address mapping(s) from -MappingCsv"
        } catch {
            Log "WARNING: Could not read -MappingCsv: $($_.Exception.Message.Split([Environment]::NewLine)[0])"
        }
    }
}

$tenantDomain = $null
if ($CustomerPrefix) {
    $cfgPath = Join-Path $env:APPDATA 'FlyMigration\config.json'
    if (Test-Path $cfgPath) {
        try {
            $cfg      = Get-Content $cfgPath -Raw | ConvertFrom-Json
            $customer = @($cfg.Customers) | Where-Object { $_.Prefix -and $_.Prefix.ToLower() -eq $CustomerPrefix.ToLower() } | Select-Object -First 1
            if ($customer -and $customer.AccountName -and $customer.AccountName -match '@(.+)$') {
                $tenantDomain = $Matches[1]
                Log "Customer '$CustomerPrefix' -> fallback tenant domain: $tenantDomain"
            } else {
                Log "WARNING: Customer '$CustomerPrefix' not found or has no usable AccountName — no fallback domain available"
            }
        } catch {
            Log "WARNING: Could not read config.json for -CustomerPrefix lookup: $($_.Exception.Message.Split([Environment]::NewLine)[0])"
        }
    }
}

function Resolve-NewAddress {
    param([string]$OldUpn, [string]$OldMail)
    foreach ($candidate in @($OldUpn, $OldMail) | Where-Object { $_ }) {
        $key = $candidate.ToLowerInvariant()
        if ($mappingTable.ContainsKey($key)) { return $mappingTable[$key] }
    }
    if ($tenantDomain) {
        $source = @($OldUpn, $OldMail) | Where-Object { $_ -match '^([^@]+)@' } | Select-Object -First 1
        if ($source -match '^([^@]+)@') { return "$($Matches[1])@$tenantDomain" }
    }
    return $null
}

# ── Graph modules + connection ──────────────────────────────────────────────────────
. (Join-Path $PSScriptRoot 'Ensure-GraphModules.ps1') -GraphModules @(
    'Microsoft.Graph.Users', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Teams', 'Microsoft.Graph.Identity.SignIns'
)

$scopes = @(
    'Team.ReadBasic.All', 'TeamMember.ReadWrite.All',
    'Channel.ReadBasic.All', 'ChannelMember.ReadWrite.All',
    'User.Read.All', 'Group.Read.All', 'User.Invite.All'
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

# ── Guest resolution (invite-or-reuse), cached per run ─────────────────────────────
$guestCache = @{}   # new address (lower) -> Entra object id, or $null if unresolvable

function Resolve-Guest {
    param([string]$NewAddress)

    $key = $NewAddress.ToLowerInvariant()
    if ($guestCache.ContainsKey($key)) { return $guestCache[$key] }

    # Already present in this tenant (a prior run, or invited some other way)?
    try {
        $existing = Get-MgUser -Filter "mail eq '$NewAddress'" -Property Id,Mail -ErrorAction Stop | Select-Object -First 1
        if (-not $existing) {
            $existing = Get-MgUser -Filter "otherMails/any(x:x eq '$NewAddress')" -Property Id,Mail -ErrorAction Stop | Select-Object -First 1
        }
        if ($existing) {
            Log "  Guest already exists: $NewAddress"
            $guestCache[$key] = $existing.Id
            return $existing.Id
        }
    } catch { }

    if ($WhatIf) {
        Log "  WhatIf: would invite guest $NewAddress"
        $guestCache[$key] = $null
        return $null
    }

    try {
        $inv = New-MgInvitation -InvitedUserEmailAddress $NewAddress `
            -InviteRedirectUrl $InviteRedirectUrl `
            -SendInvitationMessage:([bool]$SendInvitationEmail) `
            -ErrorAction Stop
        Log "  Invited guest: $NewAddress"
        $guestCache[$key] = $inv.InvitedUser.Id
        return $inv.InvitedUser.Id
    } catch {
        Log "  FAILED to invite ${NewAddress}: $($_.Exception.Message.Split([Environment]::NewLine)[0])"
        $guestCache[$key] = $null
        return $null
    }
}

# ── Add-member helper — treats "already a member" as a no-op success ──────────────
function Add-TeamOrChannelMember {
    param(
        [string]$TeamId,
        [string]$ChannelId,   # blank for team-level
        [string]$GuestUserId,
        [bool]$AsOwner
    )

    $body = @{
        '@odata.type'      = '#microsoft.graph.aadUserConversationMember'
        'roles'            = if ($AsOwner) { @('owner') } else { @() }
        'user@odata.bind'  = "https://graph.microsoft.com/v1.0/users('$GuestUserId')"
    }

    try {
        if ($ChannelId) {
            New-MgTeamChannelMember -TeamId $TeamId -ChannelId $ChannelId -BodyParameter $body -ErrorAction Stop | Out-Null
        } else {
            New-MgTeamMember -TeamId $TeamId -BodyParameter $body -ErrorAction Stop | Out-Null
        }
        return @{ Success = $true; Message = 'Added' }
    } catch {
        $msg = $_.Exception.Message.Split([Environment]::NewLine)[0]
        if ($msg -match 'already exist|Conflict|duplicate') {
            return @{ Success = $true; Message = 'Already a member' }
        }
        return @{ Success = $false; Message = $msg }
    }
}

# ── Process ─────────────────────────────────────────────────────────────────────────
$results = [System.Collections.Generic.List[object]]::new()
$ok = 0; $fail = 0; $skip = 0

foreach ($row in ($rows | Sort-Object TeamDisplayName, Scope, ChannelName, UserDisplayName)) {
    $label = "$($row.TeamDisplayName)$(if ($row.ChannelName) { " / $($row.ChannelName)" })  [$($row.UserDisplayName)]"

    $newAddress = Resolve-NewAddress -OldUpn $row.UserPrincipalName -OldMail $row.UserMail
    if (-not $newAddress) {
        Log "  $label — SKIPPED: no new address resolvable (not in mapping, no -CustomerPrefix fallback)"
        $results.Add([pscustomobject]@{
            TeamDisplayName = $row.TeamDisplayName; ChannelName = $row.ChannelName
            UserDisplayName = $row.UserDisplayName; OldAddress = ($row.UserPrincipalName, $row.UserMail | Where-Object { $_ } | Select-Object -First 1)
            NewAddress = ''; Role = $row.Role; Result = 'Skipped'; Message = 'No new address resolvable'
        })
        $skip++
        continue
    }

    $asOwner = ($row.Role -eq 'Owner')
    if ($asOwner) {
        Log "  $label — Owner role can't be preserved for a guest; adding as Member instead"
    }

    if ($WhatIf) {
        Log "  $label — WhatIf: would add $newAddress as $(if ($asOwner) { 'Member (was Owner)' } else { $row.Role })"
        $results.Add([pscustomobject]@{
            TeamDisplayName = $row.TeamDisplayName; ChannelName = $row.ChannelName
            UserDisplayName = $row.UserDisplayName; OldAddress = ($row.UserPrincipalName, $row.UserMail | Where-Object { $_ } | Select-Object -First 1)
            NewAddress = $newAddress; Role = $row.Role; Result = 'WhatIf'; Message = 'Preview only'
        })
        $ok++
        continue
    }

    $guestId = Resolve-Guest -NewAddress $newAddress
    if (-not $guestId) {
        Log "  $label — FAILED: could not resolve or invite guest $newAddress"
        $results.Add([pscustomobject]@{
            TeamDisplayName = $row.TeamDisplayName; ChannelName = $row.ChannelName
            UserDisplayName = $row.UserDisplayName; OldAddress = ($row.UserPrincipalName, $row.UserMail | Where-Object { $_ } | Select-Object -First 1)
            NewAddress = $newAddress; Role = $row.Role; Result = 'Failed'; Message = 'Guest invite failed'
        })
        $fail++
        continue
    }

    $add = Add-TeamOrChannelMember -TeamId $row.TeamId -ChannelId $row.ChannelId -GuestUserId $guestId -AsOwner:$asOwner
    if ($add.Success) {
        Log "  $label — $($add.Message): $newAddress"
        $ok++
    } else {
        Log "  $label — FAILED: $($add.Message)"
        $fail++
    }
    $results.Add([pscustomobject]@{
        TeamDisplayName = $row.TeamDisplayName; ChannelName = $row.ChannelName
        UserDisplayName = $row.UserDisplayName; OldAddress = ($row.UserPrincipalName, $row.UserMail | Where-Object { $_ } | Select-Object -First 1)
        NewAddress = $newAddress; Role = $row.Role; Result = $(if ($add.Success) { 'OK' } else { 'Failed' }); Message = $add.Message
    })
}

$results | Export-Csv -Path $ResultsCsv -NoTypeInformation -Encoding UTF8

Log ''
Log "=== Complete: ok $ok  |  failed $fail  |  skipped $skip ==="
Log "Results CSV : $ResultsCsv"
Log "Log         : $logFile"

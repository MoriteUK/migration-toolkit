#Requires -Version 7.0
<#
.SYNOPSIS
    Add-DistributionGroupMembers.ps1 — Syncs Distribution Group membership in the destination
    tenant from Discovery's captured membership, without creating any groups.

.DESCRIPTION
    Reads 03b_DistributionGroup_Members.csv (one row per group+member, written by every
    Run-Assessment.ps1 assessment, or the older search-domain.ps1) and, for each group that
    already exists in the DESTINATION tenant, adds every member captured in Discovery —
    resolving each one's NEW (destination-tenant) address the same way New-DistributionGroups.ps1
    and Restore-ProxyAddresses.ps1 do, via -MappingCsv (the Fly exchange mapping file,
    authoritative) or -CustomerPrefix (derives <local-part>@<tenant domain> as a fallback).

    Falls back to 03_DistributionGroups.csv's pipe-delimited Members column when
    03b_DistributionGroup_Members.csv isn't present (an older Discovery folder) — same source
    data, just not yet flattened to one-row-per-member.

    This is the standalone version of the membership step in New-DistributionGroups.ps1 — use
    it when the groups already exist (created earlier, or by other means) and you just need to
    (re-)sync membership, e.g. after new members were added to the source-tenant group after the
    groups were first created. A group not found in the destination tenant is logged and
    skipped — it never creates one. Existing members are left alone — idempotent, safe to re-run.

    Connects to Exchange Online — sign in with the DESTINATION tenant admin account.

.PARAMETER DiscoveryFolder
    Path to the domain's discovery output folder (or its Discovery subfolder directly).

.PARAMETER CustomerPrefix
    Settings > Customers Prefix (e.g. "AEP"). Looked up in
    %APPDATA%\FlyMigration\config.json to derive the destination tenant's onmicrosoft.com
    domain from that customer's AccountName — used as the fallback address for any member not
    found in -MappingCsv.

.PARAMETER MappingCsv
    Optional. The Fly exchange mapping file (Source/Destination columns) used to resolve each
    member's new destination-tenant address. Accepts .xlsx (via ImportExcel) or .csv. Falls
    back to -CustomerPrefix's tenant domain for anyone not found in it.

.PARAMETER WhatIf
    Preview which members would be added to which groups, without making any changes.

.EXAMPLE
    .\Add-DistributionGroupMembers.ps1 -DiscoveryFolder "C:\...\aep-italia.it" -CustomerPrefix AEP -WhatIf

.EXAMPLE
    .\Add-DistributionGroupMembers.ps1 -DiscoveryFolder "C:\...\aep-italia.it" -CustomerPrefix AEP `
        -MappingCsv ".\AEP mapping exchange.xlsx"

.NOTES
    Not yet exercised against a live tenant - run with -WhatIf first and read the results CSV
    closely before trusting a live run.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$DiscoveryFolder,

    [Parameter(Mandatory = $true)]
    [string]$CustomerPrefix,

    [string]$MappingCsv = '',
    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'

# ── Logging ──────────────────────────────────────────────────────────────────────
$_logDir = 'C:\Users\andyw\OneDrive - Andy White\Contracts\Jolera\Migrations\Logs'
if (-not (Test-Path $_logDir)) { New-Item -ItemType Directory -Path $_logDir -Force | Out-Null }
$logFile = Join-Path $_logDir "add-dg-members-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
function Log { param([string]$m) $ts = Get-Date -Format 'HH:mm:ss'; "$ts $m" | Tee-Object -FilePath $logFile -Append | Write-Host }

$_reportsDir = 'C:\Users\andyw\OneDrive - Andy White\Contracts\Jolera\Migrations\Logs'
if (-not (Test-Path $_reportsDir)) { New-Item -ItemType Directory -Path $_reportsDir -Force | Out-Null }
$resultsCsv = Join-Path $_reportsDir "AddDistributionGroupMembers-$CustomerPrefix-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"

Log "=== Add Distribution Group Members$(if ($WhatIf) { ' [WhatIf]' }) ==="
Log "Customer      : $CustomerPrefix"
if ($MappingCsv) { Log "Mapping file  : $MappingCsv" }
Log "Results CSV   : $resultsCsv"

# ── Resolve the Discovery folder + source CSV ──────────────────────────────────────
$discFolder = $DiscoveryFolder.Trim().Trim('"')
$candidate  = Join-Path $discFolder 'Discovery'
if ((Split-Path $discFolder -Leaf) -ne 'Discovery' -and (Test-Path $candidate)) {
    $discFolder = $candidate
}
$csvPath = Join-Path $discFolder '03b_DistributionGroup_Members.csv'
$legacyFallback = $false
if (-not (Test-Path $csvPath)) {
    Log "03b_DistributionGroup_Members.csv not found in: $discFolder — falling back to 03_DistributionGroups.csv's Members column (older Discovery folder)"
    $csvPath = Join-Path $discFolder '03_DistributionGroups.csv'
    $legacyFallback = $true
    if (-not (Test-Path $csvPath)) {
        Log "ERROR: Neither 03b_DistributionGroup_Members.csv nor 03_DistributionGroups.csv found in: $discFolder"
        exit 1
    }
}
$memberRows = @(Import-Csv -Path $csvPath -Encoding UTF8)
Log "Loaded $($memberRows.Count) row(s) from $(Split-Path $csvPath -Leaf)"
if ($memberRows.Count -eq 0) {
    Log "Nothing to sync."
    exit 0
}

# ── Flatten to one entry per group: Alias, DisplayName, Members[] ─────────────────
# Handles both the current engine's 03b schema (GroupAlias/GroupPrimarySmtpAddress/
# MemberAddress) and the legacy search-domain.ps1 schema (GroupEmail/MemberEmail, plus a
# '(no members)' placeholder row per empty group — skipped here via the blank-address check).
$groups = [System.Collections.Generic.List[object]]::new()
if ($legacyFallback) {
    if (-not $memberRows[0].PSObject.Properties['Members']) {
        Log "ERROR: 03_DistributionGroups.csv has no 'Members' column either - re-run Discovery so it captures DL membership."
        exit 1
    }
    foreach ($row in ($memberRows | Where-Object { $_.Alias })) {
        $members = @($row.Members -split '\|' | Where-Object { $_ })
        $groups.Add([pscustomobject]@{ Alias = $row.Alias; DisplayName = if ($row.DisplayName) { $row.DisplayName } else { $row.Alias }; Members = $members })
    }
} else {
    $hasGroupAlias = $null -ne $memberRows[0].PSObject.Properties['GroupAlias']
    $byGroup = [ordered]@{}
    foreach ($r in $memberRows) {
        $memberAddr = if ($r.PSObject.Properties['MemberAddress']) { "$($r.MemberAddress)" } else { "$($r.MemberEmail)" }
        if (-not $memberAddr) { continue }

        $alias = if ($hasGroupAlias -and $r.GroupAlias) {
            "$($r.GroupAlias)"
        } else {
            $groupAddr = if ($r.PSObject.Properties['GroupPrimarySmtpAddress']) { "$($r.GroupPrimarySmtpAddress)" } else { "$($r.GroupEmail)" }
            if ("$groupAddr" -match '^([^@]+)@') { $Matches[1] } else { $null }
        }
        if (-not $alias) { continue }

        if (-not $byGroup.Contains($alias)) {
            $byGroup[$alias] = [pscustomobject]@{ Alias = $alias; DisplayName = "$($r.GroupDisplayName)"; Members = [System.Collections.Generic.List[string]]::new() }
        }
        $byGroup[$alias].Members.Add($memberAddr)
    }
    foreach ($key in $byGroup.Keys) { $groups.Add($byGroup[$key]) }
}
Log "Grouped into $($groups.Count) distribution group(s) with members captured"
if ($groups.Count -eq 0) {
    Log "Nothing to sync."
    exit 0
}

# ── Resolve the destination tenant's onmicrosoft.com domain ───────────────────────
$cfgPath = Join-Path $env:APPDATA 'FlyMigration\config.json'
$tenantDomain = $null
if (Test-Path $cfgPath) {
    try {
        $cfg      = Get-Content $cfgPath -Raw | ConvertFrom-Json
        $customer = @($cfg.Customers) | Where-Object { $_.Prefix -and $_.Prefix.ToLower() -eq $CustomerPrefix.ToLower() } | Select-Object -First 1
        if ($customer -and $customer.AccountName -and $customer.AccountName -match '@(.+)$') {
            $tenantDomain = $Matches[1]
        }
    } catch { }
}
if (-not $tenantDomain) {
    Log "ERROR: Could not resolve a tenant domain for customer '$CustomerPrefix' from config.json. Check Settings > Customers."
    exit 1
}
Log "Destination tenant domain (fallback for unmapped members): $tenantDomain"

# ── Load the member address mapping (old -> new), if given ────────────────────────
$mappingTable = @{}
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

function Resolve-NewAddress {
    param([string]$OldAddress)
    if (-not $OldAddress) { return $null }
    $key = $OldAddress.ToLowerInvariant()
    if ($mappingTable.ContainsKey($key)) { return $mappingTable[$key] }
    if ($OldAddress -match '^([^@]+)@') { return "$($Matches[1])@$tenantDomain" }
    return $null
}

# ── Connect to Exchange Online (destination tenant) ────────────────────────────────
$mod = Get-Module -ListAvailable -Name 'ExchangeOnlineManagement' -ErrorAction SilentlyContinue
if (-not $mod) {
    Log 'ERROR: ExchangeOnlineManagement module is not installed.'
    exit 1
}
Import-Module 'ExchangeOnlineManagement' -ErrorAction Stop

# Disconnect any existing session first - a session left over from an earlier script run
# (e.g. Discovery, moments earlier, against the SOURCE tenant) is otherwise silently reused
# instead of prompting fresh, and -Organization below would be the only thing standing
# between this run and updating everything against the wrong tenant.
try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch {}

Log "Connecting to Exchange Online ($tenantDomain) — sign in with the DESTINATION tenant admin account when the browser opens..."
try {
    Connect-ExchangeOnline -Organization $tenantDomain -ShowBanner:$false -DisableWAM -ErrorAction Stop
    Log 'Connected to Exchange Online.'
} catch {
    Log "ERROR: Failed to connect to Exchange Online: $($_.Exception.Message.Split([Environment]::NewLine)[0])"
    exit 1
}

$results = [System.Collections.Generic.List[object]]::new()
$groupsNotFound = 0; $groupsSkippedNoMembers = 0
$membersAdded = 0; $membersSkipped = 0; $membersFailed = 0

foreach ($grp in $groups) {
    $alias       = $grp.Alias
    $displayName = if ($grp.DisplayName) { $grp.DisplayName } else { $alias }
    $members     = @($grp.Members)

    Log "--- $displayName [$alias] ---"

    if ($members.Count -eq 0) {
        Log "  No members captured for this group in Discovery."
        $groupsSkippedNoMembers++
        continue
    }

    $dg = $null
    try { $dg = Get-DistributionGroup -Identity $alias -ErrorAction Stop } catch { }
    if (-not $dg) {
        try { $dg = Get-Recipient -Filter "EmailAddresses -like '*$alias@*'" -ErrorAction Stop | Select-Object -First 1 } catch { }
    }

    if (-not $dg) {
        Log "  NOT FOUND in destination tenant — skipped (run Create Target DLs first)"
        $groupsNotFound++
        $results.Add([pscustomobject]@{ DisplayName = $displayName; Alias = $alias; Member = ''; Result = 'GroupNotFound'; Message = '' })
        continue
    }

    $currentMemberAddrs = @()
    if (-not $WhatIf) {
        try {
            $currentMemberAddrs = @(Get-DistributionGroupMember -Identity $alias -ResultSize Unlimited -ErrorAction Stop |
                ForEach-Object { "$($_.PrimarySmtpAddress)".ToLowerInvariant() } | Where-Object { $_ })
        } catch { }
    }

    foreach ($oldAddr in $members) {
        $newAddr = Resolve-NewAddress -OldAddress $oldAddr
        if (-not $newAddr) {
            Log "    member $oldAddr — SKIPPED: no new address resolvable"
            $membersSkipped++
            $results.Add([pscustomobject]@{ DisplayName = $displayName; Alias = $alias; Member = $oldAddr; Result = 'Unresolvable'; Message = '' })
            continue
        }

        if ($currentMemberAddrs -contains $newAddr.ToLowerInvariant()) {
            $membersSkipped++
            $results.Add([pscustomobject]@{ DisplayName = $displayName; Alias = $alias; Member = $newAddr; Result = 'AlreadyMember'; Message = '' })
            continue
        }

        if ($WhatIf) {
            Log "    member $oldAddr — WhatIf: would add $newAddr"
            $membersAdded++
            $results.Add([pscustomobject]@{ DisplayName = $displayName; Alias = $alias; Member = $newAddr; Result = 'WhatIf-WouldAdd'; Message = "was $oldAddr" })
            continue
        }

        try {
            Add-DistributionGroupMember -Identity $alias -Member $newAddr -ErrorAction Stop
            Log "    member added: $newAddr"
            $membersAdded++
            $results.Add([pscustomobject]@{ DisplayName = $displayName; Alias = $alias; Member = $newAddr; Result = 'Added'; Message = "was $oldAddr" })
        } catch {
            $msg = $_.Exception.Message.Split([Environment]::NewLine)[0]
            if ($msg -match 'already a member|already exist') {
                $membersSkipped++
                $results.Add([pscustomobject]@{ DisplayName = $displayName; Alias = $alias; Member = $newAddr; Result = 'AlreadyMember'; Message = '' })
            } else {
                Log "    member $newAddr — FAILED: $msg"
                $membersFailed++
                $results.Add([pscustomobject]@{ DisplayName = $displayName; Alias = $alias; Member = $newAddr; Result = 'Failed'; Message = $msg })
            }
        }
    }
}

try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}

$results | Export-Csv -Path $resultsCsv -NoTypeInformation -Encoding UTF8

Log ''
Log "=== Complete: members added $membersAdded  |  skipped (already present / unresolvable) $membersSkipped  |  failed $membersFailed ==="
Log "=== Groups: not found in destination $groupsNotFound  |  no members captured $groupsSkippedNoMembers ==="
Log "Results CSV : $resultsCsv"
Log "Log         : $logFile"

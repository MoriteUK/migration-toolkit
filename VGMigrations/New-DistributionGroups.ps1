#Requires -Version 7.0
<#
.SYNOPSIS
    New-DistributionGroups.ps1 — Creates any Distribution Groups / mail-enabled security groups
    from Discovery that don't already exist in the destination tenant, and adds their members.

.DESCRIPTION
    Reads 03_DistributionGroups.csv (written by every Run-Assessment.ps1 assessment, or the
    older search-domain.ps1) and, for each row:
      1. Checks whether a group with the same Alias already exists in the DESTINATION tenant.
      2. If not, creates it — as a Distribution Group or Mail-Enabled Security Group, matching
         the source GroupType — using the DESTINATION TENANT's onmicrosoft.com address as its
         primary SMTP address (e.g. sales@mbuaepts.onmicrosoft.com), since the real vanity
         domain isn't verified on the destination tenant yet at this point in a migration. Run
         Update-DistributionGroupDomain.ps1 afterwards, once the domain has cut over, to flip
         the primary address back to the real one.
      3. Adds every member captured in Discovery's Members column, resolving each one's NEW
         (destination-tenant) address the same way Restore-ProxyAddresses.ps1 and
         Restore-DomainTeamMemberships.ps1 do — via -MappingCsv (the Fly exchange mapping file,
         authoritative) or -CustomerPrefix (derives <local-part>@<tenant domain> as a fallback).
      4. Membership is synced on every run, not just at creation - a group that already exists
         still gets any missing members added. Existing members and existing groups are left
         alone (idempotent - safe to re-run).

    Connects to Exchange Online — sign in with the DESTINATION tenant admin account.

.PARAMETER DiscoveryFolder
    Path to the domain's discovery output folder (or its Discovery subfolder directly).

.PARAMETER CustomerPrefix
    Settings > Customers Prefix (e.g. "AEP"). Looked up in
    %APPDATA%\FlyMigration\config.json to derive the destination tenant's onmicrosoft.com
    domain from that customer's AccountName (e.g. itvolaris@mbuaepts.onmicrosoft.com ->
    mbuaepts.onmicrosoft.com) — required to build the temporary primary address new groups are
    created with.

.PARAMETER MappingCsv
    Optional. The Fly exchange mapping file (Source/Destination columns) used to resolve each
    member's new destination-tenant address. Accepts .xlsx (via ImportExcel) or .csv. Falls
    back to -CustomerPrefix's tenant domain for anyone not found in it.

.PARAMETER WhatIf
    Preview which groups would be created and which members would be added, without making
    any changes.

.EXAMPLE
    .\New-DistributionGroups.ps1 -DiscoveryFolder "C:\...\aep-italia.it" -CustomerPrefix AEP -WhatIf

.EXAMPLE
    .\New-DistributionGroups.ps1 -DiscoveryFolder "C:\...\aep-italia.it" -CustomerPrefix AEP `
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
$logFile = Join-Path $_logDir "new-distribution-groups-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
function Log { param([string]$m) $ts = Get-Date -Format 'HH:mm:ss'; "$ts $m" | Tee-Object -FilePath $logFile -Append | Write-Host }

$_reportsDir = 'C:\Users\andyw\OneDrive - Andy White\Contracts\Jolera\Migrations\Reports'
if (-not (Test-Path $_reportsDir)) { New-Item -ItemType Directory -Path $_reportsDir -Force | Out-Null }
$resultsCsv = Join-Path $_reportsDir "NewDistributionGroups-$CustomerPrefix-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"

Log "=== New Distribution Groups$(if ($WhatIf) { ' [WhatIf]' }) ==="
Log "Customer      : $CustomerPrefix"
if ($MappingCsv) { Log "Mapping file  : $MappingCsv" }
Log "Results CSV   : $resultsCsv"

# ── Resolve the Discovery folder + source CSV ──────────────────────────────────────
$discFolder = $DiscoveryFolder.Trim().Trim('"')
$candidate  = Join-Path $discFolder 'Discovery'
if ((Split-Path $discFolder -Leaf) -ne 'Discovery' -and (Test-Path $candidate)) {
    $discFolder = $candidate
}
$csvPath = Join-Path $discFolder '03_DistributionGroups.csv'
if (-not (Test-Path $csvPath)) {
    Log "ERROR: 03_DistributionGroups.csv not found in: $discFolder"
    exit 1
}
$rows = @(Import-Csv -Path $csvPath -Encoding UTF8 | Where-Object { $_.Alias })
Log "Loaded $($rows.Count) group(s) from 03_DistributionGroups.csv"
if ($rows.Count -eq 0) {
    Log "Nothing to create."
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
Log "Destination tenant domain: $tenantDomain"

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
# between this run and creating everything against the wrong tenant.
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
$created = 0; $existing = 0; $failed = 0
$membersAdded = 0; $membersSkipped = 0; $membersFailed = 0
$rowsWithNoMembersColumn = 0

foreach ($row in $rows) {
    $alias       = $row.Alias
    $displayName = if ($row.DisplayName) { $row.DisplayName } else { $alias }
    $isSecurity  = $row.GroupType -eq 'MailUniversalSecurityGroup'
    $tenantAddress = "$alias@$tenantDomain"

    Log "--- $displayName [$alias] ---"

    # ── Find or create ──────────────────────────────────────────────────────────
    $dg = $null
    try { $dg = Get-DistributionGroup -Identity $alias -ErrorAction Stop } catch { }
    if (-not $dg) {
        try {
            $dg = Get-Recipient -Filter "EmailAddresses -like '*$alias@*'" -ErrorAction Stop | Select-Object -First 1
        } catch { }
    }

    if ($dg) {
        Log "  Already exists: $($dg.PrimarySmtpAddress)"
        $existing++
        $results.Add([pscustomobject]@{ DisplayName = $displayName; Alias = $alias; Result = 'AlreadyExists'; Address = "$($dg.PrimarySmtpAddress)"; Message = '' })
    }
    elseif ($WhatIf) {
        Log "  WhatIf: would create as $(if ($isSecurity) { 'Mail-Enabled Security Group' } else { 'Distribution Group' }) with primary $tenantAddress"
        $results.Add([pscustomobject]@{ DisplayName = $displayName; Alias = $alias; Result = 'WhatIf-WouldCreate'; Address = $tenantAddress; Message = '' })
    }
    else {
        try {
            $newParams = @{
                Name               = $displayName
                Alias              = $alias
                PrimarySmtpAddress = $tenantAddress
                ErrorAction        = 'Stop'
            }
            if ($isSecurity) { $newParams.Type = 'Security' }
            $dg = New-DistributionGroup @newParams
            Log "  CREATED: $tenantAddress"
            $created++

            if ($row.HiddenFromAddressListsEnabled -eq 'True') {
                try { Set-DistributionGroup -Identity $alias -HiddenFromAddressListsEnabled $true -ErrorAction Stop } catch { }
            }
            $results.Add([pscustomobject]@{ DisplayName = $displayName; Alias = $alias; Result = 'Created'; Address = $tenantAddress; Message = '' })
        } catch {
            $msg = $_.Exception.Message.Split([Environment]::NewLine)[0]
            Log "  FAILED to create: $msg"
            $failed++
            $results.Add([pscustomobject]@{ DisplayName = $displayName; Alias = $alias; Result = 'Failed'; Address = $tenantAddress; Message = $msg })
            continue
        }
    }

    # ── Members ──────────────────────────────────────────────────────────────────
    if (-not $dg -and -not $WhatIf) { continue }
    if (-not $row.PSObject.Properties['Members']) {
        Log "    No 'Members' column in this CSV - it was generated before Discovery captured DL membership. Re-run Discovery, then re-run this script (it's safe to re-run - existing groups/members are left alone)."
        $rowsWithNoMembersColumn++
        continue
    }
    $members = @($row.Members -split '\|' | Where-Object { $_ })
    if ($members.Count -eq 0) {
        Log "    No members captured for this group in Discovery."
        continue
    }

    $currentMemberAddrs = @()
    if ($dg -and -not $WhatIf) {
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
            continue
        }

        if ($currentMemberAddrs -contains $newAddr.ToLowerInvariant()) {
            $membersSkipped++
            continue
        }

        if ($WhatIf) {
            Log "    member $oldAddr — WhatIf: would add $newAddr"
            $membersAdded++
            continue
        }

        try {
            Add-DistributionGroupMember -Identity $alias -Member $newAddr -ErrorAction Stop
            Log "    member added: $newAddr"
            $membersAdded++
        } catch {
            $msg = $_.Exception.Message.Split([Environment]::NewLine)[0]
            if ($msg -match 'already a member|already exist') {
                $membersSkipped++
            } else {
                Log "    member $newAddr — FAILED: $msg"
                $membersFailed++
            }
        }
    }
}

try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}

$results | Export-Csv -Path $resultsCsv -NoTypeInformation -Encoding UTF8

Log ''
Log "=== Complete: groups created $created  |  already existed $existing  |  failed $failed ==="
Log "=== Members: added $membersAdded  |  skipped (already present / unresolvable) $membersSkipped  |  failed $membersFailed ==="
if ($rowsWithNoMembersColumn -gt 0) {
    Log "=== WARNING: $rowsWithNoMembersColumn group(s) had no 'Members' column at all - this Discovery folder predates member capture. Re-run Discovery, then re-run this script. ==="
}
Log "Results CSV : $resultsCsv"
Log "Log         : $logFile"

#Requires -Version 7.0
<#
.SYNOPSIS
    Update-DistributionGroupDomain.ps1 — Flips each Distribution Group's primary SMTP address
    from its temporary destination-tenant address back to the real vanity domain, once that
    domain has been added and verified on the destination tenant (post-cutover).

.DESCRIPTION
    New-DistributionGroups.ps1 creates groups in the destination tenant using a temporary
    address on the tenant's own onmicrosoft.com domain (e.g. sales@mbuaepts.onmicrosoft.com),
    because the real vanity domain isn't verified there yet at that point in the migration.
    Once the domain has cut over — added to and verified on the destination tenant — this
    script updates each group's primary address to the real one, taken straight from
    03_DistributionGroups.csv's PrimarySmtpAddress column (the address it had in the source
    tenant before migration). The temporary tenant address is kept as a secondary alias, not
    removed — nothing here can break anything still pointed at it.

    Groups are matched by Alias, same as New-DistributionGroups.ps1 uses to find/create them.
    A group not found is logged and skipped (it may not have been created via that script, or
    may already have been renamed by a previous run of this one).

    Connects to Exchange Online — sign in with the DESTINATION tenant admin account.

.PARAMETER DiscoveryFolder
    Path to the domain's discovery output folder (or its Discovery subfolder directly) — the
    same one New-DistributionGroups.ps1 was pointed at.

.PARAMETER CustomerPrefix
    Settings > Customers Prefix (e.g. "AEP"). Looked up in
    %APPDATA%\FlyMigration\config.json to derive the destination tenant's onmicrosoft.com
    domain from that customer's AccountName — needed to identify which address is the
    temporary one being replaced.

.PARAMETER WhatIf
    Preview which groups would be updated and to what address, without making any changes.

.EXAMPLE
    .\Update-DistributionGroupDomain.ps1 -DiscoveryFolder "C:\...\aep-italia.it" -CustomerPrefix AEP -WhatIf

.NOTES
    Not yet exercised against a live tenant - run with -WhatIf first and read the results CSV
    closely before trusting a live run.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$DiscoveryFolder,

    [Parameter(Mandatory = $true)]
    [string]$CustomerPrefix,

    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'

# ── Logging ──────────────────────────────────────────────────────────────────────
$_logDir = 'C:\Users\andyw\OneDrive - Andy White\Contracts\Jolera\Migrations\Logs'
if (-not (Test-Path $_logDir)) { New-Item -ItemType Directory -Path $_logDir -Force | Out-Null }
$logFile = Join-Path $_logDir "update-dg-domain-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
function Log { param([string]$m) $ts = Get-Date -Format 'HH:mm:ss'; "$ts $m" | Tee-Object -FilePath $logFile -Append | Write-Host }

$_reportsDir = 'C:\Users\andyw\OneDrive - Andy White\Contracts\Jolera\Migrations\Reports'
if (-not (Test-Path $_reportsDir)) { New-Item -ItemType Directory -Path $_reportsDir -Force | Out-Null }
$resultsCsv = Join-Path $_reportsDir "UpdateDistributionGroupDomain-$CustomerPrefix-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"

Log "=== Update Distribution Group Domain$(if ($WhatIf) { ' [WhatIf]' }) ==="
Log "Customer      : $CustomerPrefix"
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
$rows = @(Import-Csv -Path $csvPath -Encoding UTF8 | Where-Object { $_.Alias -and $_.PrimarySmtpAddress })
Log "Loaded $($rows.Count) group(s) from 03_DistributionGroups.csv"
if ($rows.Count -eq 0) {
    Log "Nothing to update."
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
Log "Destination tenant domain (temporary address being replaced): $tenantDomain"

# ── Connect to Exchange Online (destination tenant) ────────────────────────────────
$mod = Get-Module -ListAvailable -Name 'ExchangeOnlineManagement' -ErrorAction SilentlyContinue
if (-not $mod) {
    Log 'ERROR: ExchangeOnlineManagement module is not installed.'
    exit 1
}
Import-Module 'ExchangeOnlineManagement' -ErrorAction Stop

Log 'Connecting to Exchange Online — sign in with the DESTINATION tenant admin account when the browser opens...'
try {
    Connect-ExchangeOnline -ShowBanner:$false -DisableWAM -ErrorAction Stop
    Log 'Connected to Exchange Online.'
} catch {
    Log "ERROR: Failed to connect to Exchange Online: $($_.Exception.Message.Split([Environment]::NewLine)[0])"
    exit 1
}

$results = [System.Collections.Generic.List[object]]::new()
$ok = 0; $alreadyDone = 0; $notFound = 0; $failed = 0

foreach ($row in $rows) {
    $alias         = $row.Alias
    $displayName   = if ($row.DisplayName) { $row.DisplayName } else { $alias }
    $targetAddress = $row.PrimarySmtpAddress
    $tenantAddress = "$alias@$tenantDomain"

    Log "--- $displayName [$alias] ---"

    $dg = $null
    try { $dg = Get-DistributionGroup -Identity $alias -ErrorAction Stop } catch { }
    if (-not $dg) {
        try { $dg = Get-Recipient -Filter "EmailAddresses -like '*$alias@*'" -ErrorAction Stop | Select-Object -First 1 } catch { }
    }

    if (-not $dg) {
        Log "  NOT FOUND in destination tenant — skipped"
        $notFound++
        $results.Add([pscustomobject]@{ DisplayName = $displayName; Alias = $alias; Result = 'NotFound'; TargetAddress = $targetAddress; Message = '' })
        continue
    }

    $currentPrimary = "$($dg.PrimarySmtpAddress)"
    if ($currentPrimary.ToLowerInvariant() -eq $targetAddress.ToLowerInvariant()) {
        Log "  Already on target address: $currentPrimary"
        $alreadyDone++
        $results.Add([pscustomobject]@{ DisplayName = $displayName; Alias = $alias; Result = 'AlreadyOnTarget'; TargetAddress = $targetAddress; Message = '' })
        continue
    }

    if ($WhatIf) {
        Log "  WhatIf: would set primary $currentPrimary -> $targetAddress (kept as secondary)"
        $ok++
        $results.Add([pscustomobject]@{ DisplayName = $displayName; Alias = $alias; Result = 'WhatIf-WouldUpdate'; TargetAddress = $targetAddress; Message = "was $currentPrimary" })
        continue
    }

    try {
        # Build the address list explicitly rather than relying on -PrimarySmtpAddress's
        # implicit behaviour - matches this codebase's established pattern
        # (Restore-ProxyAddresses.ps1) of never silently dropping an existing address.
        $current = @($dg.EmailAddresses | ForEach-Object { "$_" })
        $updated = [System.Collections.Generic.List[string]]::new()
        $addedTarget = $false
        foreach ($addr in $current) {
            if ($addr -match '^smtp:(.+)$' -and $Matches[1].ToLowerInvariant() -eq $currentPrimary.ToLowerInvariant()) {
                # skip - the old primary is re-added below as a lowercase secondary
                continue
            }
            if ($addr -match '^SMTP:(.+)$') {
                # demote the current primary to secondary
                $updated.Add("smtp:$($Matches[1])")
                continue
            }
            $updated.Add($addr)
        }
        $updated.Add("SMTP:$targetAddress")
        if (-not ($updated | Where-Object { $_ -ieq "smtp:$tenantAddress" -or $_ -ieq "SMTP:$tenantAddress" })) {
            $updated.Add("smtp:$tenantAddress")
        }

        Set-DistributionGroup -Identity $alias -EmailAddresses $updated -ErrorAction Stop
        Log "  UPDATED: $currentPrimary -> $targetAddress (kept as secondary: $tenantAddress)"
        $ok++
        $results.Add([pscustomobject]@{ DisplayName = $displayName; Alias = $alias; Result = 'Updated'; TargetAddress = $targetAddress; Message = "was $currentPrimary" })
    } catch {
        $msg = $_.Exception.Message.Split([Environment]::NewLine)[0]
        Log "  FAILED: $msg"
        $failed++
        $results.Add([pscustomobject]@{ DisplayName = $displayName; Alias = $alias; Result = 'Failed'; TargetAddress = $targetAddress; Message = $msg })
    }
}

try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}

$results | Export-Csv -Path $resultsCsv -NoTypeInformation -Encoding UTF8

Log ''
Log "=== Complete: updated $ok  |  already on target $alreadyDone  |  not found $notFound  |  failed $failed ==="
Log "Results CSV : $resultsCsv"
Log "Log         : $logFile"

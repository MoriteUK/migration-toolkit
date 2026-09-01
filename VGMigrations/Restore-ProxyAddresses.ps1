#Requires -Version 7.0
<#
.SYNOPSIS
    Restore-ProxyAddresses.ps1 — Re-adds proxy addresses captured before migration to the
    matching accounts in the new (destination) tenant, as secondary aliases.

.DESCRIPTION
    Reads 13_ProxyAddresses.csv from a domain's Discovery folder (written by search-domain.ps1 —
    one row per address, captured from the OLD tenant before cutover) and adds any addresses
    still missing back onto the corresponding recipient in whichever tenant the signed-in
    account belongs to — normally the NEW/destination tenant, once accounts have moved there.

    Recipients are matched by PrimarySmtpAddress first, then UserPrincipalName, then — when
    -CustomerPrefix is supplied — by local-part @ the destination tenant's onmicrosoft.com
    domain (looked up from %APPDATA%\FlyMigration\config.json's Customers array, same one
    Settings > Customers populates), then finally by local-part across any domain as a last
    resort. This matters because the old vanity domain (e.g. aep-italia.it) usually isn't
    verified on the destination tenant yet when this script runs — accounts are still sitting
    on their default onmicrosoft.com address (e.g. user@mbuaepts.onmicrosoft.com) — so matching
    on PrimarySmtpAddress/UserPrincipalName alone (both still on the OLD domain in the CSV)
    finds nothing. Restored addresses are always added as SECONDARY (lowercase prefix) — even
    if a row's IsPrimary was true in the old tenant — so nothing here can override whatever
    primary address the destination mailbox already has. Addresses already present
    (case-insensitive) are skipped.

    X500 addresses are internal Exchange routing addresses tied to the OLD mailbox's GUID and
    are skipped by default — pass -IncludeX500 to restore them too if you specifically need
    legacy-DN redirects to keep working.

.PARAMETER DiscoveryFolder
    Path to the domain's discovery output folder (or its Discovery subfolder directly).

.PARAMETER Domain
    Optional: only restore rows whose ProxyAddress contains this domain. If omitted, every row
    in 13_ProxyAddresses.csv is considered (the file is already scoped to whatever domain was
    searched when discovery ran, so this is normally unnecessary).

.PARAMETER CustomerPrefix
    Optional: the customer's Prefix as set in Settings > Customers (e.g. "AEP"). Looked up in
    %APPDATA%\FlyMigration\config.json to derive the destination tenant's onmicrosoft.com
    domain from that customer's AccountName (e.g. itvolaris@mbuaepts.onmicrosoft.com ->
    mbuaepts.onmicrosoft.com), used as an extra match strategy before accounts have been
    renamed onto their final vanity domain in the destination tenant. Strongly recommended
    when running this script before cutover/rename.

.PARAMETER SkipSMTP
    Skip restoring SMTP alias addresses.

.PARAMETER SkipSIP
    Skip restoring SIP/EUM/IM addresses.

.PARAMETER IncludeX500
    Also restore X500 addresses (skipped by default — see DESCRIPTION).

.PARAMETER WhatIf
    Preview which addresses would be added without making changes.

.EXAMPLE
    .\Restore-ProxyAddresses.ps1 -DiscoveryFolder "C:\...\contoso.com" -WhatIf
    .\Restore-ProxyAddresses.ps1 -DiscoveryFolder "C:\...\contoso.com" -Domain contoso.com
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$DiscoveryFolder,

    [string]$Domain = '',
    [string]$CustomerPrefix = '',

    [switch]$SkipSMTP,
    [switch]$SkipSIP,
    [switch]$IncludeX500,
    [switch]$WhatIf
)

$script:RootDir = $PSScriptRoot

$_logDir = 'C:\Users\andyw\OneDrive - Andy White\Contracts\Jolera\Migrations\Logs'
if (-not (Test-Path $_logDir)) { New-Item -ItemType Directory -Path $_logDir -Force | Out-Null }
$logFile = Join-Path $_logDir "restore-proxy-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
function Log { param([string]$m) $ts = Get-Date -Format 'HH:mm:ss'; "$ts $m" | Tee-Object -FilePath $logFile -Append | Write-Host }

if ($SkipSMTP -and $SkipSIP -and -not $IncludeX500) {
    Log 'ERROR: Nothing to restore — SMTP and SIP/EUM are both skipped and X500 is not included.'
    exit 1
}

# ── Resolve the Discovery folder ────────────────────────────────────────────────
$discFolder = $DiscoveryFolder.Trim().Trim('"')
$candidate  = Join-Path $discFolder 'Discovery'
if ((Split-Path $discFolder -Leaf) -ne 'Discovery' -and (Test-Path $candidate)) {
    $discFolder = $candidate
}
if (-not (Test-Path $discFolder)) {
    Log "ERROR: Discovery folder not found: $discFolder"
    exit 1
}

$csvPath = Join-Path $discFolder '13_ProxyAddresses.csv'
if (-not (Test-Path $csvPath)) {
    # Discovery folders written before search-domain.ps1 v2.10.0 numbered this file
    # off-by-one (e.g. 12_ProxyAddresses.csv) — fall back to a wildcard match so older
    # archives still work. See search-domain.ps1 changelog entry 2.10.0.
    $fallback = Get-ChildItem -Path $discFolder -Filter '*_ProxyAddresses.csv' -File -ErrorAction SilentlyContinue |
                Select-Object -First 1
    if ($fallback) {
        $csvPath = $fallback.FullName
        Log "13_ProxyAddresses.csv not found — using older-numbered file instead: $($fallback.Name)"
    }
}
if (-not (Test-Path $csvPath)) {
    Log "ERROR: No *_ProxyAddresses.csv file found in: $discFolder"
    exit 1
}
$csvName = Split-Path $csvPath -Leaf

$rows = @(Import-Csv -Path $csvPath -Encoding UTF8)
if ($rows.Count -eq 0 -or -not ($rows[0].PSObject.Properties.Name -contains 'ProxyAddress')) {
    Log "No proxy address data in $csvName — nothing to restore."
    exit 0
}

if ($Domain) {
    $rows = @($rows | Where-Object { $_.ProxyAddress -like "*@$Domain" })
    Log "Filtered to $($rows.Count) row(s) matching @$Domain"
}

# ── Resolve the destination tenant's onmicrosoft.com domain from Settings > Customers ─────
# The old vanity domain usually isn't verified on the destination tenant yet at this point in
# the migration, so PrimarySmtpAddress/UserPrincipalName from the CSV (both still on the OLD
# domain) won't match anything there — accounts are still on their default onmicrosoft.com
# address. Derived from AccountName rather than the Domain field: Domain is sometimes just a
# short label (seen in the wild: "Bravura", "Caternet") that isn't a usable domain on its own,
# while AccountName is the literal working sign-in address for that tenant.
$tenantDomain = $null
if ($CustomerPrefix) {
    $cfgPath = Join-Path $env:APPDATA 'FlyMigration\config.json'
    if (Test-Path $cfgPath) {
        try {
            $cfg      = Get-Content $cfgPath -Raw | ConvertFrom-Json
            $customer = @($cfg.Customers) | Where-Object { $_.Prefix -and $_.Prefix -eq $CustomerPrefix } | Select-Object -First 1
            if (-not $customer) {
                $customer = @($cfg.Customers) | Where-Object { $_.Prefix -and $_.Prefix.ToLower() -eq $CustomerPrefix.ToLower() } | Select-Object -First 1
            }
            if ($customer -and $customer.AccountName -and $customer.AccountName -match '@(.+)$') {
                $tenantDomain = $Matches[1]
                Log "Customer '$CustomerPrefix' -> destination tenant domain: $tenantDomain (from AccountName)"
            } elseif ($customer) {
                Log "WARNING: Customer '$CustomerPrefix' found but has no usable AccountName in config.json — skipping tenant-domain match"
            } else {
                Log "WARNING: Customer '$CustomerPrefix' not found in config.json — skipping tenant-domain match"
            }
        } catch {
            Log "Could not read config.json for -CustomerPrefix lookup: $($_.Exception.Message.Split([Environment]::NewLine)[0])"
        }
    } else {
        Log "config.json not found at $cfgPath — skipping tenant-domain match"
    }
}

$typeList = @()
if (-not $SkipSMTP) { $typeList += 'SMTP aliases' }
if (-not $SkipSIP)  { $typeList += 'SIP/EUM addresses' }
if ($IncludeX500)   { $typeList += 'X500 addresses' }

Log "=== Restore Proxy Addresses$(if ($WhatIf) { ' [WhatIf]' }) ==="
Log "Discovery folder : $discFolder"
Log "Source CSV       : $csvName ($($rows.Count) row(s))"
Log "Restoring        : $($typeList -join ', ')"

# ── Group rows into one entry per recipient ─────────────────────────────────────
$groups = $rows | Group-Object -Property PrimarySmtpAddress

Log "Found $($groups.Count) distinct recipient(s) to process"

$mod = Get-Module -ListAvailable -Name 'ExchangeOnlineManagement' -ErrorAction SilentlyContinue
if (-not $mod) {
    Log 'ERROR: ExchangeOnlineManagement module is not installed.'
    Log 'Install it with: Install-Module ExchangeOnlineManagement -Scope CurrentUser'
    exit 1
}
Import-Module 'ExchangeOnlineManagement' -ErrorAction Stop

Log 'Connecting to Exchange Online — sign in with the DESTINATION tenant admin account when the browser opens...'
try {
    $cmds = @('Get-Recipient','Get-Mailbox','Set-Mailbox',
              'Get-MailUser','Set-MailUser',
              'Get-DistributionGroup','Set-DistributionGroup',
              'Get-UnifiedGroup','Set-UnifiedGroup',
              'Get-MailContact','Set-MailContact')
    Connect-ExchangeOnline -ShowBanner:$false -CommandName $cmds -DisableWAM -ErrorAction Stop
    Log 'Connected to Exchange Online.'
} catch {
    Log "ERROR: Failed to connect to Exchange Online: $($_.Exception.Message.Split([Environment]::NewLine)[0])"
    exit 1
}

$ok = 0; $fail = 0; $nochange = 0; $notFound = 0

foreach ($group in $groups) {
    $primarySmtp = $group.Name
    $groupRows   = $group.Group
    $upn         = ($groupRows | Select-Object -First 1 -ExpandProperty UserPrincipalName)
    $displayName = ($groupRows | Select-Object -First 1 -ExpandProperty DisplayName)

    if (-not $primarySmtp -and -not $upn) { continue }

    $r = $null
    foreach ($identity in @($primarySmtp, $upn) | Where-Object { $_ }) {
        try {
            $r = Get-Recipient -Identity $identity -ErrorAction Stop
            break
        } catch { }
    }

    $localPart = @($primarySmtp, $upn) | Where-Object { $_ -match '^([^@]+)@' } | ForEach-Object { ($_ -split '@')[0] } | Select-Object -First 1

    # Try the destination tenant's onmicrosoft.com domain next (-CustomerPrefix) — a direct,
    # exact lookup, and the address that's actually correct for an account not yet renamed onto
    # its final vanity domain in the destination tenant.
    if (-not $r -and $localPart -and $tenantDomain) {
        try {
            $r = Get-Recipient -Identity "$localPart@$tenantDomain" -ErrorAction Stop
            if ($r) { Log "  $displayName  [$primarySmtp]  — matched via tenant domain: $($r.PrimarySmtpAddress)" }
        } catch { }
    }

    # Last resort: local-part across any domain. NOTE the leading '*' — EmailAddresses entries
    # are stored type-prefixed (e.g. "smtp:user@domain.com"), so a filter without it only ever
    # matches a value that starts with "user@", which none of them do; that silently made this
    # fallback match nothing, ever, prior to this fix.
    if (-not $r -and $localPart) {
        try {
            $r = Get-Recipient -Filter "EmailAddresses -like '*$localPart@*'" -ErrorAction Stop | Select-Object -First 1
            if ($r) { Log "  $displayName  [$primarySmtp]  — matched by local-part: $($r.PrimarySmtpAddress)" }
        } catch { }
    }

    if (-not $r) {
        Log "  $displayName  [$primarySmtp]  — NOT FOUND in this tenant, skipped"
        $notFound++
        continue
    }

    $current = @($r.EmailAddresses | ForEach-Object { "$_" })
    $currentLower = $current | ForEach-Object { $_.ToLowerInvariant() }
    $toAdd   = [System.Collections.Generic.List[string]]::new()
    $added   = [System.Collections.Generic.List[string]]::new()

    foreach ($row in $groupRows) {
        $type = $row.AddressType
        $addr = $row.ProxyAddress
        if (-not $type -or -not $addr) { continue }

        $wanted = switch ($type) {
            'SMTP' { -not $SkipSMTP }
            'SIP'  { -not $SkipSIP }
            'EUM'  { -not $SkipSIP }
            'X500' { $IncludeX500.IsPresent }
            default { $false }
        }
        if (-not $wanted) { continue }

        $newAddr = "$($type.ToLowerInvariant()):$addr"
        if ($currentLower -contains $newAddr.ToLowerInvariant()) { continue }
        if (($toAdd | ForEach-Object { $_.ToLowerInvariant() }) -contains $newAddr.ToLowerInvariant()) { continue }

        $toAdd.Add($newAddr)
        $added.Add("${type}:$addr")
    }

    if ($toAdd.Count -eq 0) {
        Log "  $displayName  [$primarySmtp]  — no addresses to restore"
        $nochange++
        continue
    }

    $addedList = $added -join ', '

    if ($WhatIf) {
        Log "  $displayName  [$primarySmtp]  — WhatIf: would add $addedList"
        $ok++
        continue
    }

    try {
        $updated = @($current) + @($toAdd)
        switch ($r.RecipientTypeDetails) {
            { $_ -in 'UserMailbox','SharedMailbox','RoomMailbox','EquipmentMailbox' } {
                Set-Mailbox -Identity $r.Identity -EmailAddresses $updated -ErrorAction Stop
            }
            'MailUser' {
                Set-MailUser -Identity $r.Identity -EmailAddresses $updated -ErrorAction Stop
            }
            { $_ -in 'MailUniversalDistributionGroup','MailUniversalSecurityGroup','DynamicDistributionGroup' } {
                Set-DistributionGroup -Identity $r.Identity -EmailAddresses $updated -ErrorAction Stop
            }
            'GroupMailbox' {
                Set-UnifiedGroup -Identity $r.Identity -EmailAddresses $updated -ErrorAction Stop
            }
            'MailContact' {
                Set-MailContact -Identity $r.Identity -EmailAddresses $updated -ErrorAction Stop
            }
            default {
                Log "  $displayName  [$primarySmtp]  — SKIPPED, unhandled recipient type: $($r.RecipientTypeDetails)"
                $nochange++
                continue
            }
        }
        Log "  $displayName  [$primarySmtp]  — added: $addedList"
        $ok++
    } catch {
        Log "  $displayName  [$primarySmtp]  — FAILED: $($_.Exception.Message.Split([Environment]::NewLine)[0])"
        $fail++
    }
}

try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}

Log ''
Log "=== Complete: updated $ok  |  failed $fail  |  unchanged $nochange  |  not found $notFound ==="
Log "Log saved to $logFile"

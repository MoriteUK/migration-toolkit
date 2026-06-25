#Requires -Version 7.0
<#
.SYNOPSIS
    Update-SIPDomain.ps1 — Remove or replace SIP/EUM proxy addresses across EXO recipients.

.PARAMETER OldDomain
    The domain to target (default: tarantosystems.com).

.PARAMETER Mode
    'Remove'  — removes SIP/EUM addresses matching OldDomain (default).
                Temporarily assigns a license, waits for provisioning, removes the
                addresses, then revokes the temporary license.
    'Replace' — replaces the domain suffix with NewDomain, preserving local parts.

.PARAMETER NewDomain
    Replacement domain used only in Replace mode (default: ourvolaris.onmicrosoft.com).

.PARAMETER LicenseSkuId
    GUID of the license SKU to temporarily assign (Remove mode only).
    Leave empty to auto-detect the first available Exchange / Teams SKU.

.PARAMETER WaitMinutes
    Minutes to wait after assigning the license before modifying addresses (default: 5).

.PARAMETER WhatIf
    Preview changes without applying them.

.EXAMPLE
    .\Update-SIPDomain.ps1 -WhatIf
    .\Update-SIPDomain.ps1
    .\Update-SIPDomain.ps1 -Mode Replace -OldDomain contoso.com -NewDomain contoso.onmicrosoft.com
#>

param(
    [string]$OldDomain    = 'tarantosystems.com',
    [ValidateSet('Remove','Replace')]
    [string]$Mode         = 'Remove',
    [string]$NewDomain    = 'ourvolaris.onmicrosoft.com',
    [string]$LicenseSkuId = '',
    [int]$WaitMinutes     = 5,
    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'

$_logDir = Join-Path $PSScriptRoot 'logs'
if (-not (Test-Path $_logDir)) { New-Item -ItemType Directory -Path $_logDir -Force | Out-Null }
$logFile = Join-Path $_logDir "update-sip-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
function Log { param([string]$m) $ts = Get-Date -Format 'HH:mm:ss'; "$ts $m" | Tee-Object -FilePath $logFile -Append | Write-Host }

Log "=== Update-SIPDomain [$Mode]$(if ($WhatIf) { ' [WhatIf]' }) ==="
Log "Target domain : @$OldDomain"
if ($Mode -eq 'Replace') { Log "Replace with  : @$NewDomain" }

# ── Connect Exchange Online ────────────────────────────────────────────────────
$exoCmds = @('Get-Recipient','Get-Mailbox','Set-Mailbox',
             'Get-MailUser','Set-MailUser',
             'Get-DistributionGroup','Set-DistributionGroup',
             'Get-UnifiedGroup','Set-UnifiedGroup',
             'Get-MailContact','Set-MailContact')

Log 'Connecting to Exchange Online...'
Connect-ExchangeOnline -ShowBanner:$false -CommandName $exoCmds -ErrorAction Stop
Log 'Connected to Exchange Online.'

# ── Connect Microsoft Graph (Remove mode only) ────────────────────────────────
if ($Mode -eq 'Remove') {
    Log 'Connecting to Microsoft Graph for license management...'
    Connect-MgGraph -Scopes 'User.ReadWrite.All','Organization.Read.All' -UseDeviceCode -NoWelcome -ErrorAction Stop
    Log 'Connected to Microsoft Graph.'
}

# ── Find affected recipients ──────────────────────────────────────────────────
Log "Searching for SIP/EUM addresses matching @$OldDomain ..."
$sipRecips = @(Get-Recipient -Filter "EmailAddresses -like 'SIP:*@$OldDomain'" -ResultSize Unlimited -ErrorAction Stop)
$eumRecips = @(Get-Recipient -Filter "EmailAddresses -like 'EUM:*@$OldDomain'" -ResultSize Unlimited -ErrorAction Stop)
$all       = (@($sipRecips) + @($eumRecips)) | Sort-Object -Property Identity -Unique
Log "Found $($all.Count) recipient(s) with SIP/EUM addresses on @$OldDomain"

if ($all.Count -eq 0) {
    Log 'Nothing to do.'
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    if ($Mode -eq 'Remove') { Disconnect-MgGraph -ErrorAction SilentlyContinue }
    exit 0
}

# ── Auto-detect license SKU ───────────────────────────────────────────────────
if ($Mode -eq 'Remove') {
    if (-not $LicenseSkuId) {
        Log 'Auto-detecting available license SKU...'
        $preferred = @('TEAMS_EXPLORATORY','EXCHANGEDESKLESS','EXCHANGESTANDARD',
                       'EXCHANGEENTERPRISE','SPE_E1','SPE_E3','O365_BUSINESS_PREMIUM')
        $skus   = Get-MgSubscribedSku -ErrorAction Stop
        $chosen = $null
        foreach ($p in $preferred) {
            $match = $skus | Where-Object {
                $_.SkuPartNumber -eq $p -and
                ($_.PrepaidUnits.Enabled - $_.ConsumedUnits) -gt 0
            }
            if ($match) { $chosen = $match | Select-Object -First 1; break }
        }
        if (-not $chosen) {
            $chosen = $skus | Where-Object { ($_.PrepaidUnits.Enabled - $_.ConsumedUnits) -gt 0 } | Select-Object -First 1
        }
        if (-not $chosen) {
            Log 'ERROR: No available license SKU with free seats found. Specify -LicenseSkuId manually.'
            exit 1
        }
        $LicenseSkuId = [string]$chosen.SkuId
        Log "Using license SKU: $($chosen.SkuPartNumber) ($LicenseSkuId)"
    } else {
        Log "License SKU (provided): $LicenseSkuId"
    }
}

# ── Phase 1 — Add temporary licenses ─────────────────────────────────────────
# Key: UPN, Value: $true if we added the license (so we remove it later)
$addedLicense = @{}

if ($Mode -eq 'Remove' -and -not $WhatIf) {
    Log ''
    Log '--- Phase 1: Adding temporary licenses ---'

    foreach ($r in $all) {
        $upn = $r.PrimarySmtpAddress
        if (-not $upn) { $upn = $r.Identity }
        try {
            $mgUser = Get-MgUser -UserId $upn -Property 'Id,DisplayName,AssignedLicenses' -ErrorAction Stop
            $hasIt  = $mgUser.AssignedLicenses | Where-Object { [string]$_.SkuId -eq $LicenseSkuId }
            if ($hasIt) {
                Log "  $upn — already has this license; will not remove it afterwards"
                $addedLicense[$upn] = $false
            } else {
                Set-MgUserLicense -UserId $mgUser.Id `
                    -AddLicenses @(@{ SkuId = $LicenseSkuId }) `
                    -RemoveLicenses @() -ErrorAction Stop
                Log "  $upn — license added"
                $addedLicense[$upn] = $true
            }
        } catch {
            Log "  $upn — WARNING: could not add license: $($_.Exception.Message.Split([Environment]::NewLine)[0])"
            $addedLicense[$upn] = $false
        }
    }

    Log ''
    Log "Waiting $WaitMinutes minute(s) for license provisioning..."
    $totalSecs = $WaitMinutes * 60
    $tick      = 30
    $elapsed   = 0
    while ($elapsed -lt $totalSecs) {
        Start-Sleep -Seconds $tick
        $elapsed  += $tick
        $remaining = $totalSecs - $elapsed
        if ($remaining -gt 0) { Log "  ... $remaining second(s) remaining" }
    }
    Log 'Wait complete.'
}

# ── Phase 2 — Update addresses ────────────────────────────────────────────────
$ok = 0; $fail = 0; $nochange = 0
$oldEsc = [regex]::Escape($OldDomain)

Log ''
Log '--- Phase 2: Updating addresses ---'

foreach ($r in $all) {
    Log "  $($r.DisplayName)  [$($r.RecipientTypeDetails)]  <$($r.PrimarySmtpAddress)>"

    $current = @($r.EmailAddresses | ForEach-Object { "$_" })
    $updated = [System.Collections.Generic.List[string]]::new()
    $changes = [System.Collections.Generic.List[string]]::new()

    foreach ($addr in $current) {
        $prefix    = ($addr -split ':')[0]
        $isSipEum  = $prefix -iin @('sip','eum')
        $matchesDom = $addr -imatch "@$oldEsc$"

        if ($isSipEum -and $matchesDom) {
            if ($Mode -eq 'Remove') {
                $changes.Add("    REMOVE: $addr")
                # omit from $updated — effectively deletes it
            } else {
                $newAddr = $addr -ireplace "@$oldEsc$", "@$NewDomain"
                $updated.Add($newAddr)
                $changes.Add("    REPLACE: $addr  →  $newAddr")
            }
        } else {
            $updated.Add($addr)
        }
    }

    if ($changes.Count -eq 0) {
        Log "    No matching SIP/EUM addresses found"
        $nochange++
        continue
    }

    foreach ($c in $changes) { Log $c }

    if ($WhatIf) {
        Log "    WhatIf: no changes made."
        $ok++
        continue
    }

    try {
        $addrs = $updated.ToArray()
        switch ($r.RecipientTypeDetails) {
            { $_ -in 'UserMailbox','SharedMailbox','RoomMailbox','EquipmentMailbox' } {
                Set-Mailbox -Identity $r.Identity -EmailAddresses $addrs -ErrorAction Stop
            }
            'MailUser' {
                Set-MailUser -Identity $r.Identity -EmailAddresses $addrs -ErrorAction Stop
            }
            { $_ -in 'MailUniversalDistributionGroup','MailUniversalSecurityGroup','DynamicDistributionGroup' } {
                Set-DistributionGroup -Identity $r.Identity -EmailAddresses $addrs -ErrorAction Stop
            }
            'GroupMailbox' {
                Set-UnifiedGroup -Identity $r.Identity -EmailAddresses $addrs -ErrorAction Stop
            }
            'MailContact' {
                Set-MailContact -Identity $r.Identity -EmailAddresses $addrs -ErrorAction Stop
            }
            default {
                Log "    SKIPPED — unhandled recipient type: $($r.RecipientTypeDetails)"
                $nochange++
                continue
            }
        }
        Log "    Updated OK"
        $ok++
    } catch {
        Log "    FAILED: $($_.Exception.Message.Split([Environment]::NewLine)[0])"
        $fail++
    }
}

# ── Phase 3 — Remove temporary licenses ──────────────────────────────────────
if ($Mode -eq 'Remove' -and -not $WhatIf) {
    Log ''
    Log '--- Phase 3: Removing temporary licenses ---'

    foreach ($r in $all) {
        $upn = $r.PrimarySmtpAddress
        if (-not $upn) { $upn = $r.Identity }
        if ($addedLicense[$upn]) {
            try {
                $mgUser = Get-MgUser -UserId $upn -Property 'Id' -ErrorAction Stop
                Set-MgUserLicense -UserId $mgUser.Id `
                    -AddLicenses @() `
                    -RemoveLicenses @($LicenseSkuId) -ErrorAction Stop
                Log "  $upn — license removed"
            } catch {
                Log "  $upn — WARNING: could not remove license: $($_.Exception.Message.Split([Environment]::NewLine)[0])"
            }
        }
    }
}

Log ''
Log "=== Complete: updated $ok  |  failed $fail  |  skipped/unchanged $nochange ==="
Log "Log saved to $logFile"

Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
if ($Mode -eq 'Remove') {
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
}

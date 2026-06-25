#Requires -Version 7.0
<#
.SYNOPSIS
    Update-SIPDomain.ps1 — Replaces the domain suffix on SIP/EUM proxy addresses
    across all Exchange Online recipients.

.DESCRIPTION
    Finds every recipient that has a SIP: or EUM: address ending in @OldDomain
    and replaces the domain part with @NewDomain, preserving the local-part.
    All other addresses (SMTP, smtp, X500, etc.) are left untouched.

.PARAMETER OldDomain
    The domain to replace (default: tarantosystems.com).

.PARAMETER NewDomain
    The replacement domain (default: ourvolaris.onmicrosoft.com).

.PARAMETER WhatIf
    Show what would change without making any changes.

.EXAMPLE
    .\Update-SIPDomain.ps1 -WhatIf
    .\Update-SIPDomain.ps1
    .\Update-SIPDomain.ps1 -OldDomain contoso.com -NewDomain contoso.onmicrosoft.com
#>

param(
    [string]$OldDomain = 'tarantosystems.com',
    [string]$NewDomain = 'ourvolaris.onmicrosoft.com',
    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'

$_logDir = Join-Path $PSScriptRoot 'logs'
if (-not (Test-Path $_logDir)) { New-Item -ItemType Directory -Path $_logDir -Force | Out-Null }
$logFile = Join-Path $_logDir "update-sip-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
function Log { param([string]$m) $ts = Get-Date -Format 'HH:mm:ss'; "$ts $m" | Tee-Object -FilePath $logFile -Append | Write-Host }

Log "=== Update-SIPDomain$(if ($WhatIf) { ' [WhatIf]' }) ==="
Log "Replacing : @$OldDomain  →  @$NewDomain"

# ── Connect ────────────────────────────────────────────────────────────────────
$cmds = @('Get-Recipient','Get-Mailbox','Set-Mailbox',
          'Get-MailUser','Set-MailUser',
          'Get-DistributionGroup','Set-DistributionGroup',
          'Get-UnifiedGroup','Set-UnifiedGroup',
          'Get-MailContact','Set-MailContact')

Log 'Connecting to Exchange Online...'
Connect-ExchangeOnline -ShowBanner:$false -CommandName $cmds -ErrorAction Stop
Log 'Connected.'

# ── Find affected recipients ───────────────────────────────────────────────────
Log "Searching for SIP/EUM addresses matching @$OldDomain ..."
$pattern = "SIP:*@$OldDomain"
$recipients = @(Get-Recipient -Filter "EmailAddresses -like '$pattern'" -ResultSize Unlimited -ErrorAction Stop)

# Also catch EUM addresses
$eumPattern = "EUM:*@$OldDomain"
$eumRecips  = @(Get-Recipient -Filter "EmailAddresses -like '$eumPattern'" -ResultSize Unlimited -ErrorAction Stop)

# Merge, deduplicate by Identity
$all = @($recipients) + @($eumRecips) | Sort-Object -Property Identity -Unique
Log "Found $($all.Count) recipient(s) with SIP/EUM addresses on @$OldDomain"

if ($all.Count -eq 0) {
    Log 'Nothing to do.'
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    exit 0
}

# ── Process each recipient ─────────────────────────────────────────────────────
$ok = 0; $fail = 0; $nochange = 0
$oldEsc = [regex]::Escape($OldDomain)

foreach ($r in $all) {
    Log "  $($r.DisplayName)  [$($r.RecipientTypeDetails)]  <$($r.PrimarySmtpAddress)>"

    $current = @($r.EmailAddresses | ForEach-Object { "$_" })
    $updated = [System.Collections.Generic.List[string]]::new()
    $changes = [System.Collections.Generic.List[string]]::new()

    foreach ($addr in $current) {
        $prefix = ($addr -split ':')[0]
        if ($prefix -iin @('sip','eum') -and $addr -imatch "@$oldEsc$") {
            $newAddr = $addr -ireplace "@$oldEsc$", "@$NewDomain"
            $updated.Add($newAddr)
            $changes.Add("    $addr  →  $newAddr")
        } else {
            $updated.Add($addr)
        }
    }

    if ($changes.Count -eq 0) {
        Log "    No SIP/EUM addresses to change (already updated?)"
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
        switch ($r.RecipientTypeDetails) {
            { $_ -in 'UserMailbox','SharedMailbox','RoomMailbox','EquipmentMailbox' } {
                Set-Mailbox -Identity $r.Identity -EmailAddresses $updated.ToArray() -ErrorAction Stop
            }
            'MailUser' {
                Set-MailUser -Identity $r.Identity -EmailAddresses $updated.ToArray() -ErrorAction Stop
            }
            { $_ -in 'MailUniversalDistributionGroup','MailUniversalSecurityGroup','DynamicDistributionGroup' } {
                Set-DistributionGroup -Identity $r.Identity -EmailAddresses $updated.ToArray() -ErrorAction Stop
            }
            { $_ -in 'GroupMailbox' } {
                Set-UnifiedGroup -Identity $r.Identity -EmailAddresses $updated.ToArray() -ErrorAction Stop
            }
            'MailContact' {
                Set-MailContact -Identity $r.Identity -EmailAddresses $updated.ToArray() -ErrorAction Stop
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

Log ''
Log "=== Complete: updated $ok  |  failed $fail  |  skipped/unchanged $nochange ==="
Log "Log saved to $logFile"

Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue

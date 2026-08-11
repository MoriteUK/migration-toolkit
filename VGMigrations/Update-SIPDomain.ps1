#Requires -Version 7.0
<#
.SYNOPSIS
    Update-SIPDomain.ps1 — Remove or replace SIP/EUM/IM proxy addresses matching a domain,
    across both on-premises Active Directory and Exchange Online.

.DESCRIPTION
    Some tenants are hybrid (SIP/EUM addresses live in on-prem AD's proxyAddresses attribute
    and sync to the cloud); others are cloud-only (mailboxes already migrated to Exchange
    Online, with no on-prem trace of the address at all). Rather than guessing which applies,
    this script checks both: it searches on-prem AD if the ActiveDirectory module is available,
    and Exchange Online if it can connect, and updates whichever location actually holds a
    matching address for a given recipient. Either source can be unavailable (no RSAT tools,
    or EXO sign-in skipped/failed) without stopping the other from running.

.PARAMETER OldDomain
    The domain to target in SIP/IM/EUM addresses (e.g., intranote.dk).

.PARAMETER Mode
    'Remove'  — removes SIP/EUM/IM addresses matching OldDomain (default).
    'Replace' — replaces the domain suffix with NewDomain, preserving local parts.

.PARAMETER NewDomain
    Replacement domain used only in Replace mode (default: ourvolaris.onmicrosoft.com).

.PARAMETER WhatIf
    Preview changes without applying them.

.PARAMETER SearchBase
    Optional on-prem AD search base (e.g., "OU=Users,DC=domain,DC=com"). If not specified,
    searches the entire domain. Has no effect on the Exchange Online search.

.PARAMETER SkipOnPrem
    Skip the on-premises AD search entirely (useful for known cloud-only tenants).

.PARAMETER SkipExchangeOnline
    Skip the Exchange Online search entirely (useful for known on-prem-only tenants).

.EXAMPLE
    .\Update-SIPDomain.ps1 -OldDomain intranote.dk -WhatIf
    .\Update-SIPDomain.ps1 -OldDomain contoso.com -Mode Remove
    .\Update-SIPDomain.ps1 -OldDomain contoso.com -Mode Replace -NewDomain contoso.onmicrosoft.com
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$OldDomain,
    [ValidateSet('Remove','Replace')]
    [string]$Mode = 'Remove',
    [string]$NewDomain = 'ourvolaris.onmicrosoft.com',
    [string]$SearchBase = '',
    [switch]$SkipOnPrem,
    [switch]$SkipExchangeOnline,
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

$oldEsc = [regex]::Escape($OldDomain)
$onPremOk = 0; $onPremFail = 0; $onPremNoChange = 0
$exoOk = 0; $exoFail = 0; $exoNoChange = 0
$onPremRan = $false; $exoRan = $false

# ── Phase 1 — On-premises Active Directory ────────────────────────────────────
if ($SkipOnPrem) {
    Log ''
    Log '--- On-premises AD: skipped (-SkipOnPrem) ---'
} else {
    Log ''
    Log '--- On-premises AD ---'
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        Log 'ActiveDirectory module loaded.'

        $searchParams = @{
            Filter     = '*'
            Properties = 'UserPrincipalName', 'DisplayName', 'mail', 'proxyAddresses'
        }
        if ($SearchBase) {
            $searchParams['SearchBase'] = $SearchBase
            Log "Search base   : $SearchBase"
        }

        $allUsers = @(Get-ADUser @searchParams -ErrorAction Stop)
        Log "Found $($allUsers.Count) total user(s) in AD"

        $affectedUsers = @($allUsers | Where-Object {
            $_.proxyAddresses | Where-Object {
                $_ -match ('^(sip|im|eum):.*@' + $oldEsc + '$')
            }
        })

        Log "Found $($affectedUsers.Count) on-prem AD user(s) with SIP/IM/EUM addresses on @$OldDomain"
        $onPremRan = $true

        foreach ($user in $affectedUsers) {
            Log "  $($user.DisplayName)  [$($user.UserPrincipalName)]"

            $current = @($user.proxyAddresses)
            $updated = [System.Collections.Generic.List[string]]::new()
            $changes = [System.Collections.Generic.List[string]]::new()

            foreach ($addr in $current) {
                if ($addr -match ('^(sip|im|eum):(.+)@' + $oldEsc + '$')) {
                    $prefix    = $matches[1]
                    $localPart = $matches[2]

                    if ($Mode -eq 'Remove') {
                        $changes.Add("    REMOVE: $addr")
                    } else {
                        $newAddr = "${prefix}:${localPart}@$NewDomain"
                        $updated.Add($newAddr)
                        $changes.Add("    REPLACE: $addr  →  $newAddr")
                    }
                } else {
                    $updated.Add($addr)
                }
            }

            if ($changes.Count -eq 0) {
                Log "    No matching SIP/IM/EUM addresses found"
                $onPremNoChange++
                continue
            }

            foreach ($c in $changes) { Log $c }

            if ($WhatIf) {
                Log "    WhatIf: no changes made."
                $onPremOk++
                continue
            }

            try {
                Set-ADUser -Identity $user.DistinguishedName -Replace @{ proxyAddresses = $updated.ToArray() } -ErrorAction Stop
                Log "    Updated OK"
                $onPremOk++
            } catch {
                Log "    FAILED: $($_.Exception.Message.Split([Environment]::NewLine)[0])"
                $onPremFail++
            }
        }
    } catch {
        Log "WARNING: On-premises AD search skipped: $($_.Exception.Message.Split([Environment]::NewLine)[0])"
        Log "  (Install RSAT tools with: Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0)"
    }
}

# ── Phase 2 — Exchange Online ──────────────────────────────────────────────────
if ($SkipExchangeOnline) {
    Log ''
    Log '--- Exchange Online: skipped (-SkipExchangeOnline) ---'
} else {
    Log ''
    Log '--- Exchange Online ---'
    $exoConnected = $false
    try {
        $mod = Get-Module -ListAvailable -Name 'ExchangeOnlineManagement' -ErrorAction SilentlyContinue
        if (-not $mod) { throw 'ExchangeOnlineManagement module is not installed. Install it with: Install-Module ExchangeOnlineManagement -Scope CurrentUser' }
        Import-Module 'ExchangeOnlineManagement' -ErrorAction Stop

        $cmds = @('Get-Recipient','Get-Mailbox','Set-Mailbox',
                  'Get-MailUser','Set-MailUser',
                  'Get-DistributionGroup','Set-DistributionGroup',
                  'Get-UnifiedGroup','Set-UnifiedGroup',
                  'Get-MailContact','Set-MailContact')
        Log 'Connecting to Exchange Online — sign in when the browser opens...'
        # -DisableWAM: this script is spawned headlessly (no console window) by the Electron app, and
        # WAM hard-requires a parent window handle — without it Connect-ExchangeOnline throws "A window
        # handle must be configured" immediately. -DisableWAM (EXO module >= 3.7.2) uses a normal browser
        # popup instead, which doesn't need a window belonging to this process.
        Connect-ExchangeOnline -ShowBanner:$false -CommandName $cmds -DisableWAM -ErrorAction Stop
        $exoConnected = $true
        Log 'Connected to Exchange Online.'

        Log "Searching for recipients with SIP/EUM/IM addresses matching @$OldDomain ..."
        $sipRecips = @(Get-Recipient -Filter "EmailAddresses -like 'SIP:*@$OldDomain'" -ResultSize Unlimited -ErrorAction Stop)
        $eumRecips = @(Get-Recipient -Filter "EmailAddresses -like 'EUM:*@$OldDomain'" -ResultSize Unlimited -ErrorAction Stop)
        $imRecips  = @(Get-Recipient -Filter "EmailAddresses -like 'IM:*@$OldDomain'"  -ResultSize Unlimited -ErrorAction Stop)
        $affected  = (@($sipRecips) + @($eumRecips) + @($imRecips)) | Sort-Object -Property Identity -Unique

        Log "Found $($affected.Count) Exchange Online recipient(s) with SIP/EUM/IM addresses on @$OldDomain"
        $exoRan = $true

        foreach ($r in $affected) {
            Log "  $($r.DisplayName)  [$($r.RecipientTypeDetails)]  <$($r.PrimarySmtpAddress)>"

            $current = @($r.EmailAddresses | ForEach-Object { "$_" })
            $updated = [System.Collections.Generic.List[string]]::new()
            $changes = [System.Collections.Generic.List[string]]::new()

            foreach ($addr in $current) {
                $prefix     = ($addr -split ':')[0]
                $isSipEumIm = $prefix -iin @('sip', 'eum', 'im')
                $matchesDom = $addr -imatch "@$oldEsc$"

                if ($isSipEumIm -and $matchesDom) {
                    if ($Mode -eq 'Remove') {
                        $changes.Add("    REMOVE: $addr")
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
                Log "    No matching SIP/EUM/IM addresses found"
                $exoNoChange++
                continue
            }

            foreach ($c in $changes) { Log $c }

            if ($WhatIf) {
                Log "    WhatIf: no changes made."
                $exoOk++
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
                        $exoNoChange++
                        continue
                    }
                }
                Log "    Updated OK"
                $exoOk++
            } catch {
                Log "    FAILED: $($_.Exception.Message.Split([Environment]::NewLine)[0])"
                $exoFail++
            }
        }
    } catch {
        Log "WARNING: Exchange Online search skipped: $($_.Exception.Message.Split([Environment]::NewLine)[0])"
    } finally {
        if ($exoConnected) {
            try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
        }
    }
}

# ── Summary ────────────────────────────────────────────────────────────────────
Log ''
if ($onPremRan) {
    Log "On-premises AD : updated $onPremOk  |  failed $onPremFail  |  skipped/unchanged $onPremNoChange"
}
if ($exoRan) {
    Log "Exchange Online: updated $exoOk  |  failed $exoFail  |  skipped/unchanged $exoNoChange"
}
if (-not $onPremRan -and -not $exoRan) {
    Log 'Neither on-premises AD nor Exchange Online could be searched — nothing was done.'
}
Log "Log saved to $logFile"

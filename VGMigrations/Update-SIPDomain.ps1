#Requires -Version 7.0
<#
.SYNOPSIS
    Update-SIPDomain.ps1 — Remove or replace SIP/EUM proxy addresses in on-premises Active Directory.

.PARAMETER OldDomain
    The domain to target in SIP/IM addresses (e.g., intranote.dk).

.PARAMETER Mode
    'Remove'  — removes SIP/EUM/IM addresses matching OldDomain (default).
    'Replace' — replaces the domain suffix with NewDomain, preserving local parts.

.PARAMETER NewDomain
    Replacement domain used only in Replace mode (default: ourvolaris.onmicrosoft.com).

.PARAMETER WhatIf
    Preview changes without applying them.

.PARAMETER SearchBase
    Optional AD search base (e.g., "OU=Users,DC=domain,DC=com"). If not specified, searches entire domain.

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

# ── Import Active Directory Module ────────────────────────────────────────────
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Log 'ActiveDirectory module loaded.'
} catch {
    Log "ERROR: ActiveDirectory module not available: $($_.Exception.Message)"
    Log "Install RSAT tools: Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"
    exit 1
}

# ── Find affected users ───────────────────────────────────────────────────────
Log "Searching for users with SIP/IM/EUM addresses matching @$OldDomain in on-premises AD..."

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

# Filter to only users with SIP/IM/EUM addresses matching the old domain
$affectedUsers = @($allUsers | Where-Object {
    $_.proxyAddresses | Where-Object {
        $_ -match '^(sip|im|eum):.*@' + [regex]::Escape($OldDomain) + '$'
    }
})

Log "Found $($affectedUsers.Count) user(s) with SIP/IM/EUM addresses on @$OldDomain"

if ($affectedUsers.Count -eq 0) {
    Log 'Nothing to do.'
    exit 0
}

# ── Update addresses ──────────────────────────────────────────────────────────
$ok = 0; $fail = 0; $nochange = 0
$oldEsc = [regex]::Escape($OldDomain)

Log ''
Log '--- Updating on-premises AD proxyAddresses ---'

foreach ($user in $affectedUsers) {
    Log "  $($user.DisplayName)  [$($user.UserPrincipalName)]"

    if (-not $user.proxyAddresses) {
        Log "    No proxyAddresses attribute found"
        $nochange++
        continue
    }

    $current = @($user.proxyAddresses)
    $updated = [System.Collections.Generic.List[string]]::new()
    $changes = [System.Collections.Generic.List[string]]::new()

    foreach ($addr in $current) {
        # Check if this is a SIP/IM/EUM address
        if ($addr -match '^(sip|im|eum):(.+)@' + $oldEsc + '$') {
            $prefix = $matches[1]
            $localPart = $matches[2]

            if ($Mode -eq 'Remove') {
                $changes.Add("    REMOVE: $addr")
                # Don't add to $updated — effectively removes it
            } else {
                # Replace mode
                $newAddr = "${prefix}:${localPart}@$NewDomain"
                $updated.Add($newAddr)
                $changes.Add("    REPLACE: $addr  →  $newAddr")
            }
        } else {
            # Keep all other addresses unchanged
            $updated.Add($addr)
        }
    }

    if ($changes.Count -eq 0) {
        Log "    No matching SIP/IM/EUM addresses found"
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
        Set-ADUser -Identity $user.DistinguishedName -Replace @{ proxyAddresses = $updated.ToArray() } -ErrorAction Stop
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

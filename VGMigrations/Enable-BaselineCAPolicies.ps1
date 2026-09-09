#Requires -Version 7.0
<#
.SYNOPSIS
    Enable-BaselineCAPolicies.ps1 — turns on the Conditional Access policies that
    Invoke-TenantBaseline.ps1 Step 11 creates in a disabled state, and (optionally)
    turns Security Defaults off. Run it AFTER a cutover is complete, once the tenant
    is ready to enforce them.

.DESCRIPTION
    Invoke-TenantBaseline.ps1 deploys its 8 Conditional Access policies as
    state = "disabled" on purpose, so a half-migrated tenant is never locked down
    mid-project. This script is the other half: it connects interactively, finds
    those policies by name, and switches them on.

    Default behaviour is deliberately cautious:
      * Policies are moved to "report-only" (enabledForReportingButNotEnforced)
        unless -Enforce is given. Report-only logs what WOULD have happened without
        blocking anyone — leave it a day, check the Entra sign-in logs, then re-run
        with -Enforce.
      * Before any policy is enforced, the break-glass account named by
        -BreakGlassUpn must already be on that policy's Users > Exclude list. A
        policy that doesn't exclude it is left alone (with a warning) unless -Force.
      * -Enforce with no -BreakGlassUpn is refused unless -Force — enabling
        "Require MFA for All Users" / "Block Access Outside Approved Countries" with
        no escape hatch can lock every admin out of the tenant.

    Only the 8 named baseline policies are ever touched — no other Conditional
    Access policy in the tenant is read or changed.

    -DisableSecurityDefaults additionally turns Microsoft Entra Security Defaults
    OFF. Security Defaults and Conditional Access are alternatives, not partners —
    the baseline's CA policies are its replacement — so once they are enforcing you
    normally want Security Defaults off. This only runs when the policies are being
    enforced (-Enforce), so the tenant is never left with neither protection in
    place; pass -Force to override that ordering.

.PARAMETER TenantId
    Tenant domain or GUID to sign into. Optional — only needed if your admin
    account can see more than one tenant.

.PARAMETER BreakGlassUpn
    The emergency-access account that Invoke-TenantBaseline.ps1 Step 11 excluded
    from every policy. Each policy is checked for this exclusion before it is
    enforced.

.PARAMETER Enforce
    Move the policies straight to fully enabled ("on") instead of report-only.

.PARAMETER DisableSecurityDefaults
    Also turn Microsoft Entra Security Defaults off (needs -Enforce, or -Force).

.PARAMETER Force
    Override the break-glass and ordering safety checks. Use only when you have
    confirmed an alternative emergency-access path by hand.

.EXAMPLE
    .\Enable-BaselineCAPolicies.ps1 -BreakGlassUpn breakglass@contoso.onmicrosoft.com
    Moves the 8 baseline policies to report-only (after confirming the break-glass
    exclusion on each). Security Defaults untouched.

.EXAMPLE
    .\Enable-BaselineCAPolicies.ps1 -BreakGlassUpn breakglass@contoso.onmicrosoft.com -Enforce -DisableSecurityDefaults
    Turns the policies fully on and turns Security Defaults off.
#>
# SupportsShouldProcess is here only for -WhatIf. ConfirmImpact is left at the
# default (Medium) so ShouldProcess never auto-prompts in the app's non-interactive
# streamed runner — the real safety guards below are -BreakGlassUpn and -Force.
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$TenantId,
    [string]$BreakGlassUpn,
    [switch]$Enforce,
    [switch]$DisableSecurityDefaults,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

function Write-Step {
    param([string]$Message, [ValidateSet('INFO','OK','WARN','ERROR','SKIP')][string]$Level = 'INFO')
    $color = switch ($Level) { 'OK' {'Green'} 'WARN' {'Yellow'} 'ERROR' {'Red'} 'SKIP' {'DarkYellow'} default {'Cyan'} }
    Write-Host "[$Level] $Message" -ForegroundColor $color
}

function Get-CleanError($ErrorRecord) {
    $lines = @($ErrorRecord.Exception.Message -split "`r?`n" | Where-Object { $_.Trim() })
    if ($lines.Count -eq 0) { return $ErrorRecord.Exception.GetType().Name }
    return ($lines | Select-Object -First 3) -join ' | '
}

# Kept in sync with Invoke-TenantBaseline.ps1 Step 11 and Check-TenantBaselineStatus.ps1.
$BaselineCAPolicyNames = @(
    'Block Access Outside Approved Countries'
    'Block Legacy Authentication'
    'Block Device Code Flow'
    'Require MFA for Admin Portals - 8hr'
    'Require MFA for Admin Roles - 8hr'
    'MFA for All Users - Browser Only - 8hr'
    'Require MFA for All Users'
    'Require MFA for Guest Users - 8hr'
)

Write-Host ""
Write-Host "=================================================" -ForegroundColor Cyan
Write-Host "  Enable Baseline Conditional Access Policies" -ForegroundColor Cyan
Write-Host "  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
Write-Host "=================================================" -ForegroundColor Cyan
Write-Host ""

# --- Graph modules (pinned 2.33.0 for a plain browser sign-in — see Ensure-GraphModules.ps1) ---
. (Join-Path $PSScriptRoot 'Ensure-GraphModules.ps1') -GraphModules @()

$scopes = @('Policy.Read.All', 'Policy.ReadWrite.ConditionalAccess')
if ($DisableSecurityDefaults) { $scopes += 'Policy.ReadWrite.SecurityDefaults' }

try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}

Write-Host "Connecting to Microsoft Graph ($($scopes -join ', '))..." -ForegroundColor Cyan
try {
    $connectArgs = @{ Scopes = $scopes; NoWelcome = $true; ErrorAction = 'Stop' }
    if ($TenantId) { $connectArgs.TenantId = $TenantId }
    Connect-MgGraph @connectArgs | Out-Null
} catch {
    Write-Step "Sign-in failed: $(Get-CleanError $_)" "ERROR"
    exit 1
}

$ctx = Get-MgContext
if (-not $ctx) { Write-Step "No Graph context after sign-in." "ERROR"; exit 1 }
Write-Step "Connected as $($ctx.Account) (tenant $($ctx.TenantId))." "OK"

# =========================================================================
# Enable the baseline Conditional Access policies
# =========================================================================
$targetState = if ($Enforce) { 'enabled' } else { 'enabledForReportingButNotEnforced' }
$modeLabel   = if ($Enforce) { 'ENFORCED (on)' } else { 'report-only' }
Write-Step "Target state for each policy: $modeLabel." "INFO"

# Resolve the break-glass account to an object id for the exclusion check.
$bgId = $null
if ($BreakGlassUpn) {
    try {
        $bg = Invoke-MgGraphRequest -Method GET -Uri ("https://graph.microsoft.com/v1.0/users/{0}?`$select=id,userPrincipalName" -f [uri]::EscapeDataString($BreakGlassUpn))
        $bgId = $bg.id
        Write-Step "Break-glass account '$($bg.userPrincipalName)' resolved (id $bgId) — every policy will be checked for this exclusion." "OK"
    } catch {
        Write-Step "Break-glass account '$BreakGlassUpn' not found in this tenant: $(Get-CleanError $_)" "ERROR"
        if (-not $Force) { Write-Step "Re-run with a valid -BreakGlassUpn, or -Force to skip the check." "ERROR"; exit 2 }
    }
}

if ($Enforce -and -not $bgId -and -not $Force) {
    Write-Step "Refusing to ENFORCE policies with no verified break-glass exclusion — this can lock every admin out." "ERROR"
    Write-Step "Re-run with -BreakGlassUpn <account> (recommended), drop -Enforce to stage them as report-only first, or pass -Force to override." "ERROR"
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
    exit 2
}

try {
    $allPolicies = @((Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies').value)
} catch {
    Write-Step "Could not list Conditional Access policies: $(Get-CleanError $_)" "ERROR"
    exit 1
}

# Target set = the 8 named baseline policies only. Nothing else is touched.
$targets = [System.Collections.Generic.List[object]]::new()
foreach ($name in $BaselineCAPolicyNames) {
    $p = $allPolicies | Where-Object { $_.displayName -eq $name } | Select-Object -First 1
    if ($p) { $targets.Add($p) } else { Write-Step "Baseline policy not found in tenant: '$name' (was Step 11 run?)." "WARN" }
}

if ($targets.Count -eq 0) {
    Write-Step "None of the 8 baseline policies exist in this tenant — run Tenant Baseline Config (Step 11) first." "WARN"
}

$changed = 0; $skipped = 0; $blocked = 0; $failed = 0
foreach ($p in $targets) {
    $name = $p.displayName

    if ($p.state -eq $targetState) {
        Write-Step "'$name' is already $modeLabel." "SKIP"; $skipped++; continue
    }
    if ($p.state -eq 'enabled' -and -not $Enforce) {
        Write-Step "'$name' is already fully enabled — not downgrading it to report-only." "SKIP"; $skipped++; continue
    }

    # Break-glass exclusion guard.
    if ($bgId) {
        $excl = @($p.conditions.users.excludeUsers)
        if ($excl -notcontains $bgId) {
            if ($Force) {
                Write-Step "'$name' does NOT exclude the break-glass account — proceeding anyway (-Force)." "WARN"
            } else {
                Write-Step "'$name' does NOT exclude the break-glass account '$BreakGlassUpn' — left unchanged. Add it to the policy's Users > Exclude list, or re-run with -Force." "WARN"
                $blocked++; continue
            }
        }
    }

    if ($PSCmdlet.ShouldProcess($name, "Set Conditional Access state to '$targetState'")) {
        try {
            Invoke-MgGraphRequest -Method PATCH `
                -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies/$($p.id)" `
                -Body (@{ state = $targetState } | ConvertTo-Json) -ContentType 'application/json'
            Write-Step "'$name': $($p.state) -> $targetState." "OK"; $changed++
        } catch {
            Write-Step "'$name': failed to update — $(Get-CleanError $_)" "ERROR"; $failed++
        }
    }
}

# =========================================================================
# Optionally turn Security Defaults off (CA policies are its replacement)
# =========================================================================
if ($DisableSecurityDefaults) {
    Write-Host ""
    if (-not $Enforce -and -not $Force) {
        Write-Step "Skipping -DisableSecurityDefaults: the policies are only in report-only, so turning Security Defaults off now would leave the tenant with no MFA enforcement. Re-run with -Enforce (or -Force)." "WARN"
    } else {
        $sdUri = 'https://graph.microsoft.com/v1.0/policies/identitySecurityDefaultsEnforcementPolicy'
        try {
            $sd = Invoke-MgGraphRequest -Method GET -Uri $sdUri
            if (-not $sd.isEnabled) {
                Write-Step "Security Defaults is already OFF." "SKIP"
            } elseif ($PSCmdlet.ShouldProcess("this tenant", "Turn Microsoft Entra Security Defaults OFF")) {
                Invoke-MgGraphRequest -Method PATCH -Uri $sdUri -Body (@{ isEnabled = $false } | ConvertTo-Json) -ContentType 'application/json'
                Write-Step "Security Defaults is now OFF — the Conditional Access policies are the replacement." "OK"
            }
        } catch {
            Write-Step "Failed to change Security Defaults: $(Get-CleanError $_)" "ERROR"
            $failed++
        }
    }
}

Write-Host ""
Write-Host "-------------------------------------------------" -ForegroundColor Cyan
Write-Step "Changed: $changed   Already there: $skipped   Left for missing break-glass: $blocked   Failed: $failed" "INFO"
if (-not $Enforce -and $changed -gt 0) {
    Write-Step "Policies are in REPORT-ONLY. Review Entra > Sign-in logs, then re-run with -Enforce to turn them fully on." "INFO"
}
if ($blocked -gt 0) {
    Write-Step "$blocked policy(ies) were left unchanged because the break-glass account isn't excluded from them." "WARN"
}

try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
if ($failed -gt 0) { exit 1 }

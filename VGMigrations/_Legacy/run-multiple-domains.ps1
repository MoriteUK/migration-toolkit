#Requires -Version 7.0
<#
.SYNOPSIS
    run-multiple-domains.ps1 — Runs search-domain.ps1 against a list of domains sequentially
    in a single PowerShell session, so token caching across Graph + EXO eliminates per-domain
    sign-in prompts. Designed for unattended/overnight execution.

.DESCRIPTION
    PROBLEM IT SOLVES
    Running search-domain.ps1 once per domain manually means signing into Graph + EXO + SPO
    + Power Platform every single time — fine for one or two domains, painful for ten.

    HOW THIS WORKS
    The orchestrator dot-sources search-domain.ps1 in the current PowerShell session,
    so Graph and EXO tokens cached after the first domain are reused for every domain
    after that. Combined with the tenant-sites cache (export-tenant-sites.ps1) and
    -SkipPowerPlatform, you can run dozens of domains with a SINGLE upfront sign-in.

    For truly unattended overnight runs:
      1. Run export-tenant-sites.ps1 first (interactive, ~5 min)
      2. Run this orchestrator with -SkipPowerPlatform set
      3. Sign in to Graph + EXO when prompted on the very first domain
      4. Walk away — every subsequent domain runs without prompts

.PARAMETER Domains
    Array of domains to scan. Mutually exclusive with -DomainFile.
    Example: @('expretio.com','fara.no','contoso.com')

.PARAMETER DomainFile
    Path to a text file with one domain per line. Lines starting with # or empty lines are
    ignored. Mutually exclusive with -Domains.

.PARAMETER Hybrid
    Pass -Hybrid through to each domain's scan.

.PARAMETER IncludeMembers
    Pass -IncludeMembers through to each domain's scan.

.PARAMETER BusinessUnitId
    Pass -BusinessUnitId through to each domain's scan.

.PARAMETER SkipPowerPlatform
    Pass -SkipPowerPlatform through to each domain's scan. STRONGLY RECOMMENDED for
    unattended overnight runs because the Power Platform sign-in cannot be bypassed
    and will hang waiting for input.

.PARAMETER ContinueOnError
    If a single domain fails, log it and continue with the next one instead of aborting
    the whole batch. Default behaviour is to stop on first error.

.PARAMETER WhatIf
    Show the planned execution order without running anything.

.NOTES
    Version    : 1.0.0
    Last edit  : 2026-05-09
    Author     : Andrew White / Claude (Anthropic)
    Companion  : Pairs with search-domain.ps1 v2.8.0 or later.

    REQUIRES the discovery script (search-m365domain-merged.ps1 / search-domain.ps1)
    to be present in the same folder as this orchestrator.

.EXAMPLE
    .\run-multiple-domains.ps1 -Domains @('expretio.com','fara.no') -Hybrid -SkipPowerPlatform

.EXAMPLE
    .\run-multiple-domains.ps1 -DomainFile .\domains.txt -Hybrid -IncludeMembers -SkipPowerPlatform -ContinueOnError

    Where domains.txt contains:
      # weekly migration scans
      expretio.com
      fara.no
      contoso.onmicrosoft.com

.EXAMPLE
    .\run-multiple-domains.ps1 -Domains @('a.com','b.com') -WhatIf

    Show what would run without actually running anything.
#>

[CmdletBinding(DefaultParameterSetName = 'File', SupportsShouldProcess = $true)]
param(
    [Parameter(ParameterSetName = 'List', Mandatory = $true)][string[]]$Domains,
    [Parameter(ParameterSetName = 'File', Mandatory = $false)][string]$DomainFile,
    [switch]$Hybrid,
    [switch]$IncludeMembers,
    [string]$BusinessUnitId,
    [switch]$SkipPowerPlatform,
    [switch]$ContinueOnError
)

$ErrorActionPreference = if ($ContinueOnError) { 'Continue' } else { 'Stop' }

# ─────────────────────────────────────────────────────────────
# RESOLVE DISCOVERY SCRIPT PATH
# ─────────────────────────────────────────────────────────────
$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$discovery = $null
foreach ($candidate in @('search-domain.ps1', 'search-m365domain-merged.ps1', 'search-m365domain.ps1')) {
    $tryPath = Join-Path $scriptRoot $candidate
    if (Test-Path $tryPath) { $discovery = (Resolve-Path $tryPath).Path; break }
}
if (-not $discovery) {
    Write-Host "ERROR: discovery script not found in $scriptRoot" -ForegroundColor Red
    Write-Host "Expected one of: search-domain.ps1 | search-m365domain-merged.ps1 | search-m365domain.ps1" -ForegroundColor Red
    exit 1
}

# ─────────────────────────────────────────────────────────────
# RESOLVE DOMAIN LIST
# ─────────────────────────────────────────────────────────────
if ($PSCmdlet.ParameterSetName -eq 'File') {
    # If -DomainFile wasn't passed, default to domains.txt next to this script
    if (-not $DomainFile) {
        $defaultFile = Join-Path $scriptRoot 'domains.txt'
        if (Test-Path $defaultFile) {
            $DomainFile = $defaultFile
            Write-Host "Using default domain list: $DomainFile" -ForegroundColor DarkGray
        } else {
            Write-Host "ERROR: no domain source supplied." -ForegroundColor Red
            Write-Host "  Either pass -Domains @('a.com','b.com') or -DomainFile <path>," -ForegroundColor Red
            Write-Host "  or place a 'domains.txt' file next to run-multiple-domains.ps1." -ForegroundColor Red
            exit 1
        }
    }
    if (-not (Test-Path $DomainFile)) {
        Write-Host "ERROR: -DomainFile not found: $DomainFile" -ForegroundColor Red
        exit 1
    }
    $Domains = @(
        Get-Content $DomainFile -Encoding UTF8 |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ -and -not $_.StartsWith('#') }
    )
}
$Domains = @($Domains | Where-Object { $_ -and ($_ -match '\.') } | Select-Object -Unique)
if ($Domains.Count -eq 0) {
    Write-Host "ERROR: no valid domains supplied." -ForegroundColor Red
    exit 1
}

# ─────────────────────────────────────────────────────────────
# WRAPPER LOG
# ─────────────────────────────────────────────────────────────
$batchLog = Join-Path $scriptRoot ("run-multiple-domains_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + ".log")
function Write-BatchLog {
    param([string]$Message, [ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level = 'INFO')
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$ts] [$($Level.PadRight(7))] $Message"
    $col = switch ($Level) { 'INFO'{'Cyan'} 'WARN'{'Yellow'} 'ERROR'{'Red'} 'SUCCESS'{'Green'} }
    Write-Host $entry -ForegroundColor $col
    try { Add-Content -Path $batchLog -Value $entry -Encoding UTF8 } catch {}
}

# ─────────────────────────────────────────────────────────────
# PRE-FLIGHT
# ─────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Multi-Domain Batch Orchestrator" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-BatchLog "Discovery script    : $discovery"
Write-BatchLog "Domains to scan     : $($Domains.Count)"
foreach ($i in 0..($Domains.Count - 1)) {
    Write-BatchLog ("  [{0}] {1}" -f ($i + 1), $Domains[$i])
}
Write-BatchLog "Hybrid              : $($Hybrid.IsPresent)"
Write-BatchLog "IncludeMembers      : $($IncludeMembers.IsPresent)"
Write-BatchLog "BusinessUnitId      : $(if ($BusinessUnitId) { $BusinessUnitId } else { '(none)' })"
Write-BatchLog "SkipPowerPlatform   : $($SkipPowerPlatform.IsPresent)"
Write-BatchLog "ContinueOnError     : $($ContinueOnError.IsPresent)"
Write-BatchLog "WhatIf              : $($WhatIfPreference)"
Write-BatchLog "Batch log           : $batchLog"
Write-Host ""

# Caution on unattended runs
if (-not $SkipPowerPlatform) {
    Write-BatchLog "WARNING: -SkipPowerPlatform is NOT set. Each domain will prompt for Power Platform sign-in." -Level WARN
    Write-BatchLog "  For overnight/unattended runs, re-launch with -SkipPowerPlatform." -Level WARN
}

# Surface SPO cache status — will the SPO sign-in be skipped per domain?
$cacheCands = @(Get-ChildItem -Path $scriptRoot -Filter 'tenant-sites_*.json' -File -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -notlike '*.meta.json' })
if ($cacheCands.Count -gt 0) {
    $newest = $cacheCands | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    $age = ((Get-Date) - $newest.LastWriteTime).TotalDays
    if ($age -le 7) {
        Write-BatchLog ("SPO sites cache present: {0} ({1:N1}d old) — SPO sign-in will be SKIPPED for every domain." -f $newest.Name, $age) -Level SUCCESS
    } else {
        Write-BatchLog ("SPO sites cache exists but {0:N1}d old. Each domain will re-prompt for SPO." -f $age) -Level WARN
        Write-BatchLog "  Re-run export-tenant-sites.ps1 to refresh, then re-run this orchestrator." -Level WARN
    }
} else {
    Write-BatchLog "No SPO sites cache found. Each domain will prompt for SPO sign-in." -Level WARN
    Write-BatchLog "  Run export-tenant-sites.ps1 first to avoid this." -Level WARN
}

# WhatIf: stop here
if ($WhatIfPreference) {
    Write-Host ""
    Write-BatchLog "WhatIf mode — exiting without execution."
    exit 0
}

# ─────────────────────────────────────────────────────────────
# RUN EACH DOMAIN
# ─────────────────────────────────────────────────────────────
$batchStart = Get-Date
$results = [System.Collections.Generic.List[object]]::new()
$idx = 0

foreach ($domain in $Domains) {
    $idx++
    $domStart = Get-Date
    Write-Host ""
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host "  Domain $idx / $($Domains.Count): $domain" -ForegroundColor Cyan
    Write-Host "  Started: $($domStart.ToString('HH:mm:ss'))" -ForegroundColor Cyan
    Write-Host "================================================" -ForegroundColor Cyan

    Write-BatchLog "Starting scan for: $domain ($idx/$($Domains.Count))"

    # Build splat for the discovery script
    $splat = @{ Domain = $domain }
    if ($Hybrid.IsPresent)            { $splat.Hybrid            = $true }
    if ($IncludeMembers.IsPresent)    { $splat.IncludeMembers    = $true }
    if ($BusinessUnitId)              { $splat.BusinessUnitId    = $BusinessUnitId }
    if ($SkipPowerPlatform.IsPresent) { $splat.SkipPowerPlatform = $true }

    $status = 'Unknown'
    $errorMsg = $null

    try {
        # Run in a child scope so Set-StrictMode + variable side-effects don't pollute the orchestrator
        & $discovery @splat
        if ($LASTEXITCODE -and $LASTEXITCODE -ne 0) {
            throw "Discovery script exited with code $LASTEXITCODE"
        }
        $status = 'Success'
        Write-BatchLog "Scan succeeded for: $domain" -Level SUCCESS
    } catch {
        $status = 'Failed'
        $errorMsg = $_.Exception.Message.Split([Environment]::NewLine)[0]
        Write-BatchLog "Scan FAILED for ${domain}: $errorMsg" -Level ERROR
        if (-not $ContinueOnError) {
            Write-BatchLog "Aborting batch (-ContinueOnError not set)." -Level ERROR
            $results.Add([PSCustomObject]@{
                Domain   = $domain
                Status   = $status
                StartedAt= $domStart
                Duration = ((Get-Date) - $domStart)
                Error    = $errorMsg
            }) | Out-Null
            break
        }
    }

    $domEnd = Get-Date
    $dur = $domEnd - $domStart
    $results.Add([PSCustomObject]@{
        Domain   = $domain
        Status   = $status
        StartedAt= $domStart
        Duration = $dur
        Error    = $errorMsg
    }) | Out-Null
    Write-BatchLog ("Domain $domain finished in {0:hh\:mm\:ss} ({1})" -f $dur, $status)
}

# ─────────────────────────────────────────────────────────────
# SUMMARY
# ─────────────────────────────────────────────────────────────
$batchDur = (Get-Date) - $batchStart
Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Batch complete in $($batchDur.ToString('hh\:mm\:ss'))" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan

$ok      = @($results | Where-Object { $_.Status -eq 'Success' }).Count
$failed  = @($results | Where-Object { $_.Status -eq 'Failed'  }).Count
$skipped = $Domains.Count - $results.Count

Write-BatchLog "Total runtime  : $($batchDur.ToString('hh\:mm\:ss'))"
Write-BatchLog "  Succeeded    : $ok"
Write-BatchLog "  Failed       : $failed"
Write-BatchLog "  Not started  : $skipped (after a failure with -ContinueOnError off)"

# Per-domain summary
foreach ($r in $results) {
    $line = "  {0,-30} {1,-10} {2}" -f $r.Domain, $r.Status, $r.Duration.ToString('hh\:mm\:ss')
    if ($r.Error) { $line += "   $($r.Error)" }
    $col = if ($r.Status -eq 'Success') { 'Green' } else { 'Red' }
    Write-Host $line -ForegroundColor $col
    Add-Content -Path $batchLog -Value $line -Encoding UTF8
}

# Write a CSV summary too (handy for review later)
$summaryCsv = Join-Path $scriptRoot ("run-multiple-domains_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + ".csv")
$results | Select-Object Domain, Status, StartedAt, @{n='DurationHHMMSS';e={$_.Duration.ToString('hh\:mm\:ss')}}, Error |
    Export-Csv -Path $summaryCsv -NoTypeInformation -Encoding UTF8 -Force
Write-BatchLog "Summary CSV    : $summaryCsv"

if ($failed -gt 0) { exit 1 } else { exit 0 }
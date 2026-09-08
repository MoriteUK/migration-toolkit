#Requires -Version 7.0
<#
.SYNOPSIS
    Batch wrapper for Run-Assessment.ps1 - opens one assessment per domain, sequentially.
.DESCRIPTION
    Run-Assessment.ps1 is interactive (mode menu + VBU Domain / Search Term / VBU ID / SPO
    Admin URL prompts + a final "delete Raw JSON?" prompt). It has no parameters, so this
    wrapper cannot feed answers in - instead it launches Run-Assessment.ps1 in its own new
    console window per domain and waits for that window to close before starting the next.
    Sign in and answer the prompts in each window as it opens.

    The -Domains list is only used to tell you which domain to enter at each window's prompt;
    every other value (Search Term, VBU ID, SPO Admin URL, skip choices) is entered by hand in
    the window itself.
.PARAMETER Domains
    Domain names to assess, one window per domain, in order. Shown as a reminder before each
    window opens.
.PARAMETER ContinueOnError
    Continue to the next domain if one window exits with a non-zero code, instead of stopping.
.PARAMETER SharePointAdminUrl
    Accepted for backward compatibility with existing callers; no longer used (entered in the
    Run-Assessment.ps1 window instead).
.PARAMETER SkipPowerPlatform
    Accepted for backward compatibility; no longer used (this build of Run-Assessment.ps1 has
    no Power Platform stage).
.PARAMETER SkipTeamMemberships
    Accepted for backward compatibility; no longer used (this build has no Team Memberships stage).
.PARAMETER OutputPath
    Accepted for backward compatibility; no longer used (Run-Assessment.ps1 writes next to itself).
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string[]]$Domains,
    [string]$SharePointAdminUrl,
    [switch]$SkipPowerPlatform,
    [switch]$SkipTeamMemberships,
    [switch]$ContinueOnError,
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

$runAssessmentPath = Join-Path $PSScriptRoot 'Run-Assessment.ps1'
if (-not (Test-Path $runAssessmentPath)) {
    Write-Host "Run-Assessment.ps1 not found at: $runAssessmentPath" -ForegroundColor Red
    return
}

Write-Host ''
Write-Host "Batch Assessment - $($Domains.Count) domain(s)" -ForegroundColor Cyan
Write-Host ('=' * 40) -ForegroundColor Cyan
Write-Host 'Each domain opens Run-Assessment.ps1 in its own window - sign in and answer the' -ForegroundColor DarkGray
Write-Host 'prompts there. This window resumes when that window closes.' -ForegroundColor DarkGray

$ok = 0; $fail = 0

foreach ($domain in $Domains) {
    $domain = $domain.Trim().ToLower().TrimStart('@')
    if (-not $domain) { continue }

    Write-Host ''
    Write-Host "=== $domain ===" -ForegroundColor Cyan
    Write-Host "Enter '$domain' at the 'VBU Domain' prompt in the window that opens." -ForegroundColor Yellow

    try {
        $proc = Start-Process -FilePath 'pwsh.exe' `
            -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $runAssessmentPath) `
            -WorkingDirectory $PSScriptRoot -Wait -PassThru

        if ($proc.ExitCode -eq 0) {
            Write-Host "$domain - window closed (exit 0)" -ForegroundColor Green
            $ok++
        }
        else {
            Write-Host "$domain - window closed with exit $($proc.ExitCode)" -ForegroundColor Yellow
            $fail++
            if (-not $ContinueOnError) {
                Write-Host 'Stopping batch (pass -ContinueOnError to keep going).' -ForegroundColor Yellow
                break
            }
        }
    }
    catch {
        Write-Host "Could not launch assessment for ${domain}: $($_.Exception.Message)" -ForegroundColor Red
        $fail++
        if (-not $ContinueOnError) {
            Write-Host 'Stopping batch (pass -ContinueOnError to keep going).' -ForegroundColor Yellow
            break
        }
    }
}

Write-Host ''
Write-Host "Batch complete: $ok window(s) exited cleanly, $fail with a non-zero exit" -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Yellow' })

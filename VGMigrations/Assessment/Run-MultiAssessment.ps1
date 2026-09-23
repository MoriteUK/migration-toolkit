#Requires -Version 7.0
<#
.SYNOPSIS
    Batch wrapper for Run-Assessment.ps1 - opens one assessment per domain, sequentially.
.DESCRIPTION
    Run-Assessment.ps1 now accepts -VbuDomain/-VbuSearchTerm/-VbuId/-SharePointAdminUrl/
    -OutputPath, so this wrapper feeds those in directly per domain instead of just echoing a
    reminder for the operator to retype at each window's prompts. VBU ID is looked up from
    domains.json by domain (same source the discovery-menu.ps1 GUI's single-domain picker
    uses); VBU Search Term defaults to the domain's first label (e.g. "contoso" from
    "contoso.com"). A sign-in window will still appear per domain for Exchange/Graph/SPO auth.
.PARAMETER Domains
    Domain names to assess, one window per domain, in order.
.PARAMETER ContinueOnError
    Continue to the next domain if one window exits with a non-zero code, instead of stopping.
.PARAMETER SharePointAdminUrl
    Passed through to each Run-Assessment.ps1 invocation as -SharePointAdminUrl.
.PARAMETER SkipPowerPlatform
    Accepted for backward compatibility; no longer used (this build of Run-Assessment.ps1 has
    no Power Platform stage).
.PARAMETER SkipTeamMemberships
    Accepted for backward compatibility; no longer used (this build has no Team Memberships stage).
.PARAMETER OutputPath
    Passed through to each Run-Assessment.ps1 invocation as -OutputPath.
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

# Same domains.json lookup the discovery-menu.ps1 GUI uses to auto-fill VBU ID for a
# single-domain run - applied here per-domain so batch runs get it too.
$domainVbuMap    = @{}
$domainsJsonPath = Join-Path $PSScriptRoot '..\domains.json'
if (Test-Path $domainsJsonPath) {
    try {
        foreach ($e in @(Get-Content $domainsJsonPath -Raw -Encoding UTF8 | ConvertFrom-Json)) {
            if ($e.PSObject.Properties['domain'] -and $e.domain) {
                $domainVbuMap[([string]$e.domain).ToLower()] = if ($e.PSObject.Properties['vbuId']) { [string]$e.vbuId } else { '' }
            }
        }
    } catch { Write-Host "Could not load domains.json: $($_.Exception.Message)" -ForegroundColor Yellow }
}

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

    $vbuId      = $domainVbuMap[$domain]
    $searchTerm = ($domain -split '\.')[0]

    Write-Host ''
    Write-Host "=== $domain ===" -ForegroundColor Cyan
    Write-Host "VBU Search Term='$searchTerm'  VBU ID='$vbuId'" -ForegroundColor DarkGray

    $argList = @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $runAssessmentPath,
        '-VbuDomain', $domain,
        '-VbuSearchTerm', $searchTerm
    )
    if ($vbuId)             { $argList += @('-VbuId', $vbuId) }
    if ($SharePointAdminUrl) { $argList += @('-SharePointAdminUrl', $SharePointAdminUrl) }
    if ($OutputPath)        { $argList += @('-OutputPath', $OutputPath) }

    try {
        $proc = Start-Process -FilePath 'pwsh.exe' `
            -ArgumentList $argList `
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

#Requires -Version 7.0
<#
.SYNOPSIS
    Batch wrapper for Run-Assessment.ps1 - runs one assessment per domain, sequentially.
.DESCRIPTION
    Looks up each domain's VBU ID from domains.json (same lookup discovery-menu.ps1 already
    does) and derives VBUSearchTerm from the domain's first label (e.g. "contoso" from
    contoso.com - the same convention search-domain.ps1 used internally as $DomainPrefix),
    then calls Run-Assessment.ps1 in-process per domain with every prompt bypassed.

    Unlike search-domain.ps1's batch mode, sessions are not reused across domains - each
    Run-Assessment.ps1 call connects and disconnects independently in its own finally block.
    Slower than the legacy batch mode, but correct; each domain gets a clean module/session
    state instead of carrying over whatever the previous domain left behind.
.PARAMETER Domains
    Domain names to assess, one assessment per domain, in order.
.PARAMETER ContinueOnError
    Continue to the next domain if one assessment throws, instead of stopping the batch.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string[]]$Domains,
    [string]$SharePointAdminUrl,
    [switch]$SkipPowerPlatform,
    [switch]$ContinueOnError,
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

$runAssessmentPath = Join-Path $PSScriptRoot 'Run-Assessment.ps1'
if (-not (Test-Path $runAssessmentPath)) {
    Write-Host "Run-Assessment.ps1 not found at: $runAssessmentPath" -ForegroundColor Red
    return
}

# domains.json lives one level up, alongside discovery-menu.ps1 - same lookup it performs itself
$domainsJsonPath = Join-Path $PSScriptRoot '..\domains.json'
$domainVbuMap = @{}
if (Test-Path $domainsJsonPath) {
    try {
        $entries = @(Get-Content $domainsJsonPath -Raw -Encoding UTF8 | ConvertFrom-Json)
        foreach ($e in $entries) {
            $d = if ($e.PSObject.Properties['domain']) { [string]$e.domain } else { $null }
            $v = if ($e.PSObject.Properties['vbuId'])  { [string]$e.vbuId  } else { '' }
            if ($d) { $domainVbuMap[$d.ToLower()] = $v }
        }
    }
    catch {
        Write-Host "Could not load domains.json: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

Write-Host ''
Write-Host "Batch Assessment - $($Domains.Count) domain(s)" -ForegroundColor Cyan
Write-Host ('=' * 40) -ForegroundColor Cyan

$ok = 0; $fail = 0

foreach ($domain in $Domains) {
    $domain = $domain.Trim().ToLower().TrimStart('@')
    if (-not $domain) { continue }

    Write-Host ''
    Write-Host "=== $domain ===" -ForegroundColor Cyan

    $vbuId         = if ($domainVbuMap.ContainsKey($domain)) { $domainVbuMap[$domain] } else { '' }
    $vbuSearchTerm = ($domain -split '\.')[0]

    $params = @{
        Domain            = $domain
        VBUSearchTerm     = $vbuSearchTerm
        VBUId             = $vbuId
        SkipPowerPlatform = [bool]$SkipPowerPlatform
        DeleteRawJson     = $false
    }
    if ($SharePointAdminUrl) { $params.SharePointAdminUrl = $SharePointAdminUrl }
    if ($OutputPath)         { $params.OutputPath         = $OutputPath }

    try {
        & $runAssessmentPath @params
        $ok++
    }
    catch {
        Write-Host "Assessment failed for ${domain}: $($_.Exception.Message)" -ForegroundColor Red
        $fail++
        if (-not $ContinueOnError) {
            Write-Host 'Stopping batch (pass -ContinueOnError to skip failed domains instead).' -ForegroundColor Yellow
            break
        }
    }
}

Write-Host ''
Write-Host "Batch complete: $ok succeeded, $fail failed" -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Yellow' })

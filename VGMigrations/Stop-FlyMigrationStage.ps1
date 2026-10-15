#Requires -Version 7.0
<#
.SYNOPSIS
    Stops in-progress migration jobs for a customer across all workloads
.PARAMETER CustomerPrefix
    Customer prefix, e.g. "Bravura Security"
.PARAMETER Workloads
    Comma-separated workload names to stop. Leave empty to stop all.
#>
param(
    [Parameter(Mandatory=$true)]
    [string]$CustomerPrefix,

    [Parameter(Mandatory=$false)]
    [string]$Workloads = ""
)

. "$PSScriptRoot\lib.ps1"

$ErrorActionPreference = 'Stop'

Write-Host "`n=== Stop Jobs — $CustomerPrefix ===" -ForegroundColor Cyan

$flyApiCfgPath = Join-Path $env:APPDATA "FlyMigration\config.json"
if (-not (Test-Path $flyApiCfgPath)) { Write-Error "Fly API config not found."; exit 1 }

try {
    $rawCfg = Get-Content $flyApiCfgPath -Raw | ConvertFrom-Json
    $apiUrl = $rawCfg.Url; $clientId = $rawCfg.ClientId
    if ($rawCfg.EncSecret) {
        $ss = $rawCfg.EncSecret | ConvertTo-SecureString
        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ss)
        $clientSecret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    } else { Write-Error "Client secret not found."; exit 1 }
} catch { Write-Error "Failed to load config: $_"; exit 1 }

try {
    if (-not (Get-Module -Name Fly.Client -ListAvailable)) { Write-Error "Fly.Client module not found."; exit 1 }
    Import-Module Fly.Client -ErrorAction Stop
} catch { Write-Error "Failed to import Fly.Client: $_"; exit 1 }

try {
    Write-Host "Connecting to Fly API..." -ForegroundColor Cyan
    Connect-Fly -Url $apiUrl -ClientId $clientId -ClientSecret $clientSecret -ErrorAction Stop
    Write-Host "Connected" -ForegroundColor Green
} catch { Write-Error "Failed to connect: $_"; exit 1 }

# URL segment per workload key
$typeSegment = @{
    SharePoint = 'sharepoint'
    Exchange   = 'exchange'
    OneDrive   = 'onedrive'
    Teams      = 'teams'
    TeamChat   = 'teamchat'
    Groups     = 'm365group'
}

try {
    $config = Get-FlyConfiguration
    $headers = @{
        Authorization  = "Bearer $($config['AccessToken'])"
        Accept         = 'application/json'
        'Content-Type' = 'application/json'
    }
    $body = @{ isSelectAll = $true; mappingIds = @() } | ConvertTo-Json

    $anyRun = $false
    $workloadFilter = if ($Workloads) { $Workloads -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ } } else { @() }

    foreach ($entry in $script:FlyWorkloadDefs.GetEnumerator()) {
        $workloadKey = $entry.Key
        $projectName = "$CustomerPrefix - $workloadKey"

        if ($workloadFilter.Count -gt 0 -and $workloadKey -notin $workloadFilter) { continue }

        $project = Get-FlyMigrationProject -Name $projectName -ErrorAction SilentlyContinue
        if (-not $project -or -not $project.Id) {
            Write-Host "`n[$workloadKey] Project '$projectName' not found — skipping." -ForegroundColor Yellow
            continue
        }

        $seg = $typeSegment[$workloadKey]
        if (-not $seg) {
            Write-Warning "[$workloadKey] No URL segment mapping — skipping."
            continue
        }

        Write-Host "`n[$workloadKey] Stopping jobs on '$projectName'..." -ForegroundColor Cyan

        try {
            Invoke-RestMethod -Method DELETE `
                -Uri "$apiUrl/projects/$seg/$($project.Id)/migrations" `
                -Headers $headers `
                -Body $body `
                -ErrorAction Stop | Out-Null
            Write-Host "  ✓ Stop requested for $workloadKey" -ForegroundColor Green
            $anyRun = $true
        } catch {
            $errMsg = "$_"
            if ($errMsg -match 'no.*job|not.*running|nothing' -or $_.Exception.Response.StatusCode -eq 404) {
                Write-Host "  ~ $workloadKey — no jobs running." -ForegroundColor Yellow
            } else {
                Write-Warning "  ✗ $workloadKey failed: $errMsg"
            }
        }
    }

    if (-not $anyRun) {
        Write-Host "`nNo running jobs found for '$CustomerPrefix'." -ForegroundColor Yellow
    } else {
        Write-Host "`n✓ Stop requests sent for $CustomerPrefix" -ForegroundColor Green
    }

} finally {
    Disconnect-Fly -ErrorAction SilentlyContinue
}

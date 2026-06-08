#Requires -Version 7.0
<#
.SYNOPSIS
    Runs a specific migration stage across all projects for a customer
.PARAMETER CustomerPrefix
    Customer prefix, e.g. "Fara"
.PARAMETER Stage
    Stage to run: Verify | PreScan | FullMigration | IncrementalMigration
#>
param(
    [Parameter(Mandatory=$true)]
    [string]$CustomerPrefix,

    [Parameter(Mandatory=$true)]
    [ValidateSet('Verify','PreScan','FullMigration','IncrementalMigration')]
    [string]$Stage,

    [Parameter(Mandatory=$false)]
    [string]$Workloads = ""
)

. "$PSScriptRoot\lib.ps1"

$ErrorActionPreference = 'Stop'

$stageLabel = @{
    Verify               = 'Verify'
    PreScan              = 'Pre-Scan'
    FullMigration        = 'Full Migration'
    IncrementalMigration = 'Incremental Migration'
}[$Stage]

Write-Host "`n=== $stageLabel — $CustomerPrefix ===" -ForegroundColor Cyan

# Load config
$flyApiCfgPath = Join-Path $env:APPDATA "FlyMigration\config.json"
if (-not (Test-Path $flyApiCfgPath)) { Write-Error "Fly API config not found. Configure in Settings first."; exit 1 }

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

try {
    $anyRun = $false
    $workloadFilter = if ($Workloads) { $Workloads -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ } } else { @() }

    foreach ($entry in $script:FlyWorkloadDefs.GetEnumerator()) {
        $workloadKey = $entry.Key   # e.g. 'Exchange', 'TeamChat'
        $def         = $entry.Value
        $projectName = "$CustomerPrefix - $workloadKey"

        if ($workloadFilter.Count -gt 0 -and $workloadKey -notin $workloadFilter) { continue }

        $project = Get-FlyMigrationProject -Name $projectName -ErrorAction SilentlyContinue
        if (-not $project -or -not $project.Id) {
            Write-Host "`n[$workloadKey] Project '$projectName' not found — skipping." -ForegroundColor Yellow
            continue
        }

        $mappingCount = [int]($project.mappingTotalCount ?? 0)
        if ($mappingCount -eq 0) {
            Write-Host "`n[$workloadKey] No mappings imported — skipping." -ForegroundColor Yellow
            continue
        }

        Write-Host "`n[$workloadKey] $stageLabel on '$projectName'..." -ForegroundColor Cyan

        try {
            switch ($Stage) {
                'Verify' {
                    $cmd = $def.Verify
                    if (-not $cmd) { Write-Warning "No Verify command for $workloadKey"; continue }
                    & $cmd -Project $projectName -ErrorAction Stop
                }
                'PreScan' {
                    $cmd = $def.PreScan
                    if (-not $cmd) { Write-Warning "Pre-Scan not available for $workloadKey — skipping."; continue }
                    & $cmd -Project $projectName -ErrorAction Stop
                }
                'FullMigration' {
                    & $def.Start -Project $projectName -Mode FullMigration -ErrorAction Stop
                }
                'IncrementalMigration' {
                    & $def.Start -Project $projectName -Mode IncrementalMigration -ErrorAction Stop
                }
            }
            Write-Host "  ✓ $stageLabel started for $workloadKey" -ForegroundColor Green
            $anyRun = $true
        } catch {
            $errMsg = "$_"
            if ($errMsg -match 'already in progress') {
                Write-Host "  ~ $workloadKey skipped — mappings already in progress." -ForegroundColor Yellow
            } elseif ($errMsg -match 'do not exist') {
                Write-Host "  ~ $workloadKey skipped — no valid mappings to act on." -ForegroundColor Yellow
            } else {
                Write-Warning "  ✗ $workloadKey failed: $_"
            }
        }
    }

    if (-not $anyRun) {
        Write-Warning "No projects found for '$CustomerPrefix'. Use 'Create Project' first."
        exit 1
    }

    Write-Host "`n✓ $stageLabel complete for $CustomerPrefix" -ForegroundColor Green

} finally {
    Disconnect-Fly -ErrorAction SilentlyContinue
}

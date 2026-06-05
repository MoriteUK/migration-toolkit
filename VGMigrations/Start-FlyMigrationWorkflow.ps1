#Requires -Version 7.0
<#
.SYNOPSIS
    Orchestrates a complete AvePoint Fly migration workflow
.DESCRIPTION
    Creates project, imports mappings, runs pre-scan, and starts migration
.PARAMETER CustomerPrefix
    Customer prefix (e.g., "Contoso")
.PARAMETER Workload
    Workload type: SharePoint, Exchange, OneDrive, Teams, TeamChat, Groups
.PARAMETER MappingFile
    Path to CSV file containing mappings
.PARAMETER StopAt
    Stop workflow at stage: CreateProject, ImportMappings, PreScan, Verification, StartMigration
.PARAMETER SkipPreScan
    Skip the pre-scan step
.PARAMETER SkipVerification
    Skip the verification step
.EXAMPLE
    .\Start-FlyMigrationWorkflow.ps1 -CustomerPrefix "Contoso" -Workload Exchange -MappingFile "C:\mappings\exchange.csv"
.EXAMPLE
    .\Start-FlyMigrationWorkflow.ps1 -CustomerPrefix "Contoso" -Workload SharePoint -MappingFile "C:\mappings\spo.csv" -StopAt PreScan
#>
param(
    [Parameter(Mandatory=$true)]
    [string]$CustomerPrefix,

    [Parameter(Mandatory=$true)]
    [ValidateSet('SharePoint', 'Exchange', 'OneDrive', 'Teams', 'TeamChat', 'Groups')]
    [string]$Workload,

    [Parameter(Mandatory=$true)]
    [string]$MappingFile,

    [Parameter(Mandatory=$false)]
    [ValidateSet('CreateProject', 'ImportMappings', 'PreScan', 'Verification', 'StartMigration')]
    [string]$StopAt,

    [Parameter(Mandatory=$false)]
    [switch]$SkipPreScan,

    [Parameter(Mandatory=$false)]
    [switch]$SkipVerification
)

. "$PSScriptRoot\lib.ps1"

$ErrorActionPreference = 'Stop'

Write-Host "`n╔══════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║    AvePoint Fly Migration Workflow                      ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

Write-Host "`nCustomer: $CustomerPrefix" -ForegroundColor White
Write-Host "Workload: $Workload" -ForegroundColor White
Write-Host "Mapping File: $MappingFile" -ForegroundColor White

$projectName = "$CustomerPrefix - $Workload"
$logFile = Join-Path $PSScriptRoot "logs\migration-$CustomerPrefix-$Workload-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

# Ensure logs directory exists
$logsDir = Join-Path $PSScriptRoot "logs"
if (-not (Test-Path $logsDir)) {
    New-Item -ItemType Directory -Path $logsDir -Force | Out-Null
}

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logMessage = "[$timestamp] [$Level] $Message"
    Add-Content -Path $logFile -Value $logMessage

    switch ($Level) {
        'ERROR' { Write-Host $Message -ForegroundColor Red }
        'WARNING' { Write-Host $Message -ForegroundColor Yellow }
        'SUCCESS' { Write-Host $Message -ForegroundColor Green }
        default { Write-Host $Message -ForegroundColor White }
    }
}

# Start transcript
Start-Transcript -Path $logFile -Append

try {
    # Get Fly API configuration
    Write-Log "`n[1/6] Loading configuration..." "INFO"
    $flyApiCfgPath = Join-Path $env:APPDATA "FlyMigration\config.json"
    if (-not (Test-Path $flyApiCfgPath)) {
        throw "Fly API configuration not found. Please configure in Settings first."
    }

    $rawCfg = Get-Content $flyApiCfgPath -Raw | ConvertFrom-Json
    $apiUrl = $rawCfg.Url
    $clientId = $rawCfg.ClientId

    if ($rawCfg.EncSecret) {
        $secureSecret = $rawCfg.EncSecret | ConvertTo-SecureString
        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureSecret)
        $clientSecret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    } else {
        throw "Client secret not found in configuration"
    }

    Write-Log "✓ Configuration loaded" "SUCCESS"

    # Import Fly.Client module
    Write-Log "`n[2/6] Loading Fly.Client module..." "INFO"
    if (-not (Get-Module -Name Fly.Client -ListAvailable)) {
        throw "Fly.Client module not found. Install: Install-Module -Name Fly.Client"
    }
    Import-Module Fly.Client -ErrorAction Stop
    Write-Log "✓ Module loaded" "SUCCESS"

    # Connect to Fly API
    Write-Log "`n[3/6] Connecting to Fly API..." "INFO"
    Connect-Fly -Url $apiUrl -ClientId $clientId -ClientSecret $clientSecret -ErrorAction Stop
    Write-Log "✓ Connected to Fly API" "SUCCESS"

    # STAGE 1: Create Project
    Write-Log "`n[4/6] Creating project: $projectName..." "INFO"

    $existingProject = Get-FlyMigrationProject -Name $projectName -ErrorAction SilentlyContinue
    if ($existingProject) {
        Write-Log "Project already exists (ID: $($existingProject.Id))" "WARNING"
    } else {
        $newProject = New-FlyMigrationProject -Name $projectName -ErrorAction Stop
        Write-Log "✓ Project created (ID: $($newProject.Id))" "SUCCESS"
    }

    if ($StopAt -eq 'CreateProject') {
        Write-Log "`nStopping at CreateProject stage as requested" "INFO"
        return
    }

    # STAGE 2: Import Mappings
    Write-Log "`n[5/6] Importing mappings from: $MappingFile..." "INFO"

    if (-not (Test-Path $MappingFile)) {
        throw "Mapping file not found: $MappingFile"
    }

    $importCmd = $script:FlyWorkloadDefs[$Workload].Import
    & $importCmd -Project $projectName -Path $MappingFile -ErrorAction Stop
    Write-Log "✓ Mappings imported" "SUCCESS"

    if ($StopAt -eq 'ImportMappings') {
        Write-Log "`nStopping at ImportMappings stage as requested" "INFO"
        return
    }

    # STAGE 3: Pre-Scan
    if (-not $SkipPreScan) {
        Write-Log "`n[6/6] Starting pre-scan..." "INFO"

        $preScanCmd = $script:FlyWorkloadDefs[$Workload].PreScan
        if ($preScanCmd) {
            & $preScanCmd -Project $projectName -ErrorAction Stop
            Write-Log "✓ Pre-scan started" "SUCCESS"
            Write-Log "Monitor progress in AvePoint Fly portal" "INFO"
        } else {
            Write-Log "Pre-scan not available for $Workload" "WARNING"
        }

        if ($StopAt -eq 'PreScan') {
            Write-Log "`nStopping at PreScan stage as requested" "INFO"
            return
        }
    }

    # STAGE 4: Verification
    if (-not $SkipVerification) {
        Write-Log "`n[7/6] Starting verification..." "INFO"

        $verifyCmd = $script:FlyWorkloadDefs[$Workload].Verify
        if ($verifyCmd) {
            & $verifyCmd -Project $projectName -ErrorAction Stop
            Write-Log "✓ Verification started" "SUCCESS"
        } else {
            Write-Log "Verification not available for $Workload" "WARNING"
        }

        if ($StopAt -eq 'Verification') {
            Write-Log "`nStopping at Verification stage as requested" "INFO"
            return
        }
    }

    # STAGE 5: Start Migration
    if ($StopAt -ne 'StartMigration') {
        Write-Log "`n[8/6] Starting migration..." "INFO"

        $startCmd = $script:FlyWorkloadDefs[$Workload].Start
        & $startCmd -Project $projectName -ErrorAction Stop
        Write-Log "✓ Migration started" "SUCCESS"
        Write-Log "Monitor progress in AvePoint Fly portal or using Get-MigrationData.ps1" "INFO"
    }

    Write-Log "`n╔══════════════════════════════════════════════════════════╗" "SUCCESS"
    Write-Log "║    Migration Workflow Completed Successfully            ║" "SUCCESS"
    Write-Log "╚══════════════════════════════════════════════════════════╝" "SUCCESS"

    Write-Log "`nLog file: $logFile" "INFO"

} catch {
    Write-Log "`n✗ Workflow failed: $_" "ERROR"
    Write-Log "Check log file for details: $logFile" "ERROR"
    exit 1
} finally {
    Stop-Transcript
    Disconnect-Fly -ErrorAction SilentlyContinue
}

#Requires -Version 7.0
<#
.SYNOPSIS
    Runs the complete Fly migration workflow for one workload
.DESCRIPTION
    1. Connects to Fly and looks up the destination connection by customer name
    2. Creates the migration project if it doesn't already exist
    3. Imports the CSV mappings
    4. Starts a pre-scan (unless -SkipPreScan)
    5. Starts the migration
.PARAMETER CustomerPrefix
    Customer prefix, e.g. "Fara"
.PARAMETER Workload
    Workload type: SharePoint, Exchange, OneDrive, Teams, TeamChat, Groups
.PARAMETER MappingFile
    Path to CSV file containing mappings
.PARAMETER SkipPreScan
    Skip the pre-scan step
.PARAMETER SkipVerification
    Skip the verification step
#>
param(
    [Parameter(Mandatory=$true)]
    [string]$CustomerPrefix,

    [Parameter(Mandatory=$true)]
    [ValidateSet('SharePoint','Exchange','OneDrive','Teams','TeamChat','Groups')]
    [string]$Workload,

    [Parameter(Mandatory=$true)]
    [string]$MappingFile,

    [Parameter(Mandatory=$false)]
    [string]$CustomerDomain = "",

    [Parameter(Mandatory=$false)]
    [switch]$SkipPreScan,

    [Parameter(Mandatory=$false)]
    [switch]$SkipVerification
)

. "$PSScriptRoot\lib.ps1"

$ErrorActionPreference = 'Stop'

# Workload command maps
$importCmds = @{
    Exchange   = 'Import-FlyExchangeMappings'
    SharePoint = 'Import-FlySharePointMappings'
    OneDrive   = 'Import-FlyOneDriveMappings'
    Teams      = 'Import-FlyTeamsMappings'
    TeamChat   = 'Import-FlyTeamChatMappings'
    Groups     = 'Import-FlyM365GroupMappings'
}
$preScanCmds = @{
    Exchange   = 'Start-FlyExchangePreScan'
    SharePoint = 'Start-FlySharePointPreScan'
    OneDrive   = 'Start-FlyOneDrivePreScan'
    Teams      = 'Start-FlyTeamsPreScan'
    TeamChat   = ''
    Groups     = 'Start-FlyM365GroupPreScan'
}
$startCmds = @{
    Exchange   = 'Start-FlyExchangeMigration'
    SharePoint = 'Start-FlySharePointMigration'
    OneDrive   = 'Start-FlyOneDriveMigration'
    Teams      = 'Start-FlyTeamsMigration'
    TeamChat   = 'Start-FlyTeamChatMigration'
    Groups     = 'Start-FlyM365GroupMigration'
}
$sourceConnMap = @{
    Exchange   = 'OurVolaris - EXO'
    SharePoint = 'OurVolaris - SPO'
    OneDrive   = 'OurVolaris - OneDrive'
    Teams      = 'OurVolaris - MS Teams'
    TeamChat   = 'OurVolaris - Teams Chats'
    Groups     = 'OurVolaris - M365 Groups'
}

$projectName    = "$CustomerPrefix - $Workload"
$sourceConn     = $sourceConnMap[$Workload]

Write-Host "`n╔══════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║    Fly Migration Workflow                               ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host "Project:  $projectName" -ForegroundColor White
Write-Host "Mapping:  $MappingFile" -ForegroundColor White

# Validate mapping file
if (-not (Test-Path $MappingFile)) {
    Write-Error "Mapping file not found: $MappingFile"
    exit 1
}

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

# Load module
try {
    if (-not (Get-Module -Name Fly.Client -ListAvailable)) { Write-Error "Fly.Client module not found."; exit 1 }
    Import-Module Fly.Client -ErrorAction Stop
} catch { Write-Error "Failed to import Fly.Client: $_"; exit 1 }

# Set up log file
$logsDir = Join-Path $env:APPDATA "FlyMigration\Logs"
if (-not (Test-Path $logsDir)) { New-Item -ItemType Directory -Path $logsDir -Force | Out-Null }
$logFile = Join-Path $logsDir "workflow-$CustomerPrefix-$Workload-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

function Write-StepLog {
    param([string]$Message, [string]$Level = 'INFO')
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    "[$ts] [$Level] $Message" | Add-Content -Path $logFile -Encoding UTF8
    $colour = switch ($Level) {
        'ERROR'   { 'Red'    }
        'WARNING' { 'Yellow' }
        'SUCCESS' { 'Green'  }
        default   { 'White'  }
    }
    Write-Host $Message -ForegroundColor $colour
}

try {
    # ── STEP 1: Connect ──────────────────────────────────────────────────────
    Write-StepLog "`n[1/5] Connecting to Fly API..."
    Connect-Fly -Url $apiUrl -ClientId $clientId -ClientSecret $clientSecret -ErrorAction Stop
    Write-StepLog "Connected" 'SUCCESS'

    # ── STEP 2: Find destination connection (auto-creates if missing) ────────
    Write-StepLog "`n[2/5] Locating destination connection..."
    $destConn = Find-FlyDestinationConnection `
        -ProjectName    $projectName `
        -Workload       $Workload `
        -ApiUrl         $apiUrl `
        -CustomerDomain $CustomerDomain

    # ── STEP 3: Ensure project exists ────────────────────────────────────────
    Write-StepLog "`n[3/5] Checking project '$projectName'..."
    $project = Get-FlyMigrationProject -Name $projectName -ErrorAction SilentlyContinue
    if (-not $project -or -not $project.Id) {
        Write-StepLog "Project not found — creating..." 'WARNING'
        $policy = Find-FlyPolicy -Workload $Workload -ApiUrl $apiUrl
        $newProj = New-FlyMigrationProject `
            -Name                  $projectName `
            -SourceConnection      $sourceConn `
            -DestinationConnection $destConn `
            -Policy                $policy `
            -ErrorAction Stop
        Write-StepLog "Project created (ID: $($newProj.Id))" 'SUCCESS'
    } else {
        Write-StepLog "Project found (ID: $($project.Id))" 'SUCCESS'
    }

    # ── STEP 4: Import mappings ───────────────────────────────────────────────
    Write-StepLog "`n[4/5] Importing mappings from: $MappingFile..."
    $importCmd = $importCmds[$Workload]
    & $importCmd -Project $projectName -Path $MappingFile -ErrorAction Stop
    Write-StepLog "Mappings imported" 'SUCCESS'

    # ── STEP 5: Pre-scan ─────────────────────────────────────────────────────
    if (-not $SkipPreScan) {
        $preScanCmd = $preScanCmds[$Workload]
        if ($preScanCmd) {
            Write-StepLog "`n[5/5] Starting pre-scan..."
            & $preScanCmd -Project $projectName -ErrorAction Stop
            Write-StepLog "Pre-scan started — monitor progress in the Fly portal" 'SUCCESS'
        } else {
            Write-StepLog "`n[5/5] Pre-scan not available for $Workload — skipping." 'WARNING'
        }
    } else {
        Write-StepLog "`n[5/5] Pre-scan skipped." 'WARNING'

        # If pre-scan skipped, start migration directly
        if (-not $SkipVerification) {
            Write-StepLog "Starting migration..."
            $startCmd = $startCmds[$Workload]
            & $startCmd -Project $projectName -ErrorAction Stop
            Write-StepLog "Migration started — monitor progress in the Fly portal" 'SUCCESS'
        }
    }

    Write-StepLog "`n✓ Workflow complete for $projectName" 'SUCCESS'
    Write-StepLog "Log: $logFile" 'INFO'

} catch {
    Write-StepLog "`n✗ Workflow failed: $_" 'ERROR'
    Write-StepLog "Log: $logFile" 'ERROR'
    exit 1
} finally {
    Disconnect-Fly -ErrorAction SilentlyContinue
}

#Requires -Version 7.0
<#
.SYNOPSIS
    fly-migrator.ps1 — headless migration runner for the Fly web UI.
    Reads a JSON param file, runs Fly.Client cmdlets, and emits JSON events
    to stdout so server.js can forward them to the browser via WebSocket.

.PARAMETER ParamFile
    Path to a temporary JSON file written by server.js containing:
    {
      "flyUrl":       "https://...",
      "clientId":     "...",
      "clientSecret": "...",
      "prefix":       "Contoso",
      "ops":    { "SharePoint": "migrate", "Exchange": "verify", ... },
      "csvPaths": { "SharePoint": "/tmp/abc.csv", ... }
    }
    The file is deleted by server.js once this script exits.

.NOTES
    ops values: "import" | "prescan" | "verify" | "migrate"
    For Teams Chat, "prescan" is not supported — treated as "import".
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ParamFile
)

$ErrorActionPreference = 'Stop'

# ── Helpers ──────────────────────────────────────────────────────────
function Emit([string]$Event, [string]$Message) {
    $obj = [PSCustomObject]@{ event = $Event; message = $Message }
    Write-Output ($obj | ConvertTo-Json -Compress)
}

function EmitTask([string]$Id, [string]$Status, [string]$Message) {
    $obj = [PSCustomObject]@{ id = $Id; status = $Status; message = $Message }
    Write-Output ($obj | ConvertTo-Json -Compress)
}

# ── Load params ───────────────────────────────────────────────────────
try {
    $p = Get-Content $ParamFile -Raw | ConvertFrom-Json
} catch {
    Emit 'fatal' "Could not read param file: $($_.Exception.Message)"
    exit 1
}

# ── Connect ───────────────────────────────────────────────────────────
Emit 'info' 'Connecting to Fly API...'
try {
    $sec = ConvertTo-SecureString $p.clientSecret -AsPlainText -Force
    Connect-Fly -Url $p.flyUrl -ClientId $p.clientId -ClientSecret $sec | Out-Null
    Emit 'info' 'Connected.'
} catch {
    Emit 'fatal' "Connect-Fly failed: $($_.Exception.Message)"
    exit 2
}

# ── Cmdlet map ────────────────────────────────────────────────────────
$cmdlets = @{
    SharePoint = @{ import = 'Import-FlySharePointMappings'; prescan = 'Start-FlySharePointPreScan';  verify = 'Start-FlySharePointVerification';  migrate = 'Start-FlySharePointMigration'  }
    Exchange   = @{ import = 'Import-FlyExchangeMappings';   prescan = 'Start-FlyExchangePreScan';    verify = 'Start-FlyExchangeVerification';    migrate = 'Start-FlyExchangeMigration'    }
    OneDrive   = @{ import = 'Import-FlyOneDriveMappings';   prescan = 'Start-FlyOneDrivePreScan';    verify = 'Start-FlyOneDriveVerification';    migrate = 'Start-FlyOneDriveMigration'    }
    Teams      = @{ import = 'Import-FlyTeamsMappings';      prescan = 'Start-FlyTeamsPreScan';       verify = 'Start-FlyTeamsVerification';      migrate = 'Start-FlyTeamsMigration'       }
    TeamChat   = @{ import = 'Import-FlyTeamChatMappings';   prescan = $null;                         verify = 'Start-FlyTeamChatVerification';    migrate = 'Start-FlyTeamChatMigration'    }
    Groups     = @{ import = 'Import-FlyM365GroupMappings';  prescan = 'Start-FlyM365GroupPreScan';   verify = 'Start-FlyM365GroupVerification';   migrate = 'Start-FlyM365GroupMigration'   }
}

# ── Workload table from workloads.json (source/destination connections) ─
$workloadsFile = Join-Path $PSScriptRoot 'workloads.json'
$workloadCfg   = @{}
if (Test-Path $workloadsFile) {
    try { $workloadCfg = Get-Content $workloadsFile -Raw | ConvertFrom-Json } catch {}
}

# ── Process each workload ─────────────────────────────────────────────
$ops = $p.ops.PSObject.Properties   # enumerate the ops object
foreach ($entry in $ops) {
    $workload = $entry.Name
    $op       = $entry.Value

    if (-not $op -or $op -eq 'none') { continue }

    $cmds = $cmdlets[$workload]
    if (-not $cmds) {
        EmitTask $workload 'SKIPPED' "Unknown workload key"
        continue
    }

    $csvPath = $p.csvPaths.$workload
    if (-not $csvPath -or -not (Test-Path $csvPath)) {
        EmitTask $workload 'SKIPPED' 'No CSV file provided'
        continue
    }

    # Source / destination from workloads.json
    $cfg      = $workloadCfg.$workload
    $srcConn  = if ($cfg) { $cfg.Source }      else { '' }
    $dstConn  = if ($cfg) { $cfg.Destination } else { '' }
    $policy   = if ($cfg) { $cfg.Policy }      else { '' }

    # TeamChat key maps to "Teams Chat" project name in Fly
    $wlDisplayName = if ($workload -eq 'TeamChat') { 'Teams Chat' } else { $workload }
    $projectName = "$($p.prefix) - $wlDisplayName"

    EmitTask $workload 'WORKING' ''

    try {
        # Import mappings
        $importCmd = $cmds['import']
        Emit 'info' "[$workload] Importing mappings from CSV..."
        & $importCmd `
            -Project               $projectName `
            -Path                  $csvPath `
            -SourceConnection      $srcConn `
            -DestinationConnection $dstConn `
            -ErrorAction Stop | Out-Null

        Emit 'info' "[$workload] Mappings imported."

        # Run the requested operation
        switch ($op) {
            'import' {
                EmitTask $workload 'DONE' 'Mappings imported'
            }
            'prescan' {
                $cmd = $cmds['prescan']
                if (-not $cmd) {
                    Emit 'warn' "[$workload] Pre-scan not supported; skipping."
                } else {
                    Emit 'info' "[$workload] Starting pre-scan..."
                    & $cmd -Project $projectName -ErrorAction Stop | Out-Null
                }
                EmitTask $workload 'DONE' 'Pre-scan started'
            }
            'verify' {
                Emit 'info' "[$workload] Starting verification..."
                & $cmds['verify'] -Project $projectName -ErrorAction Stop | Out-Null
                EmitTask $workload 'DONE' 'Verification started'
            }
            'migrate' {
                Emit 'info' "[$workload] Starting migration..."
                & $cmds['migrate'] -Project $projectName -ErrorAction Stop | Out-Null
                EmitTask $workload 'DONE' 'Migration started'
            }
            default {
                EmitTask $workload 'SKIPPED' "Unknown operation: $op"
            }
        }

    } catch {
        EmitTask $workload 'FAILED' $_.Exception.Message
        Emit 'error' "[$workload] $($_.Exception.Message)"
    }
}

Emit 'done' 'All workloads processed.'

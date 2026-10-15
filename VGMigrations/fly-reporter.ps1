#Requires -Version 7.0
<#
.SYNOPSIS
    fly-reporter.ps1 — fetches migration/mapping reports from the Fly API
    and emits each row as a JSON event so server.js can stream them to the browser.

.PARAMETER ParamFile
    Path to a temporary JSON file written by server.js containing:
    {
      "flyUrl":       "https://...",
      "clientId":     "...",
      "clientSecret": "...",
      "prefix":       "Contoso",
      "reportType":   "migration" | "mapping",
      "workloads":    ["SharePoint", "Exchange", ...]
    }
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ParamFile
)

$ErrorActionPreference = 'Stop'

function Emit([string]$Event, [string]$Message) {
    $obj = [PSCustomObject]@{ event = $Event; message = $Message }
    Write-Output ($obj | ConvertTo-Json -Compress)
}

function EmitRow([string]$Workload, [object]$Row) {
    $obj = [PSCustomObject]@{ event = 'row'; workload = $Workload; data = $Row }
    Write-Output ($obj | ConvertTo-Json -Compress -Depth 5)
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

# ── Report cmdlet map ─────────────────────────────────────────────────
$reportCmdlets = @{
    SharePoint  = @{ migration = 'Export-FlySharePointMigrationReport'; mapping = 'Export-FlySharePointMappingStatus' }
    Exchange    = @{ migration = 'Export-FlyExchangeMigrationReport';   mapping = 'Export-FlyExchangeMappingStatus'   }
    OneDrive    = @{ migration = 'Export-FlyOneDriveMigrationReport';   mapping = 'Export-FlyOneDriveMappingStatus'   }
    Teams       = @{ migration = 'Export-FlyTeamsMigrationReport';      mapping = 'Export-FlyTeamsMappingStatus'      }
    TeamChat    = @{ migration = 'Export-FlyTeamChatMigrationReport';   mapping = 'Export-FlyTeamChatMappingStatus'   }
    Groups      = @{ migration = 'Export-FlyM365GroupMigrationReport';  mapping = 'Export-FlyM365GroupMappingStatus'  }
}

$reportType = if ($p.reportType) { $p.reportType } else { 'migration' }

# ── Process each workload ─────────────────────────────────────────────
foreach ($workload in $p.workloads) {
    $cmds = $reportCmdlets[$workload]
    if (-not $cmds) {
        Emit 'warn' "[$workload] Unknown workload — skipped"
        continue
    }

    $cmdName = $cmds[$reportType]
    if (-not $cmdName) {
        Emit 'warn' "[$workload] No $reportType report cmdlet — skipped"
        continue
    }

    # TeamChat key → project name "Teams Chat"
    $wlDisplayName = if ($workload -eq 'TeamChat') { 'Teams Chat' } else { $workload }
    $projectName   = "$($p.prefix) - $wlDisplayName"

    $tempCsv = [System.IO.Path]::GetTempFileName() -replace '\.tmp$', '.csv'

    try {
        Emit 'info' "[$workload] Fetching $reportType report for project '$projectName'..."
        & $cmdName -Project $projectName -OutFile $tempCsv -ErrorAction Stop | Out-Null

        if (-not (Test-Path $tempCsv)) {
            Emit 'warn' "[$workload] No report file produced — project may have no data yet"
            continue
        }

        $rows = Import-Csv -Path $tempCsv -ErrorAction Stop
        Emit 'info' "[$workload] $($rows.Count) rows retrieved"

        foreach ($row in $rows) {
            EmitRow $workload $row
        }

    } catch {
        Emit 'error' "[$workload] $($_.Exception.Message)"
    } finally {
        if (Test-Path $tempCsv) { Remove-Item $tempCsv -Force -ErrorAction SilentlyContinue }
    }
}

Emit 'done' 'Report export complete.'

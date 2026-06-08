#Requires -Version 7.0
<#
.SYNOPSIS
    Imports mappings into an AvePoint Fly migration project
.DESCRIPTION
    Imports user/mailbox/site mappings from a CSV file into a Fly project
.PARAMETER ProjectName
    Name of the Fly project
.PARAMETER Workload
    Workload type: SharePoint, Exchange, OneDrive, Teams, TeamChat, Groups
.PARAMETER MappingFile
    Path to CSV file containing mappings
.PARAMETER SkipValidation
    Skip validation of mappings before import
.EXAMPLE
    .\Import-FlyMappings.ps1 -ProjectName "Contoso - Exchange" -Workload Exchange -MappingFile "C:\mappings\exchange.csv"
#>
param(
    [Parameter(Mandatory=$true)]
    [string]$ProjectName,

    [Parameter(Mandatory=$true)]
    [ValidateSet('SharePoint', 'Exchange', 'OneDrive', 'Teams', 'TeamChat', 'Groups')]
    [string]$Workload,

    [Parameter(Mandatory=$true)]
    [string]$MappingFile,

    [Parameter(Mandatory=$false)]
    [switch]$SkipValidation
)

. "$PSScriptRoot\lib.ps1"

Write-Host "`nImporting Fly Mappings..." -ForegroundColor Cyan
Write-Host "Project: $ProjectName" -ForegroundColor White
Write-Host "Workload: $Workload" -ForegroundColor White
Write-Host "Mapping File: $MappingFile" -ForegroundColor White

# Validate mapping file exists
if (-not (Test-Path $MappingFile)) {
    Write-Error "Mapping file not found: $MappingFile"
    exit 1
}

# Get Fly API configuration
$flyApiCfgPath = Join-Path $env:APPDATA "FlyMigration\config.json"
if (-not (Test-Path $flyApiCfgPath)) {
    Write-Error "Fly API configuration not found. Please configure in Settings first."
    exit 1
}

try {
    $rawCfg = Get-Content $flyApiCfgPath -Raw | ConvertFrom-Json
    $apiUrl = $rawCfg.Url
    $clientId = $rawCfg.ClientId

    if ($rawCfg.EncSecret) {
        $secureSecret = $rawCfg.EncSecret | ConvertTo-SecureString
        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureSecret)
        $clientSecret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    } else {
        Write-Error "Client secret not found in configuration"
        exit 1
    }
} catch {
    Write-Error "Failed to load Fly API configuration: $_"
    exit 1
}

# Import Fly.Client module
try {
    if (-not (Get-Module -Name Fly.Client -ListAvailable)) {
        Write-Error "Fly.Client module not found. Please install it first: Install-Module -Name Fly.Client"
        exit 1
    }
    Import-Module Fly.Client -ErrorAction Stop
} catch {
    Write-Error "Failed to import Fly.Client module: $_"
    exit 1
}

# Connect to Fly API
try {
    Write-Host "`nConnecting to Fly API..." -ForegroundColor Cyan
    Connect-Fly -Url $apiUrl -ClientId $clientId -ClientSecret $clientSecret -ErrorAction Stop
    Write-Host "Connected successfully" -ForegroundColor Green
} catch {
    Write-Error "Failed to connect to Fly API: $_"
    exit 1
}

# Verify project exists — Get-FlyMigrationProject returns an empty object (not
# null and not an error) when the project doesn't exist, so check .Id explicitly.
$project = Get-FlyMigrationProject -Name $ProjectName -ErrorAction SilentlyContinue
if (-not $project -or -not $project.Id) {
    Write-Error "Project '$ProjectName' not found. Use 'Create Project' first."
    exit 1
}
Write-Host "Project found: $($project.Id)" -ForegroundColor Green

# Pre-validate and clean the CSV — remove rows where source == destination
$cleanedCsvPath = $null
try {
    $allRows = Import-Csv $MappingFile -ErrorAction Stop

    if ($allRows.Count -gt 0) {
        # Detect source/destination column names dynamically — different workloads use
        # different headers (e.g. "Source" vs "Source user" vs "Source site")
        $headers = $allRows[0].PSObject.Properties.Name
        $srcCol  = $headers | Where-Object { $_ -imatch '^source' }      | Select-Object -First 1
        $dstCol  = $headers | Where-Object { $_ -imatch '^destination' } | Select-Object -First 1

        if ($srcCol -and $dstCol) {
            # Pass 1 — drop rows where source == destination
            $cleanRows = $allRows | Where-Object {
                $_.$srcCol -and $_.$dstCol -and
                $_.$srcCol.Trim().ToLower() -ne $_.$dstCol.Trim().ToLower()
            }
            $sameCount = $allRows.Count - $cleanRows.Count
            if ($sameCount -gt 0) {
                Write-Warning "$sameCount row(s) skipped — source and destination are identical:"
                $allRows | Where-Object {
                    $_.$srcCol -and $_.$dstCol -and
                    $_.$srcCol.Trim().ToLower() -eq $_.$dstCol.Trim().ToLower()
                } | ForEach-Object { Write-Host "  SKIP (same): $($_.$srcCol)" -ForegroundColor Yellow }
            }

            # Pass 2 — drop rows where the same source appears more than once
            # (Fly rejects the entire batch if any source is duplicated)
            $seen     = @{}
            $dedupRows = [System.Collections.Generic.List[psobject]]::new()
            $dupCount  = 0
            foreach ($row in $cleanRows) {
                $key = $row.$srcCol.Trim().ToLower()
                if ($seen.ContainsKey($key)) {
                    Write-Host "  SKIP (dup src): $($row.$srcCol) → $($row.$dstCol)" -ForegroundColor Yellow
                    $dupCount++
                } else {
                    $seen[$key] = $true
                    $dedupRows.Add($row)
                }
            }
            if ($dupCount -gt 0) {
                Write-Warning "$dupCount row(s) skipped — duplicate source entries (Fly rejects the whole batch if any source appears twice)."
            }

            $finalRows = $dedupRows.ToArray()
            if ($finalRows.Count -eq 0) {
                Write-Error "No valid rows to import after removing same-identity and duplicate-source rows."
                exit 1
            }

            $totalSkipped = $allRows.Count - $finalRows.Count
            if ($totalSkipped -gt 0) {
                $cleanedCsvPath = [System.IO.Path]::GetTempFileName() -replace '\.tmp$', '.csv'
                $finalRows | Export-Csv $cleanedCsvPath -NoTypeInformation -Encoding UTF8
                Write-Host "$($finalRows.Count) of $($allRows.Count) rows will be imported." -ForegroundColor Cyan
                $MappingFile = $cleanedCsvPath
            }
        } else {
            Write-Warning "Could not detect source/destination columns in CSV — skipping pre-validation."
        }
    }
} catch {
    Write-Error "Failed to validate mapping file: $_"
    exit 1
}

# Import mappings
try {
    Write-Host "`nImporting mappings..." -ForegroundColor Cyan

    $importCmd = $script:FlyWorkloadDefs[$Workload].Import
    if (-not $importCmd) {
        Write-Error "No import command found for workload: $Workload"
        exit 1
    }

    $importParams = @{
        Project = $ProjectName
        Path = $MappingFile
        ErrorAction = 'Stop'
    }

    if ($SkipValidation) {
        $importParams.SkipValidation = $true
    }

    Write-Host "Running: $importCmd" -ForegroundColor Gray
    try {
        & $importCmd @importParams
        Write-Host "`n✓ Mappings imported successfully!" -ForegroundColor Green
    } catch {
        # Fly throws a 500/ProjectMappingDuplicated when all submitted rows already exist.
        # Treat this as "already imported" — not a failure.
        $errText = "$_" + ($_.ErrorDetails.Message ?? '')
        if ($errText -match 'ProjectMappingDuplicated') {
            Write-Warning "All mappings already exist in project '$ProjectName' — nothing new to import."
        } else {
            throw
        }
    }

    # Get mapping count
    $statusCmd = $script:FlyWorkloadDefs[$Workload].Status
    if ($statusCmd) {
        try {
            $tempFile = [System.IO.Path]::GetTempFileName()
            & $statusCmd -Project $ProjectName -OutFile $tempFile -ErrorAction SilentlyContinue | Out-Null

            if (Test-Path $tempFile) {
                $mappings = Import-Csv $tempFile -ErrorAction SilentlyContinue
                $mappingCount = $mappings.Count
                Remove-Item $tempFile -Force -ErrorAction SilentlyContinue

                Write-Host "Total mappings: $mappingCount" -ForegroundColor White
            }
        } catch {
            # Ignore errors getting count
        }
    }

    Write-Host "`n=== Next Steps ===" -ForegroundColor Cyan
    Write-Host "1. Review mappings in AvePoint Fly portal"
    Write-Host "2. Run pre-scan: Start-Fly$($Workload)PreScan -Project '$ProjectName'"
    Write-Host "3. Review scan results"
    Write-Host "4. Start migration: Start-Fly$($Workload)Migration -Project '$ProjectName'"

} catch {
    Write-Error "Failed to import mappings: $_"
    Write-Host "`nTroubleshooting:" -ForegroundColor Yellow
    Write-Host "- Verify CSV file format matches Fly requirements"
    Write-Host "- Check for duplicate mappings"
    Write-Host "- Ensure source/destination values are valid"
    exit 1
} finally {
    Disconnect-Fly -ErrorAction SilentlyContinue
    if ($cleanedCsvPath -and (Test-Path $cleanedCsvPath)) {
        Remove-Item $cleanedCsvPath -Force -ErrorAction SilentlyContinue
    }
}

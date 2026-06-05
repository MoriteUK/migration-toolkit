#Requires -Version 7.0
<#
.SYNOPSIS
    Gets migration data for dashboard display
.DESCRIPTION
    Retrieves real migration status data from AvePoint Fly for a specific project prefix
.PARAMETER ProjectPrefix
    The project prefix to query (e.g., "Contoso")
#>
param(
    [Parameter(Mandatory=$true)]
    [string]$ProjectPrefix
)

# Load library
. "$PSScriptRoot\lib.ps1"

# Get Fly API configuration
$flyApiCfgPath = Join-Path $env:APPDATA "FlyMigration\config.json"
if (-not (Test-Path $flyApiCfgPath)) {
    Write-Error "Fly API configuration not found. Please configure in Settings."
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
        Write-Error "Fly.Client module not found. Please install it first."
        exit 1
    }
    Import-Module Fly.Client -ErrorAction Stop
} catch {
    Write-Error "Failed to import Fly.Client module: $_"
    exit 1
}

# Connect to Fly API
try {
    Connect-Fly -Url $apiUrl -ClientId $clientId -ClientSecret $clientSecret -ErrorAction Stop
} catch {
    Write-Error "Failed to connect to Fly API: $_"
    exit 1
}

# Initialize result object
$result = @{
    TotalItems = 0
    Completed = 0
    InProgress = 0
    NotStarted = 0
    Failed = 0
    Warnings = 0
    Workloads = @{}
    FailedItems = @()
    WarningItems = @()
    InProgressItems = @()
    CompletedItems = @()
}

# Workload definitions
$workloads = @{
    SharePoint = 'SharePoint'
    Exchange = 'Exchange'
    OneDrive = 'OneDrive'
    Teams = 'Teams'
    'Teams Chat' = 'Teams Chat'
}

# Query each workload
foreach ($wlKey in $workloads.Keys) {
    $wlName = $workloads[$wlKey]
    $projectName = "$ProjectPrefix - $wlName"

    try {
        # Check if project exists
        $project = Get-FlyMigrationProject -Name $projectName -ErrorAction SilentlyContinue

        if (-not $project) {
            Write-Verbose "Project not found: $projectName"
            continue
        }

        # Get status data
        $tempFile = [System.IO.Path]::GetTempFileName()
        $statusCmd = $script:FlyWorkloadDefs[$wlKey].Status

        if ($statusCmd) {
            & $statusCmd -Project $projectName -OutFile $tempFile -ErrorAction SilentlyContinue | Out-Null

            if (Test-Path $tempFile) {
                $rows = Import-Csv $tempFile -ErrorAction SilentlyContinue
                Remove-Item $tempFile -Force -ErrorAction SilentlyContinue

                if ($rows) {
                    $wlTotal = $rows.Count
                    $wlCompleted = 0
                    $wlInProgress = 0
                    $wlNotStarted = 0
                    $wlFailed = 0
                    $wlWarnings = 0

                    # Find status column
                    $statusCol = $null
                    foreach ($colName in @('Stage status', 'StageStatus', 'Stage Status', 'Status', 'MigrationStatus', 'Migration Status', 'State')) {
                        if ($rows[0].PSObject.Properties.Name -contains $colName) {
                            $statusCol = $colName
                            break
                        }
                    }

                    # Count statuses
                    foreach ($row in $rows) {
                        $status = if ($statusCol) { ([string]$row.$statusCol).Trim() } else { '' }
                        $sourceUser = if ($row.PSObject.Properties['SourceUserPrincipalName']) { $row.SourceUserPrincipalName } `
                                      elseif ($row.PSObject.Properties['Source']) { $row.Source } else { 'Unknown' }

                        switch -Regex ($status) {
                            '^(Finished|Complete|Completed|Successful|Success)$' {
                                $wlCompleted++
                                # Add to recent completed (limit to last 5)
                                if ($result.CompletedItems.Count -lt 5) {
                                    $result.CompletedItems += @{
                                        Name = $sourceUser
                                        Workload = $wlName
                                        Status = $status
                                    }
                                }
                            }
                            '^(Exceptions|Exceptioned|CompletedWithException|FinishedWithException)$' {
                                $wlWarnings++
                                $wlCompleted++
                                $result.WarningItems += @{
                                    Name = $sourceUser
                                    Workload = $wlName
                                    Warning = "Completed with exceptions"
                                    Status = $status
                                }
                            }
                            '^(Failed)$' {
                                $wlFailed++
                                $errorMsg = if ($row.PSObject.Properties['ErrorMessage']) { $row.ErrorMessage } `
                                           elseif ($row.PSObject.Properties['Error']) { $row.Error } else { 'Migration failed' }
                                $result.FailedItems += @{
                                    Name = $sourceUser
                                    Workload = $wlName
                                    Error = $errorMsg
                                }
                            }
                            '^(In progress|In queue|In queue with priority|Scheduled|InProgress|Waiting|Queued)$' {
                                $wlInProgress++
                                if ($result.InProgressItems.Count -lt 10) {
                                    $result.InProgressItems += @{
                                        Name = $sourceUser
                                        Workload = $wlName
                                        Status = $status
                                    }
                                }
                            }
                            default {
                                $wlNotStarted++
                            }
                        }
                    }

                    # Update totals
                    $result.TotalItems += $wlTotal
                    $result.Completed += $wlCompleted
                    $result.InProgress += $wlInProgress
                    $result.NotStarted += $wlNotStarted
                    $result.Failed += $wlFailed
                    $result.Warnings += $wlWarnings

                    # Store workload breakdown
                    $result.Workloads[$wlName] = @{
                        Total = $wlTotal
                        Completed = $wlCompleted
                        InProgress = $wlInProgress
                        NotStarted = $wlNotStarted
                        Failed = $wlFailed
                        Warnings = $wlWarnings
                    }
                }
            }
        }
    } catch {
        Write-Verbose "Error querying $wlName : $_"
    }
}

# Output as JSON for the web app to consume
$result | ConvertTo-Json -Depth 10 -Compress

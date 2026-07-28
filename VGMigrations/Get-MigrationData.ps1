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

. "$PSScriptRoot\lib.ps1"

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

try {
    Connect-Fly -Url $apiUrl -ClientId $clientId -ClientSecret $clientSecret -ErrorAction Stop
} catch {
    Write-Error "Failed to connect to Fly API: $_"
    exit 1
}

# Get bearer token directly for raw REST calls (Get-FlyConfiguration is private to the module)
$restBaseUrl = $apiUrl.TrimEnd('/')
$restHeaders = $null
$identityHost = if ($apiUrl -match '\-gov\.') { 'identity-gov.avepointonlineservices.com' } else { 'identity.avepointonlineservices.com' }
try {
    $tokenBody = "grant_type=client_credentials&scope=fly.graph.readwrite.all" +
                 "&client_id=$([uri]::EscapeDataString($clientId))" +
                 "&client_secret=$([uri]::EscapeDataString($clientSecret))"
    $tokenResp   = Invoke-RestMethod -Uri "https://$identityHost/connect/token" `
        -Method POST -ContentType 'application/x-www-form-urlencoded' -Body $tokenBody -ErrorAction Stop
    $restHeaders = @{ Authorization = "Bearer $($tokenResp.access_token)"; Accept = 'application/json' }
} catch {
    Write-Warning "Could not obtain access token for direct REST calls: $_"
}

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

foreach ($wlKey in $script:FlyWorkloadDefs.Keys) {
    $result.Workloads[$wlKey] = @{
        Total = 0; Completed = 0; InProgress = 0; NotStarted = 0; Failed = 0; Warnings = 0
        ProjectFound = $false
    }
}

foreach ($wlKey in $script:FlyWorkloadDefs.Keys) {
    $projectName = "$ProjectPrefix - $wlKey"

    try {
        $project = Get-FlyMigrationProject -Name $projectName -ErrorAction SilentlyContinue

        if (-not $project -or -not $project.Id) {
            Write-Verbose "Project not found: $projectName"
            continue
        }

        # Use the counts already on the ProjectSummaryModel — no need to page through all mappings
        $wlTotal     = [int]($project.mappingTotalCount              ?? 0)
        $wlFailed    = [int]($project.mappingFailedCount             ?? 0)
        $wlWarnings  = [int]($project.mappingCompletedWithExceptionCount ?? 0)
        $wlCompleted = [int](($project.mappingCompletedCount         ?? 0) + $wlWarnings)
        $wlNotStart  = [int]($project.mappingNotMigratedCount        ?? 0)
        $wlInProg    = [int](($project.mappingWaitingCount           ?? 0) +
                             ($project.mappingInProgressCount        ?? 0) +
                             ($project.mappingStoppedCount           ?? 0) +
                             ($project.mappingScheduledCount         ?? 0))

        $result.TotalItems += $wlTotal
        $result.Completed  += $wlCompleted
        $result.InProgress += $wlInProg
        $result.NotStarted += $wlNotStart
        $result.Failed     += $wlFailed
        $result.Warnings   += $wlWarnings

        $result.Workloads[$wlKey] = @{
            Total = $wlTotal; Completed = $wlCompleted; InProgress = $wlInProg
            NotStarted = $wlNotStart; Failed = $wlFailed; Warnings = $wlWarnings
            ProjectFound = $true; ProjectId = "$($project.Id)"
        }

        # Fetch all mappings for this project and sort client-side.
        # Server-side stageStatuses filtering is unreliable (parameter encoding varies), so we
        # fetch a broad page and categorise in PowerShell.
        if ($wlFailed -gt 0 -or $wlWarnings -gt 0 -or ($wlInProg -gt 0 -and $result.InProgressItems.Count -lt 10)) {
            try {
                if (-not $restHeaders) { throw 'No access token available for REST calls' }
                $uri = "$restBaseUrl/projects/$($project.Id)/mappings/summaries?top=500"
                $page = Invoke-RestMethod -Uri $uri -Method GET -Headers $restHeaders -ErrorAction Stop
                foreach ($mapping in ($page.data.data ?? @())) {
                    # Handle stageStatus as integer OR string enum (API may return either)
                    $stageVal = $mapping.stageStatus
                    $statusCode = switch -Regex ("$stageVal") {
                        '^6$|^Failed$'      { 6 }
                        '^5$|^Exceptioned$' { 5 }
                        '^3$|^InProgress$'  { 3 }
                        '^1$|^Waiting$'     { 1 }
                        '^2$|^Queued$'      { 2 }
                        '^7$|^Stopped$'     { 7 }
                        '^4$|^Successful$'  { 4 }
                        default             { 0 }
                    }
                    $sourceUser  = ($mapping.sourceName ?? $mapping.sourceIdentity ?? $mapping.identity) ?? 'Unknown'
                    $destUser    = ($mapping.destinationName ?? $mapping.destinationIdentity) ?? ''
                    $lastRunTime = if ($mapping.lastMigrationStartTime -gt 0) {
                        try { Convert-TicksToDateTime -Ticks $mapping.lastMigrationStartTime -ShowHourFormat $true } catch { '' }
                    } else { '' }

                    switch ($statusCode) {
                        6 {  # Failed
                            $errCount = [int]($mapping.errorItemCount ?? 0)
                            $result.FailedItems += @{
                                Name = $sourceUser; Destination = $destUser
                                Workload = $wlKey; Project = $projectName
                                Status = 'Failed'; ErrorCount = $errCount; LastRunTime = $lastRunTime
                                ProjectId = "$($project.Id)"; MappingId = "$($mapping.id)"
                            }
                        }
                        5 {  # Exceptioned / completed with warnings
                            $errCount = [int]($mapping.errorItemCount ?? 0)
                            $result.WarningItems += @{
                                Name = $sourceUser; Destination = $destUser
                                Workload = $wlKey; Project = $projectName
                                Warning = 'Completed with exceptions'; Status = 'Exceptions'
                                ErrorCount = $errCount; LastRunTime = $lastRunTime
                                ProjectId = "$($project.Id)"; MappingId = "$($mapping.id)"
                            }
                        }
                        { $_ -in @(1, 2, 3, 7) -and $result.InProgressItems.Count -lt 10 } {
                            $status = switch ($_) { 1 { 'In queue' } 2 { 'In queue' } 3 { 'In progress' } 7 { 'Stopped' } }
                            $result.InProgressItems += @{
                                Name = $sourceUser; Destination = $destUser
                                Workload = $wlKey; Project = $projectName; Status = $status; LastRunTime = $lastRunTime
                            }
                        }
                    }
                }
            } catch {
                Write-Warning "Error fetching mappings for $wlKey : $_"
            }
        }

    } catch {
        Write-Warning "Error querying $wlKey : $_"
    }
}

$result | ConvertTo-Json -Depth 10 -Compress

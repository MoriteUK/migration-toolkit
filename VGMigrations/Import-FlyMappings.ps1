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
    Write-Host "  API URL: $apiUrl" -ForegroundColor Gray
    # Test whether the API hostname resolves before attempting the connection
    try {
        $apiHost = ([System.Uri]$apiUrl).Host
        $addrs   = [System.Net.Dns]::GetHostAddresses($apiHost)
        Write-Host "  Hostname '$apiHost' resolves to: $($addrs.IPAddressToString -join ', ')" -ForegroundColor Gray
    } catch {
        Write-Warning "  DNS lookup failed for API hostname: $_"
        Write-Warning "  The Fly API may be unreachable from this machine (firewall / proxy / VPN required)."
    }
    Connect-Fly -Url $apiUrl -ClientId $clientId -ClientSecret $clientSecret -ErrorAction Stop
    Write-Host "Connected successfully" -ForegroundColor Green
} catch {
    Write-Error "Failed to connect to Fly API: $_"
    exit 1
}

Write-Host "Importing into project '$ProjectName'..." -ForegroundColor Cyan

# Detect CSV encoding from BOM so non-ASCII characters (Hebrew, Spanish, etc.)
# are read correctly regardless of how the file was saved.
function Get-CsvEncoding([string]$Path) {
    # Open with FileShare.ReadWrite so OneDrive/Excel locks don't block us
    try {
        $fs    = [System.IO.FileStream]::new($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $bytes = New-Object byte[] 4
        $read  = $fs.Read($bytes, 0, 4)
        $fs.Close()
        if ($read -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { return 'UTF8'             }
        if ($read -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) { return 'Unicode'          }  # UTF-16 LE
        if ($read -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) { return 'BigEndianUnicode' }  # UTF-16 BE
    } catch {
        Write-Warning "Could not detect CSV encoding ($($_.Exception.Message)) -- defaulting to system ANSI."
    }
    return 'Default'  # ANSI / Windows system code page (e.g. Windows-1252)
}

# Pre-validate and clean the CSV -- remove rows where source == destination
$cleanedCsvPath = $null
try {
    $csvEncoding = Get-CsvEncoding $MappingFile
    Write-Host "CSV encoding detected: $csvEncoding" -ForegroundColor Gray
    $allRows = Import-Csv $MappingFile -Encoding $csvEncoding -ErrorAction Stop

    if ($allRows.Count -gt 0) {
        # Detect source/destination column names dynamically -- different workloads use
        # different headers (e.g. "Source" vs "Source user" vs "Source site")
        $headers = $allRows[0].PSObject.Properties.Name
        # Prefer identity columns (email/URL/user) over plain name columns so that
        # groups/teams with the same display name but different email aren't skipped.
        $srcCol  = $headers | Where-Object { $_ -imatch '^source'      -and $_ -imatch 'email|url|user' } | Select-Object -First 1
        if (-not $srcCol) { $srcCol = $headers | Where-Object { $_ -imatch '^source' }      | Select-Object -First 1 }
        $dstCol  = $headers | Where-Object { $_ -imatch '^destination' -and $_ -imatch 'email|url|user' } | Select-Object -First 1
        if (-not $dstCol) { $dstCol = $headers | Where-Object { $_ -imatch '^destination' } | Select-Object -First 1 }

        Write-Host "CSV columns: $($headers -join ', ')" -ForegroundColor Gray

        if ($srcCol -and $dstCol) {
            Write-Host "Using '$srcCol'  '$dstCol' as source/destination columns." -ForegroundColor Gray
            # Pass 1 -- drop rows where source == destination.
            # Use OrdinalIgnoreCase so accented/Unicode characters compare correctly.
            $cleanRows = $allRows | Where-Object {
                $_.$srcCol -and $_.$dstCol -and
                -not [string]::Equals($_.$srcCol.Trim(), $_.$dstCol.Trim(), [System.StringComparison]::OrdinalIgnoreCase)
            }
            $sameCount = $allRows.Count - $cleanRows.Count
            if ($sameCount -gt 0) {
                Write-Warning "$sameCount row(s) skipped -- source and destination are identical:"
                $allRows | Where-Object {
                    $_.$srcCol -and $_.$dstCol -and
                    [string]::Equals($_.$srcCol.Trim(), $_.$dstCol.Trim(), [System.StringComparison]::OrdinalIgnoreCase)
                } | ForEach-Object { Write-Host "  SKIP (same): $($_.$srcCol)" -ForegroundColor Yellow }
            }

            # Pass 2 -- drop rows where the same source appears more than once
            # (Fly rejects the entire batch if any source is duplicated)
            $seen      = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $dedupRows = [System.Collections.Generic.List[psobject]]::new()
            $dupCount  = 0
            foreach ($row in $cleanRows) {
                $key = $row.$srcCol.Trim()
                if (-not $seen.Add($key)) {
                    Write-Host "  SKIP (dup src): $($row.$srcCol) -> $($row.$dstCol)" -ForegroundColor Yellow
                    $dupCount++
                } else {
                    $dedupRows.Add($row)
                }
            }
            if ($dupCount -gt 0) {
                Write-Warning "$dupCount row(s) skipped -- duplicate source entries (Fly rejects the whole batch if any source appears twice)."
            }

            $finalRows = $dedupRows.ToArray()
            if ($finalRows.Count -eq 0) {
                Write-Error "No valid rows to import after removing same-identity and duplicate-source rows."
                exit 1
            }

            $totalSkipped = $allRows.Count - $finalRows.Count
            if ($totalSkipped -gt 0) {
                Write-Host "$($finalRows.Count) of $($allRows.Count) rows will be imported." -ForegroundColor Cyan
            }

            # Always write a normalised UTF-8 BOM copy -- the Fly.Client .NET module
            # reads the file as UTF-8 regardless of the original encoding, so passing
            # an ANSI file directly causes a silent parse failure even for ASCII content.
            $cleanedCsvPath = [System.IO.Path]::GetTempFileName() -replace '\.tmp$', '.csv'
            $finalRows | Export-Csv $cleanedCsvPath -NoTypeInformation -Encoding utf8BOM -UseQuotes AsNeeded
            Write-Host "Normalised to UTF-8: $cleanedCsvPath" -ForegroundColor Gray
            $MappingFile = $cleanedCsvPath
        } else {
            Write-Warning "Could not detect source/destination columns in CSV -- skipping pre-validation."
        }
    }
} catch {
    Write-Error "Failed to validate mapping file: $_"
    exit 1
}

# Import mappings by calling Fly.Client private functions via module scope.
# The public Import-Fly*Mappings cmdlets all call Get-FlyProjectByName internally,
# which searches the Fly API with the FULL project name ("Pcentra - Exchange"). The Fly
# API returns no results when the search term contains " - ", causing the lookup to throw
# "Failed to retrieve the project." We fix this by calling Get-FlyProjects ourselves with
# just the customer prefix, then calling Add-Fly*Mappings directly once we have the project ID.
# We use & $flyModule { param(...) ... } to invoke private module functions in module scope,
# which keeps all proxy/TLS/header handling inside the module's own Invoke-FlyApiClient.
try {
    Write-Host "`nImporting mappings..." -ForegroundColor Cyan

    $flyModule = Get-Module Fly.Client

    # Search with just the customer prefix so the Fly API search endpoint doesn't choke
    # on the " - " separator, then filter the result by exact project name.
    $searchPrefix = ($ProjectName -split ' - ')[0].Trim()
    Write-Host "Locating project '$ProjectName' (search prefix: '$searchPrefix')..." -ForegroundColor Cyan

    $targetProject = $null
    $skip = 0
    try {
        do {
            $page = & $flyModule { param($pfx, $top, $sk)
                Get-FlyProjects -Search $pfx -Top $top -Skip $sk
            } $searchPrefix 200 $skip

            $targetProject = $page.data | Where-Object { $_.name -ieq $ProjectName } | Select-Object -First 1
            $skip += 200
        } while (-not $targetProject -and $page.nextLink)
    } catch {
        Write-Error "Fly API search failed: $_"
        Write-Host "  Error type : $($_.Exception.GetType().FullName)" -ForegroundColor Yellow
        Write-Host "  Inner error: $($_.Exception.InnerException?.Message)" -ForegroundColor Yellow
        Write-Host "  Stack      : $($_.ScriptStackTrace)" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "This usually means the Fly API host ('$apiUrl') is not reachable from" -ForegroundColor Yellow
        Write-Host "this PowerShell process. Check firewall rules, VPN, or proxy settings." -ForegroundColor Yellow
        exit 1
    }

    if (-not $targetProject) {
        Write-Error "Project '$ProjectName' not found in Fly. Use 'Create Project' first."
        exit 1
    }
    Write-Host "Project found (ID: $($targetProject.id))" -ForegroundColor Green

    # Data type integer values -- replicates Get-FlyDataType from the Fly.Client module
    $dataTypeMap = @{
        'Site collection'             = 600
        'Site'                        = 400
        'List'                        = 200
        'Folder'                      = 100
        'User mailbox'                = 1001
        'Archive mailbox'             = 1002
        'Distribution list'           = 1007
        'Microsoft 365 Group mailbox' = 1003
        'Resource mailbox'            = 1004
        'Shared mailbox'              = 1005
        'Mail-enabled security group' = 1008
        'Microsoft 365 Group'         = 1006
    }

    $csvRows     = Import-Csv $MappingFile -Encoding UTF8 -ErrorAction Stop
    $mappingBody = [System.Collections.Generic.List[psobject]]::new()
    $projectId   = $targetProject.id

    # Detect the actual CSV column layout so we can handle both Teams mapping formats:
    # - Team-level: "Source Team email address" / "Destination Team email address"
    # - User-level: "Source user" / "Destination user" (also valid for Teams projects)
    $csvHeaders = if ($csvRows.Count -gt 0) { $csvRows[0].PSObject.Properties.Name } else { @() }

    # For workloads other than Teams, validate required columns up front
    $requiredCols = @{
        Exchange   = @('Source', 'Source type', 'Destination', 'Destination type')
        SharePoint = @('Source URL', 'Source object level', 'Destination URL', 'Destination object level')
        OneDrive   = @('Source user', 'Destination user')
        TeamChat   = @('Source user', 'Destination user')
        Groups     = @('Source group name', 'Source group email address', 'Destination group name', 'Destination group email address')
    }
    if ($requiredCols.ContainsKey($Workload)) {
        $missing = $requiredCols[$Workload] | Where-Object { $_ -notin $csvHeaders }
        if ($missing) {
            Write-Error "CSV columns do not match the '$Workload' workload format."
            Write-Host "  Expected : $($requiredCols[$Workload] -join ', ')" -ForegroundColor Yellow
            Write-Host "  Found    : $($csvHeaders -join ', ')" -ForegroundColor Yellow
            Write-Host "  Missing  : $($missing -join ', ')" -ForegroundColor Yellow
            exit 1
        }
    }

    # For Teams, detect which column layout is in use
    $teamsUserLevel = $Workload -eq 'Teams' -and ('Source user' -in $csvHeaders)

    switch ($Workload) {
        'Exchange' {
            foreach ($row in $csvRows) {
                $srcType = $dataTypeMap[$row.'Source type']
                $dstType = $dataTypeMap[$row.'Destination type']
                if ($null -eq $srcType -or $null -eq $dstType) {
                    Write-Warning "Skipping row -- unrecognised type: '$($row.'Source type')' / '$($row.'Destination type')'"
                    continue
                }
                $mappingBody.Add([PSCustomObject]@{
                    sourceIdentity      = $row.Source
                    sourceType          = $srcType
                    destinationIdentity = $row.Destination
                    destinationType     = $dstType
                })
            }
        }
        'SharePoint' {
            foreach ($row in $csvRows) {
                $srcType = $dataTypeMap[$row.'Source object level']
                $dstType = $dataTypeMap[$row.'Destination object level']
                $method  = if ($row.Method -eq 'Merge') { 1 } else { 0 }
                $mappingBody.Add([PSCustomObject]@{
                    sourceIdentity      = $row.'Source URL'
                    sourceType          = $srcType
                    destinationIdentity = $row.'Destination URL'
                    destinationType     = $dstType
                    method              = $method
                })
            }
        }
        'OneDrive' {
            foreach ($row in $csvRows) {
                $mappingBody.Add([PSCustomObject]@{
                    sourceIdentity      = $row.'Source user'
                    destinationIdentity = $row.'Destination user'
                })
            }
        }
        'Teams' {
            if ($teamsUserLevel) {
                Write-Host "Using user-level Teams mapping format (Source user / Destination user)." -ForegroundColor Gray
            }
            foreach ($row in $csvRows) {
                if ($teamsUserLevel) {
                    # User-level Teams mapping: map by user email address
                    $mappingBody.Add([PSCustomObject]@{
                        enableChannelMapping = $false
                        includeOtherChannels = $false
                        sourceName           = $row.'Source user'
                        sourceIdentity       = $row.'Source user'
                        destinationName      = $row.'Destination user'
                        destinationIdentity  = $row.'Destination user'
                        channelMapping       = @()
                    })
                    continue
                }
                # Team-level mapping: map by team name and email address
                $item = [PSCustomObject]@{
                    enableChannelMapping = $false
                    includeOtherChannels = $false
                    sourceName           = $row.'Source Team name'
                    sourceIdentity       = $row.'Source Team email address'
                    destinationName      = $row.'Destination Team name'
                    destinationIdentity  = $row.'Destination Team email address'
                    channelMapping       = @()
                }
                if ($row.'Channel mappings' -eq 'Enabled') {
                    $item.enableChannelMapping = $true
                    $item.channelMapping = @([PSCustomObject]@{
                        sourceName      = $row.'Source channel name'
                        destinationType = if ($row.'Destination channel type' -eq 'Standard') { 0 } else { 1 }
                        destinationName = $row.'Destination channel name'
                    })
                }
                $mappingBody.Add($item)
            }
        }
        'TeamChat' {
            foreach ($row in $csvRows) {
                $mappingBody.Add([PSCustomObject]@{
                    sourceIdentity      = $row.'Source user'
                    destinationIdentity = $row.'Destination user'
                })
            }
        }
        'Groups' {
            foreach ($row in $csvRows) {
                $mappingBody.Add([PSCustomObject]@{
                    sourceName          = $row.'Source Group name'
                    sourceIdentity      = $row.'Source Group email address'
                    destinationName     = $row.'Destination Group name'
                    destinationIdentity = $row.'Destination Group email address'
                })
            }
        }
    }

    if ($mappingBody.Count -eq 0) {
        Write-Error "No valid mappings to import after processing CSV rows."
        exit 1
    }

    Write-Host "Posting $($mappingBody.Count) mapping(s)..." -ForegroundColor Cyan
    [System.Net.HttpWebRequest]::DefaultMaximumErrorResponseLength = -1

    $mappingArr = [object[]]$mappingBody.ToArray()

    try {
        switch ($Workload) {
            'Exchange'   { & $flyModule { param($id, $body) Add-FlyExchangeMappings   -ProjectId $id -ExchangeMappingCreationModel  $body } $projectId $mappingArr }
            'SharePoint' { & $flyModule { param($id, $body) Add-FlySharePointMappings -ProjectId $id -SharePointMappingCreationModel $body } $projectId $mappingArr }
            'OneDrive'   { & $flyModule { param($id, $body) Add-FlyOneDriveMappings   -ProjectId $id -OneDriveMappingCreationModel   $body } $projectId $mappingArr }
            'Teams'      { & $flyModule { param($id, $body) Add-FlyTeamsMappings      -ProjectId $id -TeamsMappingCreationModel      $body } $projectId $mappingArr }
            'TeamChat'   { & $flyModule { param($id, $body) Add-FlyTeamChatMappings   -ProjectId $id -TeamChatMappingCreationModel   $body } $projectId $mappingArr }
            'Groups'     { & $flyModule { param($id, $body) Add-FlyM365GroupMappings  -ProjectId $id -M365GroupMappingCreationModel  $body } $projectId $mappingArr }
        }
        Write-Host "`n✓ Mappings imported successfully!" -ForegroundColor Green
    } catch {
        $errText = "$_"
        $errBody = $_.ErrorDetails.Message
        if ($errBody -match 'ProjectMappingDuplicated') {
            Write-Warning "All mappings already exist in project '$ProjectName' -- nothing new to import."
        } else {
            Write-Error "Import failed: $errText"
            if ($errBody) { Write-Host "API error detail: $errBody" -ForegroundColor Yellow }
            exit 1
        }
    }

    Write-Host "`n=== Next Steps ===" -ForegroundColor Cyan
    Write-Host "1. Review mappings in AvePoint Fly portal"
    Write-Host "2. Run pre-scan for this workload via the Connections view"
    Write-Host "3. Review scan results, then start migration"

} catch {
    Write-Error "Failed to import mappings: $_"
    exit 1
} finally {
    Disconnect-Fly -ErrorAction SilentlyContinue
    if ($cleanedCsvPath -and (Test-Path $cleanedCsvPath)) {
        Remove-Item $cleanedCsvPath -Force -ErrorAction SilentlyContinue
    }
}

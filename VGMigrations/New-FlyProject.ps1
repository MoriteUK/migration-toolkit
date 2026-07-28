#Requires -Version 7.0
<#
.SYNOPSIS
    Creates a new AvePoint Fly migration project
.DESCRIPTION
    Looks up the destination connection in Fly by matching the customer prefix
    against existing connection names, then creates the project.
.PARAMETER ProjectName
    Project name in the format "{CustomerPrefix} - {Workload}" (e.g., "Fara - Exchange")
.PARAMETER Workload
    Workload type: SharePoint, Exchange, OneDrive, Teams, TeamChat, Groups
.PARAMETER SourceConnection
    Source connection name (e.g., "OurVolaris - EXO")
.PARAMETER CustomerDomain
    Customer domain prefix (not used for matching; kept for compatibility)
.PARAMETER Policy
    Policy name to assign to the project
#>
param(
    [Parameter(Mandatory=$true)]  [string]$ProjectName,
    [Parameter(Mandatory=$true)]  [ValidateSet('SharePoint','Exchange','OneDrive','Teams','TeamChat','Groups')] [string]$Workload,
    [Parameter(Mandatory=$true)]  [string]$SourceConnection,
    [Parameter(Mandatory=$false)] [string]$CustomerDomain = "",
    [Parameter(Mandatory=$false)] [string]$Description = ""
)

. "$PSScriptRoot\lib.ps1"

Write-Host "`n=== $ProjectName ===" -ForegroundColor Cyan

# Load Fly API config
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
    } else { Write-Error "Client secret not found in config."; exit 1 }
} catch { Write-Error "Failed to load config: $_"; exit 1 }

# Import module
try {
    if (-not (Get-Module -Name Fly.Client -ListAvailable)) { Write-Error "Fly.Client module not found."; exit 1 }
    Import-Module Fly.Client -ErrorAction Stop
} catch { Write-Error "Failed to import Fly.Client: $_"; exit 1 }

# Connect
try {
    Write-Host "Connecting to Fly API..." -ForegroundColor Cyan
    Connect-Fly -Url $apiUrl -ClientId $clientId -ClientSecret $clientSecret -ErrorAction Stop
    Write-Host "Connected" -ForegroundColor Green
} catch { Write-Error "Failed to connect: $_"; exit 1 }

try {
    # Look up destination connection (auto-creates via fly-connector.js if missing)
    $DestinationConnection = Find-FlyDestinationConnection `
        -ProjectName    $ProjectName `
        -Workload       $Workload `
        -ApiUrl         $apiUrl `
        -CustomerDomain $CustomerDomain

    # Find default policy for this workload
    $Policy = Find-FlyPolicy -Workload $Workload -ApiUrl $apiUrl

    # Skip if project already exists
    $existing = Get-FlyMigrationProject -Name $ProjectName -ErrorAction SilentlyContinue
    if ($existing -and $existing.Id) {
        Write-Host "Project '$ProjectName' already exists (ID: $($existing.Id)) — skipping." -ForegroundColor Yellow
        exit 0
    }

    # Create project
    Write-Host "Creating project '$ProjectName'..." -ForegroundColor Cyan
    $params = @{
        Name                  = $ProjectName
        SourceConnection      = $SourceConnection
        DestinationConnection = $DestinationConnection
        Policy                = $Policy
    }
    if ($Description) { $params.Description = $Description }
    $newProject = New-FlyMigrationProject @params -ErrorAction Stop

    Write-Host "✓ Project created (ID: $($newProject.Id))" -ForegroundColor Green

} catch { Write-Error "Failed: $_"; exit 1 }
finally { Disconnect-Fly -ErrorAction SilentlyContinue }

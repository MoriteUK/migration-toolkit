#Requires -Version 7.0
<#
.SYNOPSIS
    Creates a new AvePoint Fly migration project
.DESCRIPTION
    Creates a migration project for a specific workload using the Fly.Client module
.PARAMETER ProjectName
    Name of the project (e.g., "Contoso - SharePoint")
.PARAMETER Workload
    Workload type: SharePoint, Exchange, OneDrive, Teams, TeamChat, Groups
.PARAMETER Description
    Optional project description
.EXAMPLE
    .\New-FlyProject.ps1 -ProjectName "Contoso - SharePoint" -Workload SharePoint
#>
param(
    [Parameter(Mandatory=$true)]
    [string]$ProjectName,

    [Parameter(Mandatory=$true)]
    [ValidateSet('SharePoint', 'Exchange', 'OneDrive', 'Teams', 'TeamChat', 'Groups')]
    [string]$Workload,

    [Parameter(Mandatory=$false)]
    [string]$Description = ""
)

. "$PSScriptRoot\lib.ps1"

Write-Host "`nCreating Fly Migration Project..." -ForegroundColor Cyan
Write-Host "Project Name: $ProjectName" -ForegroundColor White
Write-Host "Workload: $Workload" -ForegroundColor White

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

# Check if project already exists
try {
    $existingProject = Get-FlyMigrationProject -Name $ProjectName -ErrorAction SilentlyContinue
    if ($existingProject) {
        Write-Warning "Project '$ProjectName' already exists!"
        Write-Host "Project ID: $($existingProject.Id)" -ForegroundColor Yellow
        exit 0
    }
} catch {
    # Project doesn't exist, which is what we want
}

# Create the project
try {
    Write-Host "`nCreating project..." -ForegroundColor Cyan

    $projectParams = @{
        Name = $ProjectName
    }

    if ($Description) {
        $projectParams.Description = $Description
    }

    $newProject = New-FlyMigrationProject @projectParams -ErrorAction Stop

    Write-Host "`n✓ Project created successfully!" -ForegroundColor Green
    Write-Host "Project ID: $($newProject.Id)" -ForegroundColor White
    Write-Host "Name: $($newProject.Name)" -ForegroundColor White

    if ($newProject.Description) {
        Write-Host "Description: $($newProject.Description)" -ForegroundColor White
    }

    # Create default policy for the workload
    Write-Host "`nCreating default policy for $Workload..." -ForegroundColor Cyan

    $policyCmd = $script:FlyWorkloadDefs[$Workload].PolicyCmd
    if ($policyCmd) {
        try {
            $policyName = "$ProjectName - Default Policy"
            & "New-Fly$($Workload)Policy" -Name $policyName -Project $ProjectName -ErrorAction Stop
            Write-Host "✓ Policy created: $policyName" -ForegroundColor Green
        } catch {
            Write-Warning "Failed to create policy: $_"
        }
    }

    Write-Host "`n=== Next Steps ===" -ForegroundColor Cyan
    Write-Host "1. Import mappings using: Import-Fly$($Workload)Mappings"
    Write-Host "2. Configure policy settings in AvePoint Fly portal"
    Write-Host "3. Start pre-scan: Start-Fly$($Workload)PreScan"
    Write-Host "4. Start migration: Start-Fly$($Workload)Migration"

} catch {
    Write-Error "Failed to create project: $_"
    exit 1
} finally {
    Disconnect-Fly -ErrorAction SilentlyContinue
}

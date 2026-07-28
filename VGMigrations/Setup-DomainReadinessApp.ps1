#Requires -Version 7.0
<#
.SYNOPSIS
    One-time setup: Creates app registration for domain readiness checks
.DESCRIPTION
    Run this script ONCE per target tenant to create the app registration.
    After setup completes, Check-DomainMigrationReadiness.ps1 will use the
    stored credentials automatically.
.PARAMETER TenantId
    The target tenant ID (e.g., intranotedk.onmicrosoft.com)
.PARAMETER TenantName
    Friendly name for this tenant (e.g., "Intranote")
.EXAMPLE
    .\Setup-DomainReadinessApp.ps1 -TenantId "intranotedk.onmicrosoft.com" -TenantName "Intranote"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$TenantId,

    [Parameter(Mandatory)]
    [string]$TenantName,

    [string]$AppName = "VG-DomainReadiness-Reporter",

    [string]$CredentialStorePath = $PSScriptRoot
)

$ErrorActionPreference = 'Stop'

function Write-Log {
    param([string]$Msg, [string]$Level = 'INFO')
    $colors = @{
        'INFO' = 'Cyan'
        'SUCCESS' = 'Green'
        'WARN' = 'Yellow'
        'ERROR' = 'Red'
    }
    Write-Host "[$Level] $Msg" -ForegroundColor $colors[$Level]
}

Write-Host ""
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "  Domain Readiness App Setup" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host ""
Write-Log "Tenant: $TenantName ($TenantId)"
Write-Log "App Name: $AppName"
Write-Host ""

# Check if already set up
$credFile = Join-Path $CredentialStorePath "domaincheck_$($TenantId).json"
if (Test-Path $credFile) {
    $existing = Get-Content $credFile | ConvertFrom-Json
    Write-Log "App registration already exists!" "WARN"
    Write-Log "  App ID: $($existing.AppId)" "WARN"
    Write-Log "  Created: $($existing.CreatedDate)" "WARN"
    Write-Host ""
    $continue = Read-Host "Create new credentials anyway? (y/N)"
    if ($continue -ne 'y') {
        Write-Log "Setup cancelled" "WARN"
        exit 0
    }
}

# Check Microsoft.Graph module
Write-Log "Checking Microsoft.Graph module..."
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Applications)) {
    Write-Log "Installing Microsoft.Graph.Applications..." "WARN"
    Install-Module -Name Microsoft.Graph.Applications -Scope CurrentUser -Force -AllowClobber
}

try {
    # Connect with interactive browser auth
    Write-Log "Opening browser for authentication..."
    Write-Log "Sign in as a Global Administrator or Application Administrator" "WARN"
    Write-Host ""

    Import-Module Microsoft.Graph.Authentication
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Connect-MgGraph -TenantId $TenantId -Scopes "Application.ReadWrite.All" -NoWelcome

    Write-Log "Connected successfully!" "SUCCESS"
    Write-Host ""

    # Check if app exists
    Write-Log "Checking for existing app registration..."
    $existingApp = Get-MgApplication -Filter "displayName eq '$AppName'" -ErrorAction SilentlyContinue

    if ($existingApp) {
        Write-Log "App exists: $($existingApp.AppId)" "SUCCESS"
        $appId = $existingApp.AppId
        $objectId = $existingApp.Id
    } else {
        Write-Log "Creating new app registration..." "WARN"

        # Required permissions
        $requiredResourceAccess = @(
            @{
                ResourceAppId = "00000003-0000-0000-c000-000000000000" # MS Graph
                ResourceAccess = @(
                    @{ Id = "7e05723c-0bb0-42da-be95-ae9f08a6e53c"; Type = "Role" } # Domain.ReadWrite.All
                )
            }
        )

        $newApp = New-MgApplication -DisplayName $AppName -SignInAudience "AzureADMyOrg" -RequiredResourceAccess $requiredResourceAccess
        $appId = $newApp.AppId
        $objectId = $newApp.Id
        Write-Log "App created: $appId" "SUCCESS"
        Start-Sleep -Seconds 3
    }

    # Create client secret
    Write-Log "Creating new client secret..."
    $passwordCred = @{
        DisplayName = "Auto-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
        EndDateTime = (Get-Date).AddYears(2)
    }
    $secret = Add-MgApplicationPassword -ApplicationId $objectId -PasswordCredential $passwordCred
    $clientSecret = $secret.SecretText
    Write-Log "Secret created (expires: $($secret.EndDateTime.ToString('yyyy-MM-dd')))" "SUCCESS"

    # Ensure service principal exists
    Write-Log "Checking service principal..."
    $sp = Get-MgServicePrincipal -Filter "appId eq '$appId'" -ErrorAction SilentlyContinue
    if (-not $sp) {
        Write-Log "Creating service principal..."
        $sp = New-MgServicePrincipal -AppId $appId
        Start-Sleep -Seconds 2
        Write-Log "Service principal created" "SUCCESS"
    } else {
        Write-Log "Service principal exists" "SUCCESS"
    }

    # Grant admin consent
    Write-Log "Granting admin consent for Domain.ReadWrite.All..."
    $graphSP = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
    try {
        New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -BodyParameter @{
            PrincipalId = $sp.Id
            ResourceId = $graphSP.Id
            AppRoleId = "7e05723c-0bb0-42da-be95-ae9f08a6e53c" # Domain.ReadWrite.All
        } -ErrorAction SilentlyContinue | Out-Null
        Write-Log "Admin consent granted" "SUCCESS"
    } catch {
        if ($_ -match "already exists") {
            Write-Log "Permission already granted" "SUCCESS"
        } else {
            throw
        }
    }

    # Store credentials
    Write-Log "Saving credentials..."
    $credObject = @{
        TenantId = $TenantId
        TenantName = $TenantName
        AppId = $appId
        ClientSecret = $clientSecret
        SecretExpires = $secret.EndDateTime
        CreatedDate = Get-Date
    }
    $credObject | ConvertTo-Json | Out-File -FilePath $credFile -Encoding UTF8 -Force
    Write-Log "Credentials saved to: $credFile" "SUCCESS"

    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

    Write-Host ""
    Write-Host "=========================================" -ForegroundColor Green
    Write-Host "  Setup Complete!" -ForegroundColor Green
    Write-Host "=========================================" -ForegroundColor Green
    Write-Host ""
    Write-Log "App ID: $appId" "SUCCESS"
    Write-Log "Secret expires: $($secret.EndDateTime.ToString('yyyy-MM-dd'))" "SUCCESS"
    Write-Host ""
    Write-Log "Check-DomainMigrationReadiness.ps1 will now use these credentials automatically" "SUCCESS"
    Write-Host ""

} catch {
    Write-Log "Setup failed: $_" "ERROR"
    exit 1
}

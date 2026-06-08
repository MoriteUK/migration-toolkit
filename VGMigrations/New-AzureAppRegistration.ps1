#Requires -Version 7.0
<#
.SYNOPSIS
    Creates an Azure AD App Registration for AvePoint Fly
.DESCRIPTION
    Creates an app registration with required Microsoft Graph and SharePoint permissions
.PARAMETER TenantId
    Target tenant ID where the app will be registered
.PARAMETER AppName
    Name for the application (default: "AvePoint Fly Migration")
.PARAMETER Interactive
    Use interactive browser authentication
.EXAMPLE
    .\New-AzureAppRegistration.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -Interactive
#>
param(
    [Parameter(Mandatory=$true)]
    [string]$TenantId,

    [Parameter(Mandatory=$false)]
    [string]$AppName = "AvePoint Fly Migration",

    [Parameter(Mandatory=$false)]
    [switch]$Interactive,

    [Parameter(Mandatory=$false)]
    [switch]$SkipSavePrompt
)

Write-Host "`n╔══════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║    Azure AD App Registration for AvePoint Fly            ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

Write-Host "`nTenant ID: $TenantId" -ForegroundColor White
Write-Host "App Name: $AppName" -ForegroundColor White

# Check if Microsoft.Graph module is installed
$graphModule = Get-Module -Name Microsoft.Graph.Applications -ListAvailable
if (-not $graphModule) {
    Write-Host "`n❌ Microsoft.Graph.Applications module not found" -ForegroundColor Red
    Write-Host "`nInstalling Microsoft Graph PowerShell SDK..." -ForegroundColor Yellow
    Write-Host "This may take a few minutes..." -ForegroundColor Gray

    try {
        Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
        Write-Host "✓ Microsoft Graph module installed" -ForegroundColor Green
    } catch {
        Write-Error "Failed to install Microsoft Graph module: $_"
        Write-Host "`nManual installation:" -ForegroundColor Yellow
        Write-Host "  Install-Module -Name Microsoft.Graph -Scope CurrentUser" -ForegroundColor Gray
        exit 1
    }
}

# Import required modules
try {
    Import-Module Microsoft.Graph.Applications -ErrorAction Stop
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Write-Host "✓ Microsoft Graph modules loaded" -ForegroundColor Green
} catch {
    Write-Error "Failed to import Microsoft Graph modules: $_"
    exit 1
}

# Connect to Microsoft Graph
Write-Host "`n📡 Connecting to Microsoft Graph..." -ForegroundColor Cyan
Write-Host "`n⚠️  DEVICE CODE AUTHENTICATION" -ForegroundColor Yellow
Write-Host "A code will appear below. Copy it and visit https://microsoft.com/devicelogin" -ForegroundColor Yellow
Write-Host "to complete authentication.`n" -ForegroundColor Yellow

try {
    # Always use device code authentication - it's more reliable when called from Electron
    Connect-MgGraph -TenantId $TenantId -Scopes "Application.ReadWrite.All" -UseDeviceAuthentication -ErrorAction Stop
    Write-Host "`n✓ Connected to tenant: $TenantId" -ForegroundColor Green
} catch {
    Write-Error "Failed to connect to Microsoft Graph: $_"
    Write-Host "`nTroubleshooting:" -ForegroundColor Yellow
    Write-Host "- Ensure you have 'Application Administrator' or 'Global Administrator' role" -ForegroundColor Gray
    Write-Host "- Check that the Tenant ID is correct" -ForegroundColor Gray
    Write-Host "- Make sure you completed the device code authentication" -ForegroundColor Gray
    exit 1
}

# Define required permissions
Write-Host "`n🔐 Configuring permissions..." -ForegroundColor Cyan

$requiredResourceAccess = @(
    # Microsoft Graph permissions
    @{
        ResourceAppId = "00000003-0000-0000-c000-000000000000" # Microsoft Graph
        ResourceAccess = @(
            @{ Id = "9e3f62cf-ca93-4989-b6ce-bf83c28f9fe8"; Type = "Role" }  # RoleManagement.ReadWrite.Directory
            @{ Id = "19dbc75e-c2e2-444c-a770-ec69d8559fc7"; Type = "Role" }  # Directory.ReadWrite.All
            @{ Id = "62a82d76-70ea-41e2-9197-370581804d09"; Type = "Role" }  # Group.ReadWrite.All
            @{ Id = "741f803b-c850-494e-b5df-cde7c675a1ca"; Type = "Role" }  # User.ReadWrite.All
            @{ Id = "1bfefb4e-e0b5-418b-a88f-73c46d2cc8e9"; Type = "Role" }  # Application.ReadWrite.All
            @{ Id = "06b708a9-e830-4db3-a914-8e69da51d44f"; Type = "Role" }  # AppRoleAssignment.ReadWrite.All
        )
    }
    # SharePoint permissions
    @{
        ResourceAppId = "00000003-0000-0ff1-ce00-000000000000" # SharePoint
        ResourceAccess = @(
            @{ Id = "678536fe-1083-478a-9c59-b99265e6b0d3"; Type = "Role" }  # Sites.FullControl.All
            @{ Id = "741f803b-c850-494e-b5df-cde7c675a1ca"; Type = "Role" }  # User.ReadWrite.All
        )
    }
)

# Create app registration
Write-Host "`n📝 Creating app registration..." -ForegroundColor Cyan

try {
    $app = New-MgApplication `
        -DisplayName $AppName `
        -SignInAudience "AzureADMyOrg" `
        -RequiredResourceAccess $requiredResourceAccess `
        -ErrorAction Stop

    Write-Host "✓ App registration created" -ForegroundColor Green
    Write-Host "  Application ID: $($app.AppId)" -ForegroundColor White
    Write-Host "  Object ID: $($app.Id)" -ForegroundColor White

    # Create client secret
    Write-Host "`n🔑 Creating client secret..." -ForegroundColor Cyan

    $passwordCred = Add-MgApplicationPassword -ApplicationId $app.Id -PasswordCredential @{
        DisplayName = "AvePoint Fly Secret"
        EndDateTime = (Get-Date).AddYears(2)
    } -ErrorAction Stop

    Write-Host "✓ Client secret created" -ForegroundColor Green
    Write-Host "  Expires: $($passwordCred.EndDateTime)" -ForegroundColor Gray

    # Display results
    Write-Host "`n╔══════════════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "║    App Registration Created Successfully                 ║" -ForegroundColor Green
    Write-Host "╚══════════════════════════════════════════════════════════╝" -ForegroundColor Green

    Write-Host "`n📋 IMPORTANT: Save these credentials securely!" -ForegroundColor Yellow
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Yellow

    Write-Host "`nTenant ID:" -ForegroundColor Cyan
    Write-Host "  $TenantId" -ForegroundColor White

    Write-Host "`nApplication (Client) ID:" -ForegroundColor Cyan
    Write-Host "  $($app.AppId)" -ForegroundColor White

    Write-Host "`nClient Secret:" -ForegroundColor Cyan
    Write-Host "  $($passwordCred.SecretText)" -ForegroundColor White

    Write-Host "`nSecret Expiry:" -ForegroundColor Cyan
    Write-Host "  $($passwordCred.EndDateTime)" -ForegroundColor White

    Write-Host "`n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Yellow

    # Save to config
    $save = if ($SkipSavePrompt) { 'N' } else {
        Write-Host "`n💾 Do you want to save these credentials to Migration Toolkit config? (Y/N): " -ForegroundColor Cyan -NoNewline
        Read-Host
    }

    if ($save -eq 'Y' -or $save -eq 'y') {
        $configPath = Join-Path $env:APPDATA "FlyMigration\config.json"
        $configDir = Split-Path $configPath

        if (-not (Test-Path $configDir)) {
            New-Item -ItemType Directory -Path $configDir -Force | Out-Null
        }

        $config = @{}
        if (Test-Path $configPath) {
            $config = Get-Content $configPath -Raw | ConvertFrom-Json -AsHashtable
        }

        $config.TenantId = $TenantId
        $config.ClientId = $app.AppId
        # Note: Secret will be encrypted when saved via UI

        $config | ConvertTo-Json -Depth 10 | Set-Content $configPath -Force
        Write-Host "✓ Credentials saved to config (encrypt secret via Settings)" -ForegroundColor Green
    }

    Write-Host "`n⚠️  NEXT STEPS:" -ForegroundColor Yellow
    Write-Host "1. In Azure Portal, grant admin consent for the API permissions" -ForegroundColor White
    Write-Host "   Portal URL: https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$($app.AppId)" -ForegroundColor Gray
    Write-Host "2. Enter the Client ID and Secret in Migration Toolkit > Settings > Config tab" -ForegroundColor White
    Write-Host "3. Test the connection using the 'Test Connection' button" -ForegroundColor White

} catch {
    Write-Error "Failed to create app registration: $_"
    exit 1
} finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}

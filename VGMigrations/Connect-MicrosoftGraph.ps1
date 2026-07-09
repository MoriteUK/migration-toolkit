#Requires -Version 7.0
<#
.SYNOPSIS
    Connects to Microsoft Graph with the scopes needed for OneDrive provisioning.
    Run this once per session, then use Provision-OneDrive-User.ps1 as many times as needed.
#>

$ErrorActionPreference = 'Stop'

$graphMods = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Files', 'Microsoft.Graph.Users')
foreach ($m in $graphMods) {
    if (-not (Get-Module -ListAvailable -Name $m -ErrorAction SilentlyContinue)) {
        Write-Error "Required module not installed: $m`nRun: Install-Module Microsoft.Graph -Scope CurrentUser -Force"
        exit 1
    }
    Import-Module $m -ErrorAction Stop
}

# Check if already connected with adequate scopes
$ctx = Get-MgContext -ErrorAction SilentlyContinue
if ($ctx) {
    Write-Host "Already connected as: $($ctx.Account)  (tenant: $($ctx.TenantId))" -ForegroundColor Green
    Write-Host "Scopes: $($ctx.Scopes -join ', ')" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Run Provision-OneDrive-User.ps1 to provision." -ForegroundColor Cyan
    exit 0
}

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Write-Host ">>> Visit https://microsoft.com/devicelogin and enter the code shown below <<<" -ForegroundColor Yellow
Write-Host ""

Connect-MgGraph -Scopes 'Sites.ReadWrite.All', 'User.Read.All' -UseDeviceCode -NoWelcome -ErrorAction Stop

$ctx = Get-MgContext
Write-Host ""
Write-Host "Connected as: $($ctx.Account)" -ForegroundColor Green
Write-Host "Tenant:       $($ctx.TenantId)" -ForegroundColor Green
Write-Host ""
Write-Host "You can now run Provision-OneDrive-User.ps1 without re-authenticating." -ForegroundColor Cyan

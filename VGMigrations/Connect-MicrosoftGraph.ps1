#Requires -Version 7.0
<#
.SYNOPSIS
    Connects to Microsoft Graph with the scopes needed for OneDrive provisioning.
    Run this once per session, then use Provision-OneDrive-User.ps1 as many times as needed.
#>

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'Ensure-GraphModules.ps1') -GraphModules @('Microsoft.Graph.Files','Microsoft.Graph.Users')

# Check if already connected with adequate scopes
$ctx = Get-MgContext -ErrorAction SilentlyContinue
if ($ctx) {
    Write-Host "Already connected as: $($ctx.Account)  (tenant: $($ctx.TenantId))" -ForegroundColor Green
    Write-Host "Scopes: $($ctx.Scopes -join ', ')" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Run Provision-OneDrive-User.ps1 to provision." -ForegroundColor Cyan
    exit 0
}

Write-Host "Connecting to Microsoft Graph — sign in with the browser window that opens..." -ForegroundColor Cyan
Write-Host ""

Connect-MgGraph -Scopes 'Sites.ReadWrite.All', 'User.Read.All' -NoWelcome -ErrorAction Stop

$ctx = Get-MgContext
Write-Host ""
Write-Host "Connected as: $($ctx.Account)" -ForegroundColor Green
Write-Host "Tenant:       $($ctx.TenantId)" -ForegroundColor Green
Write-Host ""
Write-Host "You can now run Provision-OneDrive-User.ps1 without re-authenticating." -ForegroundColor Cyan

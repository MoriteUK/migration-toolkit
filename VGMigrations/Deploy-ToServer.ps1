#Requires -Version 7.0
<#
.SYNOPSIS
    Deploy-ToServer.ps1 - Prepare configuration files for server deployment

.DESCRIPTION
    Copies essential configuration files from AppData to the MigrationToolkit folder
    for easy transfer to the server.

.EXAMPLE
    .\Deploy-ToServer.ps1
#>

$ErrorActionPreference = 'Stop'

Write-Host ""
Write-Host "═══════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Server Deployment Preparation" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

$destFolder = $PSScriptRoot

# ── Copy Fly API Config ────────────────────────────────────────────────────────
Write-Host "Copying Fly API configuration..." -ForegroundColor Yellow

$sourceConfig = "$env:APPDATA\FlyMigration\config.json"
$destConfig = Join-Path $destFolder "fly-config.json"

if (Test-Path $sourceConfig) {
    Copy-Item -Path $sourceConfig -Destination $destConfig -Force
    Write-Host "  ✓ Copied Fly API config to: fly-config.json" -ForegroundColor Green

    # Show content (with secret masked)
    $config = Get-Content $sourceConfig | ConvertFrom-Json
    Write-Host "    URL: $($config.Url)" -ForegroundColor Gray
    Write-Host "    ClientId: $($config.ClientId)" -ForegroundColor Gray
    Write-Host "    Secret: [ENCRYPTED]" -ForegroundColor Gray
} else {
    Write-Host "  ✗ No Fly API config found in AppData" -ForegroundColor Red
    Write-Host "    This is normal if you haven't set up Fly credentials yet" -ForegroundColor Yellow
}

Write-Host ""

# ── Check Config Files ─────────────────────────────────────────────────────────
Write-Host "Checking configuration files..." -ForegroundColor Yellow

$configFiles = @(
    @{ Name = 'domains.json';        Description = 'Domain to VBU ID mappings' }
    @{ Name = 'workloads.json';      Description = 'Workload configuration' }
    @{ Name = 'shared-config.json';  Description = 'Customer prefixes and portal URL' }
    @{ Name = 'tenant-sites.json';   Description = 'SharePoint site data' }
    @{ Name = 'version.json';        Description = 'Current version' }
)

foreach ($file in $configFiles) {
    $path = Join-Path $destFolder $file.Name
    if (Test-Path $path) {
        $item = Get-Item $path
        $sizeKB = [math]::Round($item.Length / 1KB, 1)
        Write-Host "  ✓ $($file.Name)" -NoNewline -ForegroundColor Green
        Write-Host " - $sizeKB KB - $($file.Description)" -ForegroundColor Gray
    } else {
        Write-Host "  ○ $($file.Name)" -NoNewline -ForegroundColor Yellow
        Write-Host " - Missing - $($file.Description)" -ForegroundColor Gray
    }
}

Write-Host ""

# ── Create README ──────────────────────────────────────────────────────────────
Write-Host "Deployment package ready!" -ForegroundColor Green
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host "  1. Copy the entire MigrationToolkit folder to your server" -ForegroundColor White
Write-Host "  2. On the server, run this to install Fly API config:" -ForegroundColor White
Write-Host ""
Write-Host '     $destFolder = "$env:APPDATA\FlyMigration"' -ForegroundColor Gray
Write-Host '     New-Item -ItemType Directory -Path $destFolder -Force' -ForegroundColor Gray
Write-Host '     Copy-Item "fly-config.json" -Destination "$destFolder\config.json"' -ForegroundColor Gray
Write-Host ""
Write-Host "  3. Launch the toolkit: .\main-menu.ps1" -ForegroundColor White
Write-Host "  4. Check for updates in Settings > Config" -ForegroundColor White
Write-Host ""
Write-Host "See SERVER-DEPLOYMENT.md for detailed instructions." -ForegroundColor Yellow
Write-Host ""

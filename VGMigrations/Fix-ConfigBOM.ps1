#Requires -Version 5.1
<#
.SYNOPSIS
    Strips the UTF-8 BOM from %APPDATA%\FlyMigration\config.json if present.

.DESCRIPTION
    PowerShell 5.1's Set-Content -Encoding UTF8 writes a BOM header that
    the Electron app's JSON.parse cannot handle.  Run this once on the server
    to clean the file; subsequent saves from the Electron UI will be BOM-free.
#>

$configPath = Join-Path $env:APPDATA 'FlyMigration\config.json'

if (-not (Test-Path $configPath)) {
    Write-Host "No config.json found at $configPath — nothing to fix." -ForegroundColor Yellow
    exit 0
}

$bytes = [System.IO.File]::ReadAllBytes($configPath)

if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
    $stripped = $bytes[3..($bytes.Length - 1)]
    [System.IO.File]::WriteAllBytes($configPath, $stripped)
    Write-Host "BOM removed from config.json." -ForegroundColor Green
} else {
    Write-Host "config.json has no BOM — already clean." -ForegroundColor Green
}

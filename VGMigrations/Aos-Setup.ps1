#Requires -Version 7.0
. "$PSScriptRoot\lib.ps1"

$node = Get-Command node.exe -ErrorAction SilentlyContinue
if (-not $node) {
    $candidates = @(
        "$env:ProgramFiles\nodejs\node.exe",
        "${env:ProgramFiles(x86)}\nodejs\node.exe",
        "$env:APPDATA\nvm\current\node.exe"
    )
    foreach ($c in $candidates) {
        if (Test-Path $c) { $node = [pscustomobject]@{ Source = $c }; break }
    }
}
if (-not $node) {
    Write-Host "ERROR: Node.js not found. Install Node 18+ from https://nodejs.org" -ForegroundColor Red
    exit 1
}

$connector = Join-Path $PSScriptRoot "fly-connector.js"
if (-not (Test-Path $connector)) {
    Write-Host "ERROR: fly-connector.js not found at $connector" -ForegroundColor Red
    exit 1
}

$cfg = Read-SharedConfig
if (-not $cfg.TenantName -or -not $cfg.TenantSearch) {
    Write-Host "ERROR: Tenant details missing. Fill in Display Name and Search Code in AOS Setup first." -ForegroundColor Red
    exit 1
}

$appProfileName = if ($cfg.AppProfileName) { $cfg.AppProfileName } else { "$($cfg.TenantName) App" }

$id   = [guid]::NewGuid().ToString('N').Substring(0, 8)
$task = [pscustomobject]@{
    id                = $id
    tenantDisplayName = $cfg.TenantName
    tenantSearch      = $cfg.TenantSearch
    appProfileName    = $appProfileName
}

Write-Host "AOS App Profile Setup" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""
Write-Host "Tenant      : $($cfg.TenantName) ($($cfg.TenantSearch))" -ForegroundColor Gray
Write-Host "App Profile : $appProfileName" -ForegroundColor Gray
Write-Host ""
Write-Host "Browser automation will run. Approve any consent prompts that appear." -ForegroundColor Yellow
Write-Host ""

$task | ConvertTo-Json -Compress |
    & $node.Source $connector --mode=setup "--display-name=$($cfg.TenantName -replace '[^A-Za-z0-9._-]','_')"

if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "ERROR: Setup failed (exit code $LASTEXITCODE)." -ForegroundColor Red
    exit $LASTEXITCODE
}

Write-Host ""
Write-Host "Setup complete." -ForegroundColor Green

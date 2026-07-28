#Requires -Version 7.0
. "$PSScriptRoot\lib.ps1"

$node = Get-Command node.exe -ErrorAction SilentlyContinue
if (-not $node) {
    # Node may not be in PATH for spawned processes; check common install locations
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

Write-Host "AOS Sign-In" -ForegroundColor Cyan
Write-Host "═══════════════════════════════" -ForegroundColor Cyan
Write-Host ""
Write-Host "A browser window will open. Complete Microsoft SSO, then close the browser." -ForegroundColor Yellow
Write-Host ""

& $node.Source $connector --mode=login

if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "ERROR: Sign-in failed (exit code $LASTEXITCODE)." -ForegroundColor Red
    exit $LASTEXITCODE
}

Write-Host ""
Write-Host "Sign-in complete. Session saved." -ForegroundColor Green

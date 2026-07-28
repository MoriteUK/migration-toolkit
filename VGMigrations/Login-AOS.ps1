#Requires -Version 7.0
<#
.SYNOPSIS
    Opens a browser window so you can sign in to AvePoint Online Services (AOS).
.DESCRIPTION
    Launches fly-connector.js in login mode using Microsoft Edge.
    After you complete SSO the session is saved to auth/storageState.json.
    This saved session is used by the toolkit to auto-create Fly destination connections.
#>

$nodeExe = (Get-Command node.exe -ErrorAction SilentlyContinue)?.Source
if (-not $nodeExe) {
    Write-Error "Node.js not found on PATH. Install Node.js 18+ from nodejs.org."
    exit 1
}

$connectorJs = Join-Path $PSScriptRoot 'fly-connector.js'
if (-not (Test-Path $connectorJs)) {
    Write-Error "fly-connector.js not found at '$PSScriptRoot'."
    exit 1
}

Write-Host "Opening Edge browser for AOS sign-in..." -ForegroundColor Cyan
Write-Host "(Complete Microsoft SSO in the browser window that opens)" -ForegroundColor Gray

& $nodeExe $connectorJs --mode=login
$code = $LASTEXITCODE

if ($code -eq 0) {
    Write-Host "Sign-in complete. AOS session saved." -ForegroundColor Green
} else {
    Write-Error "Sign-in failed or was cancelled (exit code $code)."
    exit $code
}

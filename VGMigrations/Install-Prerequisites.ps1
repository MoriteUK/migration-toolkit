#Requires -Version 5.1
<#
.SYNOPSIS
    Installs all prerequisites required by the Migration Toolkit.

.DESCRIPTION
    Installs:
      - PowerShell 7 (pwsh.exe)
      - Node.js LTS
      - PowerShell modules: ExchangeOnlineManagement, Microsoft.Graph,
        Microsoft.Online.SharePoint.PowerShell, ImportExcel, Fly.Client
        (and optionally Microsoft.PowerApps.Administration.PowerShell)
      - npm packages for the Electron app
      - Playwright + Chromium (for AOS connector automation)
      - Optionally: RSAT Active Directory tools (for Hybrid discovery mode)

    Safe to re-run — already-installed components are skipped.

.PARAMETER ToolkitRoot
    Path to the toolkit root folder. Defaults to the parent of VGMigrations.

.PARAMETER IncludeAD
    Also installs the RSAT Active Directory PowerShell tools (requires restart).

.PARAMETER IncludePowerApps
    Also installs Microsoft.PowerApps.Administration.PowerShell.
#>
param(
    [string]$ToolkitRoot    = (Split-Path $PSScriptRoot -Parent),
    [switch]$IncludeAD,
    [switch]$IncludePowerApps
)

Set-StrictMode -Off
$ErrorActionPreference = 'Continue'

# ── Colour helpers ────────────────────────────────────────────────────────────
function Write-Step  { param([string]$msg) Write-Host "`n── $msg" -ForegroundColor Cyan }
function Write-OK    { param([string]$msg) Write-Host "  ✓  $msg" -ForegroundColor Green }
function Write-Skip  { param([string]$msg) Write-Host "  –  $msg (already installed)" -ForegroundColor DarkGray }
function Write-Warn  { param([string]$msg) Write-Host "  ⚠  $msg" -ForegroundColor Yellow }
function Write-Fail  { param([string]$msg) Write-Host "  ✗  $msg" -ForegroundColor Red }

# ── Admin check ───────────────────────────────────────────────────────────────
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Warn "Not running as Administrator. Some installs (Node.js, system-wide PS modules) may fail."
    Write-Warn "Re-run from an elevated prompt for best results."
}

Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Migration Toolkit — Prerequisites Installer   " -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Toolkit root : $ToolkitRoot"
Write-Host "  Running as   : $($env:USERNAME)  (admin=$isAdmin)"
Write-Host ""

# ── Helper: find winget ───────────────────────────────────────────────────────
function Get-Winget {
    $wg = Get-Command winget.exe -ErrorAction SilentlyContinue
    if ($wg) { return $wg.Source }
    $candidates = @(
        "$env:LOCALAPPDATA\Microsoft\WindowsApps\winget.exe",
        "$env:ProgramFiles\WindowsApps\Microsoft.DesktopAppInstaller*\winget.exe"
    )
    foreach ($c in $candidates) {
        $resolved = (Resolve-Path $c -ErrorAction SilentlyContinue | Select-Object -First 1)
        if ($resolved) { return $resolved.Path }
    }
    return $null
}

$winget = Get-Winget

# ── Helper: install a winget package if not already present ───────────────────
function Install-WingetPackage {
    param([string]$Id, [string]$Name, [string]$TestCmd)
    if ($TestCmd) {
        $found = Get-Command $TestCmd -ErrorAction SilentlyContinue
        if ($found) { Write-Skip $Name; return $true }
    }
    if (-not $winget) {
        Write-Warn "$Name — winget not found. Install manually from https://aka.ms/getwinget"
        return $false
    }
    Write-Host "  Installing $Name via winget..." -ForegroundColor White
    & $winget install --id $Id --source winget --accept-package-agreements --accept-source-agreements --silent 2>&1 |
        ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
    if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq -1978335212) {  # 0=ok, -1978335212=already installed
        Write-OK $Name
        return $true
    }
    Write-Warn "$Name install may have failed (exit $LASTEXITCODE). Check winget logs."
    return $false
}

# ── Helper: install a PS module if not already present ────────────────────────
function Install-PSModule {
    param([string]$Name, [string]$Scope = 'CurrentUser', [string]$MinVersion = '')
    $existing = Get-Module -Name $Name -ListAvailable -ErrorAction SilentlyContinue |
                Sort-Object Version -Descending | Select-Object -First 1
    if ($existing) {
        if ($MinVersion -and ([version]$existing.Version -lt [version]$MinVersion)) {
            Write-Host "  Updating $Name ($($existing.Version) → $MinVersion+)..." -ForegroundColor White
        } else {
            Write-Skip "$Name $($existing.Version)"; return
        }
    } else {
        Write-Host "  Installing $Name..." -ForegroundColor White
    }
    $params = @{ Name = $Name; Scope = $Scope; Force = $true; AllowClobber = $true; ErrorAction = 'Stop' }
    if ($MinVersion) { $params['MinimumVersion'] = $MinVersion }
    try {
        Install-Module @params
        $v = (Get-Module -Name $Name -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1).Version
        Write-OK "$Name $v"
    } catch {
        Write-Fail "$Name — $($_.Exception.Message)"
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# 1. POWERSHELL 7
# ─────────────────────────────────────────────────────────────────────────────
Write-Step "PowerShell 7"
$pwsh = Get-Command pwsh.exe -ErrorAction SilentlyContinue
if ($pwsh) {
    $ver = & pwsh.exe -NoProfile -Command '$PSVersionTable.PSVersion.ToString()'
    Write-Skip "PowerShell $ver"
} else {
    Install-WingetPackage -Id 'Microsoft.PowerShell' -Name 'PowerShell 7' -TestCmd 'pwsh.exe'
    if (-not (Get-Command pwsh.exe -ErrorAction SilentlyContinue)) {
        # Fallback: direct MSI download
        Write-Warn "winget failed — downloading PS7 MSI directly..."
        $msiUrl  = 'https://github.com/PowerShell/PowerShell/releases/download/v7.4.6/PowerShell-7.4.6-win-x64.msi'
        $msiPath = Join-Path $env:TEMP 'PS7.msi'
        try {
            Write-Host "  Downloading from $msiUrl ..." -ForegroundColor White
            Invoke-WebRequest $msiUrl -OutFile $msiPath -UseBasicParsing
            Start-Process msiexec.exe -ArgumentList "/i `"$msiPath`" /quiet ADD_EXPLORER_CONTEXT_MENU_OPENPOWERSHELL=1 ENABLE_PSREMOTING=1 REGISTER_MANIFEST=1" -Wait
            Write-OK 'PowerShell 7 (MSI)'
        } catch {
            Write-Fail "PowerShell 7 — $($_.Exception.Message)"
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# 2. NODE.JS LTS
# ─────────────────────────────────────────────────────────────────────────────
Write-Step "Node.js LTS"
$nodeOk = Install-WingetPackage -Id 'OpenJS.NodeJS.LTS' -Name 'Node.js LTS' -TestCmd 'node.exe'
if (-not $nodeOk) {
    # Fallback direct download
    Write-Warn "winget failed — downloading Node.js MSI directly..."
    $nodeUrl  = 'https://nodejs.org/dist/v22.14.0/node-v22.14.0-x64.msi'
    $nodePath = Join-Path $env:TEMP 'nodejs.msi'
    try {
        Write-Host "  Downloading Node.js v22 LTS..." -ForegroundColor White
        Invoke-WebRequest $nodeUrl -OutFile $nodePath -UseBasicParsing
        Start-Process msiexec.exe -ArgumentList "/i `"$nodePath`" /quiet" -Wait
        Write-OK 'Node.js (MSI)'
    } catch {
        Write-Fail "Node.js — $($_.Exception.Message)"
    }
}

# Refresh PATH so node/npm are available in this session
$env:Path = [System.Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' +
            [System.Environment]::GetEnvironmentVariable('Path', 'User')

# ─────────────────────────────────────────────────────────────────────────────
# 3. POWERSHELL MODULE REPOSITORY
# ─────────────────────────────────────────────────────────────────────────────
Write-Step "PowerShell Gallery (TLS + PSRepository)"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$psgallery = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
if (-not $psgallery -or $psgallery.InstallationPolicy -ne 'Trusted') {
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
}
$pm = Get-Command Install-Module -ErrorAction SilentlyContinue
if (-not $pm) {
    Write-Warn "Install-Module not available. Installing PowerShellGet..."
    try { Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser | Out-Null } catch {}
    try { Install-Module -Name PowerShellGet -Scope CurrentUser -Force -AllowClobber } catch {}
}
Write-OK "PSGallery trusted"

# ─────────────────────────────────────────────────────────────────────────────
# 4. POWERSHELL MODULES
# ─────────────────────────────────────────────────────────────────────────────
Write-Step "Exchange Online Management"
Install-PSModule -Name 'ExchangeOnlineManagement' -MinVersion '3.0.0'

Write-Step "Microsoft Graph"
# Install the full SDK or just the needed sub-modules
$graphMods = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Groups',
    'Microsoft.Graph.Sites',
    'Microsoft.Graph.Applications',
    'Microsoft.Graph.DeviceManagement',
    'Microsoft.Graph.Planner',
    'Microsoft.Graph.Identity.DirectoryManagement',
    'Microsoft.Graph.Identity.SignIns'
)
foreach ($mod in $graphMods) {
    Install-PSModule -Name $mod
}

Write-Step "SharePoint Online Management Shell"
Install-PSModule -Name 'Microsoft.Online.SharePoint.PowerShell'

Write-Step "ImportExcel"
Install-PSModule -Name 'ImportExcel'

Write-Step "Fly.Client"
Install-PSModule -Name 'Fly.Client'

if ($IncludePowerApps) {
    Write-Step "Power Apps Administration"
    Install-PSModule -Name 'Microsoft.PowerApps.Administration.PowerShell'
    Install-PSModule -Name 'Microsoft.PowerApps.PowerShell'
}

# ─────────────────────────────────────────────────────────────────────────────
# 5. NPM PACKAGES (Electron app)
# ─────────────────────────────────────────────────────────────────────────────
Write-Step "npm packages (Electron app)"
$appDir = Join-Path $ToolkitRoot 'MigrationToolkit-Web'
if (Test-Path $appDir) {
    $nodeExe = Get-Command node.exe -ErrorAction SilentlyContinue
    if ($nodeExe) {
        Write-Host "  Running npm install in $appDir..." -ForegroundColor White
        Push-Location $appDir
        try {
            & npm install --prefer-offline 2>&1 | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
            Write-OK 'npm packages installed'
        } catch {
            Write-Warn "npm install failed: $($_.Exception.Message)"
        } finally {
            Pop-Location
        }
    } else {
        Write-Warn "node.exe not on PATH — skipping npm install. Restart shell after Node.js install and re-run."
    }
} else {
    Write-Warn "MigrationToolkit-Web folder not found at $appDir — skipping npm install."
}

# ─────────────────────────────────────────────────────────────────────────────
# 6. PLAYWRIGHT + CHROMIUM (AOS connector)
# ─────────────────────────────────────────────────────────────────────────────
Write-Step "Playwright + Chromium"
$psRoot    = Join-Path $PSScriptRoot 'node_modules'
$playwrightCli = Get-Command npx -ErrorAction SilentlyContinue
if ($playwrightCli) {
    Write-Host "  Checking Playwright install..." -ForegroundColor White
    Push-Location $PSScriptRoot
    try {
        # Install playwright package locally if not present
        if (-not (Test-Path (Join-Path $PSScriptRoot 'node_modules\playwright'))) {
            & npm install playwright --save-dev 2>&1 | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
        }
        & npx playwright install chromium 2>&1 | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
        Write-OK 'Playwright Chromium'
    } catch {
        Write-Warn "Playwright install failed: $($_.Exception.Message)"
    } finally {
        Pop-Location
    }
} else {
    Write-Warn "npx not found — skipping Playwright install. Install Node.js first."
}

# ─────────────────────────────────────────────────────────────────────────────
# 7. RSAT ACTIVE DIRECTORY (optional — needed for Hybrid discovery mode)
# ─────────────────────────────────────────────────────────────────────────────
if ($IncludeAD) {
    Write-Step "RSAT Active Directory Tools"
    $adMod = Get-Module -Name ActiveDirectory -ListAvailable -ErrorAction SilentlyContinue
    if ($adMod) {
        Write-Skip "ActiveDirectory module"
    } else {
        Write-Host "  Installing RSAT-AD-PowerShell (requires internet)..." -ForegroundColor White
        try {
            Add-WindowsCapability -Online -Name 'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0' -ErrorAction Stop | Out-Null
            Write-OK 'RSAT Active Directory Tools (restart may be required)'
        } catch {
            Write-Warn "RSAT install failed: $($_.Exception.Message)"
            Write-Warn "On Windows Server, use: Install-WindowsFeature RSAT-AD-PowerShell"
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# SUMMARY
# ─────────────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Installation complete." -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Next steps:" -ForegroundColor White
Write-Host "  1. If Node.js was freshly installed, close and reopen your shell."
Write-Host "  2. Launch the app:  cd MigrationToolkit-Web && npm start"
Write-Host "  3. For Hybrid discovery mode, re-run with:  -IncludeAD"
Write-Host ""

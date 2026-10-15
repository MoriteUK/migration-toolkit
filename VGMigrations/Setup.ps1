<#
.SYNOPSIS
    One-time setup for the Fly Migration Toolkit.

    Checks for Node.js, installs it from nodejs.org if missing, then runs
    'npm install' and 'npx playwright install chromium' in this folder.

.NOTES
    - Run this from the same folder as menu.ps1 and fly-connector.js / package.json.
    - Admin rights are recommended (system-wide Node install); the script
      falls back to a per-user temp install if not elevated.
    - Re-running is safe; existing components are detected and skipped.
#>

[CmdletBinding()]
param(
    [string]$NodeVersion = '24.15.0'   # Latest Active LTS at time of writing
)

$ErrorActionPreference = 'Stop'

# Folder this script lives in - all subsequent work happens here
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
Set-Location $scriptDir

function Write-Step($message) { Write-Host "`n==> $message" -ForegroundColor Cyan }
function Write-Ok  ($message) { Write-Host "    $message"        -ForegroundColor Green }
function Write-Warn($message) { Write-Host "    $message"        -ForegroundColor Yellow }
function Write-Err ($message) { Write-Host "    $message"        -ForegroundColor Red }

function Test-Admin {
    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $pr = New-Object System.Security.Principal.WindowsPrincipal($id)
    return $pr.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Update-EnvironmentPath {
    # Reload PATH from registry so newly-installed tools become available
    # without restarting the shell.
    $machinePath = [System.Environment]::GetEnvironmentVariable('Path','Machine')
    $userPath    = [System.Environment]::GetEnvironmentVariable('Path','User')
    $env:Path    = ($machinePath, $userPath) -join ';'
}

function Find-NodeExe {
    try {
        $cmd = Get-Command node.exe -ErrorAction Stop
        if ($cmd -and $cmd.Source) { return $cmd.Source }
    } catch { }

    $candidates = @(
        "$env:ProgramFiles\nodejs\node.exe",
        "${env:ProgramFiles(x86)}\nodejs\node.exe",
        "$env:LOCALAPPDATA\Programs\nodejs\node.exe"
    )
    foreach ($p in $candidates) {
        if ($p -and (Test-Path $p)) { return $p }
    }
    return $null
}

# -------------------------------------------------------------------
# Step 1: Verify required project files
# -------------------------------------------------------------------
Write-Step "Verifying project files in $scriptDir"
$required = @(
    'menu.ps1', 'fly-connector.js', 'package.json',
    'install-flymodules.ps1', 'workloads.json'
)
$missing  = $required | Where-Object { -not (Test-Path (Join-Path $scriptDir $_)) }
if ($missing) {
    Write-Err "Missing files: $($missing -join ', ')"
    Write-Err "Place this Setup.ps1 in the same folder as the project files and re-run."
    exit 1
}
Write-Ok "All required files present."

# -------------------------------------------------------------------
# Step 2: Node.js
# -------------------------------------------------------------------
Write-Step "Checking for Node.js"
$node = Find-NodeExe

if ($node) {
    $ver = & $node -v 2>$null
    Write-Ok "Found Node $ver at $node"
} else {
    Write-Warn "Node.js not found. Downloading Node $NodeVersion LTS..."

    $arch     = if ([Environment]::Is64BitOperatingSystem) { 'x64' } else { 'x86' }
    $msiName  = "node-v$NodeVersion-$arch.msi"
    $msiUrl   = "https://nodejs.org/dist/v$NodeVersion/$msiName"
    $msiPath  = Join-Path $env:TEMP $msiName

    Write-Host "    Downloading $msiUrl"
    try {
        # Force TLS 1.2 for older PowerShell sessions
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $msiUrl -OutFile $msiPath -UseBasicParsing
    } catch {
        Write-Err "Download failed: $($_.Exception.Message)"
        Write-Err "Download manually from https://nodejs.org and re-run Setup.ps1."
        exit 2
    }
    Write-Ok "Downloaded to $msiPath"

    Write-Host "    Installing Node.js (silent)..."
    $msiArgs = @('/i', "`"$msiPath`"", '/qn', '/norestart', 'ADDLOCAL=ALL')
    if (-not (Test-Admin)) {
        Write-Warn "Not running as Administrator - installer will prompt for elevation."
    }
    $proc = Start-Process -FilePath 'msiexec.exe' -ArgumentList $msiArgs -Wait -PassThru
    if ($proc.ExitCode -ne 0) {
        Write-Err "Node MSI install failed (exit $($proc.ExitCode))."
        Write-Err "Run the MSI manually: $msiPath"
        exit 3
    }
    Remove-Item $msiPath -ErrorAction SilentlyContinue

    Update-EnvironmentPath
    $node = Find-NodeExe
    if (-not $node) {
        Write-Err "Node installed but still not found on PATH. Open a new PowerShell and re-run this Setup.ps1."
        exit 4
    }
    Write-Ok "Installed Node $(& $node -v) at $node"
}

# Resolve npm and npx alongside node.exe
$nodeDir = Split-Path -Parent $node
$npmCmd  = Join-Path $nodeDir 'npm.cmd'
$npxCmd  = Join-Path $nodeDir 'npx.cmd'

if (-not (Test-Path $npmCmd)) {
    Write-Err "npm.cmd not found next to node.exe. Node install may be incomplete."
    exit 5
}

# -------------------------------------------------------------------
# Step 3: npm install (Playwright)
# -------------------------------------------------------------------
Write-Step "Installing npm dependencies (Playwright)"

if (Test-Path (Join-Path $scriptDir 'node_modules\playwright')) {
    Write-Ok "node_modules already present. Running 'npm install' anyway to verify..."
} else {
    Write-Host "    Running 'npm install'..."
}

& $npmCmd install --no-fund --no-audit
if ($LASTEXITCODE -ne 0) {
    Write-Err "npm install failed (exit $LASTEXITCODE)."
    exit 6
}
Write-Ok "npm dependencies installed."

# -------------------------------------------------------------------
# Step 4: Playwright Chromium browser
# -------------------------------------------------------------------
Write-Step "Installing Playwright's Chromium browser"

# Quick check: is chromium already cached?
$pwCache = Join-Path $env:LOCALAPPDATA 'ms-playwright'
$haveChromium = (Test-Path $pwCache) -and
                (Get-ChildItem -Path $pwCache -Filter 'chromium-*' -Directory -ErrorAction SilentlyContinue)

if ($haveChromium) {
    Write-Ok "Chromium already installed under $pwCache. Re-checking..."
} else {
    Write-Host "    Downloading Chromium (this can take a couple of minutes)..."
}

& $npxCmd playwright install chromium
if ($LASTEXITCODE -ne 0) {
    Write-Err "Playwright Chromium install failed (exit $LASTEXITCODE)."
    Write-Err "Try manually: npx playwright install chromium"
    exit 7
}
Write-Ok "Chromium ready."

# -------------------------------------------------------------------
# Step 5: Fly.Client PowerShell module
# -------------------------------------------------------------------
Write-Step "Checking for Fly.Client PowerShell module"

$flyModule = Get-Module -ListAvailable -Name Fly.Client | Sort-Object Version -Descending | Select-Object -First 1
if ($flyModule) {
    Write-Ok "Fly.Client version $($flyModule.Version) already installed."
} else {
    Write-Warn "Fly.Client not found. Running install-flymodules.ps1..."
    $installScript = Join-Path $scriptDir 'install-flymodules.ps1'
    $proc = Start-Process pwsh.exe `
        -ArgumentList "-NoProfile -ExecutionPolicy Bypass -NonInteractive -File `"$installScript`"" `
        -Wait -PassThru -NoNewWindow
    if ($proc.ExitCode -ne 0) {
        Write-Err "Fly.Client installation failed (exit $($proc.ExitCode))."
        Write-Err "Run install-flymodules.ps1 manually to see the full error."
        exit 8
    }
    $flyModule = Get-Module -ListAvailable -Name Fly.Client | Sort-Object Version -Descending | Select-Object -First 1
    if ($flyModule) {
        Write-Ok "Fly.Client version $($flyModule.Version) installed successfully."
    } else {
        Write-Warn "install-flymodules.ps1 completed but Fly.Client is not yet detectable."
        Write-Warn "Open a new PowerShell session and re-run Setup.ps1 to verify."
    }
}

# -------------------------------------------------------------------
# Step 6: Microsoft Graph PowerShell modules
# -------------------------------------------------------------------
Write-Step "Checking Microsoft Graph PowerShell modules"

$graphModules = @(
    'Microsoft.Graph.Authentication'               # Connect-MgGraph, Get-MgContext
    'Microsoft.Graph.Identity.DirectoryManagement' # Get-MgDomain, Get-MgDevice, Remove-MgDevice
    'Microsoft.Graph.Users'                        # Get-MgUser, Update-MgUser
    'Microsoft.Graph.Applications'                 # Remove-MgApplication, Remove-MgServicePrincipal
)

$graphFailed = [System.Collections.Generic.List[string]]::new()

foreach ($mod in $graphModules) {
    $installed = Get-Module -ListAvailable -Name $mod |
                 Sort-Object Version -Descending | Select-Object -First 1

    if ($installed) {
        Write-Ok "$mod $($installed.Version) already installed."
    } else {
        Write-Warn "$mod not found — installing from PSGallery (CurrentUser scope)..."
        try {
            Install-Module -Name $mod -Scope CurrentUser -Force -AllowClobber `
                           -Repository PSGallery -ErrorAction Stop
            $installed = Get-Module -ListAvailable -Name $mod |
                         Sort-Object Version -Descending | Select-Object -First 1
            if ($installed) {
                Write-Ok "$mod $($installed.Version) installed."
            } else {
                Write-Warn "$mod install reported success but module not yet visible — may need a new PS session."
            }
        } catch {
            Write-Err "$mod installation FAILED: $($_.Exception.Message)"
            $graphFailed.Add($mod)
        }
    }
}

if ($graphFailed.Count -gt 0) {
    Write-Warn ""
    Write-Warn "The following Graph module(s) could not be installed automatically:"
    $graphFailed | ForEach-Object { Write-Warn "  - $_" }
    Write-Warn "Install them manually with:"
    Write-Warn "  Install-Module $($graphFailed -join ', ') -Scope CurrentUser -Force"
    Write-Warn "The toolkit will still start but Graph-dependent features will not work."
} else {
    Write-Ok "All Microsoft Graph modules present."
}

# -------------------------------------------------------------------
# Done
# -------------------------------------------------------------------
Write-Step "Setup complete"
Write-Host ""
Write-Host "    Double-click MigrationTools.exe to launch the toolkit." -ForegroundColor White
Write-Host "    (Or: .\main-menu.ps1 directly from PowerShell)" -ForegroundColor Gray
Write-Host ""

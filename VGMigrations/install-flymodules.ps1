#Requires -Version 7.0
<#
.SYNOPSIS
    install-fly-modules.ps1 — Installs the AvePoint Fly.Client PowerShell module from the
    PowerShell Gallery. Run this once per machine; companion fly-migration.ps1 then dot-sources
    the installed module.

.DESCRIPTION
    Performs all the prep needed to use the Fly.Client cmdlets:
      - Verifies PowerShell 5.1 or 7+
      - Trusts the PSGallery repository (silent, scoped to current process)
      - Installs the NuGet provider if missing
      - Installs or updates Fly.Client (-Scope CurrentUser by default)
      - Imports the module to verify it loads
      - Lists the cmdlets exposed for sanity-checking

.PARAMETER Scope
    Install scope. 'CurrentUser' (default) needs no admin rights. 'AllUsers' requires
    an elevated PowerShell session.

.PARAMETER Force
    Reinstall even if a recent version is already present.

.NOTES
    Version    : 1.0.0
    Last edit  : 2026-05-09
    Author     : Andrew White / Claude (Anthropic)
    Reference  : https://www.powershellgallery.com/packages/Fly.Client
                 https://github.com/AvePoint/fly-client/tree/main/powershell/docs

.EXAMPLE
    .\install-fly-modules.ps1

.EXAMPLE
    .\install-fly-modules.ps1 -Scope AllUsers -Force
    # Run from an elevated PowerShell session.
#>

[CmdletBinding()]
param(
    [ValidateSet('CurrentUser','AllUsers')]
    [string]$Scope = 'CurrentUser',
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

function Write-Step { param([string]$M, [string]$L = 'INFO')
    $col = switch ($L) { 'OK'{'Green'} 'WARN'{'Yellow'} 'ERROR'{'Red'} default{'Cyan'} }
    Write-Host "[$L] $M" -ForegroundColor $col
}

Write-Host ""
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "  AvePoint Fly.Client module installer" -ForegroundColor Cyan
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host ""

# ── 1. PowerShell version check ──
$psVer = $PSVersionTable.PSVersion
if ($psVer.Major -lt 5 -or ($psVer.Major -eq 5 -and $psVer.Minor -lt 1)) {
    Write-Step "PowerShell $psVer detected. Fly.Client requires 5.1 or 7+." 'ERROR'
    exit 1
}
Write-Step "PowerShell $psVer ($($PSVersionTable.PSEdition)) — OK" 'OK'

# ── 2. Admin elevation check (only relevant for AllUsers scope) ──
if ($Scope -eq 'AllUsers') {
    $isAdmin = (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Step "Scope=AllUsers requires an elevated PowerShell session. Re-launch as administrator or use -Scope CurrentUser." 'ERROR'
        exit 1
    }
    Write-Step "Running as administrator — OK for AllUsers scope" 'OK'
}

# ── 3. NuGet provider ──
try {
    $nuget = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
    if (-not $nuget -or $nuget.Version -lt [version]'2.8.5.201') {
        Write-Step "Installing NuGet package provider..."
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope $Scope -ErrorAction Stop | Out-Null
        Write-Step "NuGet installed." 'OK'
    } else {
        Write-Step "NuGet provider $($nuget.Version) — already present" 'OK'
    }
} catch {
    Write-Step "Failed to install NuGet provider: $($_.Exception.Message)" 'ERROR'
    exit 1
}

# ── 4. Trust the PSGallery (silent, current process only) ──
try {
    $repo = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
    if ($repo -and $repo.InstallationPolicy -ne 'Trusted') {
        Write-Step "Setting PSGallery as Trusted (so install runs without prompts)..."
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop
        Write-Step "PSGallery now trusted." 'OK'
    } else {
        Write-Step "PSGallery already trusted (or repo entry not found — Install-Module will create it)" 'OK'
    }
} catch {
    Write-Step "Could not adjust PSGallery trust: $($_.Exception.Message). Continuing — will get a confirm prompt during install." 'WARN'
}

# ── 5. Install or update Fly.Client ──
$moduleName = 'Fly.Client'
$existing = Get-Module -ListAvailable -Name $moduleName | Sort-Object Version -Descending | Select-Object -First 1

if ($existing -and -not $Force) {
    Write-Step "Existing $moduleName found: version $($existing.Version) at $($existing.ModuleBase)"
    Write-Step "Checking PSGallery for newer version..."
    try {
        $latest = Find-Module -Name $moduleName -Repository PSGallery -ErrorAction Stop
        if ([version]$latest.Version -gt [version]$existing.Version) {
            Write-Step "Newer version available: $($latest.Version) (installed: $($existing.Version)). Updating..."
            Install-Module -Name $moduleName -Scope $Scope -Force -AllowClobber -ErrorAction Stop
            Write-Step "Updated to $($latest.Version)" 'OK'
        } else {
            Write-Step "Already on the latest version ($($existing.Version)). Use -Force to reinstall anyway." 'OK'
        }
    } catch {
        Write-Step "Couldn't reach PSGallery to check latest. Existing copy will be used: $($_.Exception.Message)" 'WARN'
    }
} else {
    Write-Step "Installing $moduleName from PSGallery (Scope=$Scope)..."
    try {
        Install-Module -Name $moduleName -Scope $Scope -Force:$Force -AllowClobber -ErrorAction Stop
        Write-Step "$moduleName installed." 'OK'
    } catch {
        Write-Step "Install failed: $($_.Exception.Message)" 'ERROR'
        exit 1
    }
}

# ── 6. Import + sanity check ──
try {
    Import-Module $moduleName -Force -ErrorAction Stop
    $loaded = Get-Module $moduleName | Select-Object -First 1
    Write-Step "Imported $moduleName version $($loaded.Version) from $($loaded.ModuleBase)" 'OK'
} catch {
    Write-Step "Module installed but failed to import: $($_.Exception.Message)" 'ERROR'
    exit 1
}

# ── 7. Probe expected cmdlets ──
$expected = @(
    'Connect-Fly','Disconnect-Fly',
    'New-FlyMigrationProject','Get-FlyMigrationProject','Import-FlyMigrationProjects',
    'Import-FlyExchangeMappings','Start-FlyExchangePreScan','Start-FlyExchangeVerification','Start-FlyExchangeMigration',
    'Import-FlyOneDriveMappings','Start-FlyOneDrivePreScan','Start-FlyOneDriveVerification','Start-FlyOneDriveMigration',
    'Import-FlySharePointMappings','Start-FlySharePointPreScan','Start-FlySharePointVerification','Start-FlySharePointMigration',
    'Import-FlyM365GroupMappings','Start-FlyM365GroupPreScan','Start-FlyM365GroupVerification','Start-FlyM365GroupMigration',
    'Import-FlyTeamChatMappings','Start-FlyTeamChatVerification','Start-FlyTeamChatMigration'
)
$missing = $expected | Where-Object { -not (Get-Command $_ -ErrorAction SilentlyContinue) }
if ($missing.Count -eq 0) {
    Write-Step "All $($expected.Count) expected Fly cmdlets are available." 'OK'
} else {
    Write-Step "These expected cmdlets are missing — version $($loaded.Version) may have renamed them: $($missing -join ', ')" 'WARN'
}

Write-Host ""
Write-Host "========================================================" -ForegroundColor Green
Write-Host "  Done. Fly.Client is ready." -ForegroundColor Green
Write-Host "========================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Next: dot-source fly-migration.ps1 in your PowerShell session:" -ForegroundColor Cyan
Write-Host "  . .\fly-migration.ps1" -ForegroundColor Gray
Write-Host "Then run:" -ForegroundColor Cyan
Write-Host "  Save-FlyCredential       # one-time, encrypts client secret" -ForegroundColor Gray
Write-Host "  Connect-FlyTenant" -ForegroundColor Gray
Write-Host ""
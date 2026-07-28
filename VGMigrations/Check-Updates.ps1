#Requires -Version 7.0
<#
.SYNOPSIS
    Check-Updates.ps1 - Check for and install updates from GitHub

.DESCRIPTION
    Checks the GitHub repository for updates to the Migration Toolkit.
    If updates are available, downloads and installs them automatically.
    Preserves user configuration files.

.PARAMETER GitHubRepo
    GitHub repository in format "username/repository"

.PARAMETER Silent
    Run in silent mode (no UI, auto-install updates)

.PARAMETER Force
    Force check for updates even if recently checked
#>

param(
    [string]$GitHubRepo = "MoriteUK/migration-toolkit",
    [switch]$Silent,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$ScriptRoot = $PSScriptRoot

# ── Logging ────────────────────────────────────────────────────────────────────
$_logDir = Join-Path $ScriptRoot 'logs'
if (-not (Test-Path $_logDir)) { New-Item -ItemType Directory -Path $_logDir -Force | Out-Null }
$LogFile = Join-Path $_logDir "updates-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

function Write-UpdateLog {
    param([string]$Message, [string]$Level = 'INFO')
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logLine = "[$timestamp] [$Level] $Message"
    $logLine | Add-Content -Path $LogFile -Encoding UTF8 -ErrorAction SilentlyContinue

    if (-not $Silent) {
        switch ($Level) {
            'ERROR' { Write-Host $Message -ForegroundColor Red }
            'WARN'  { Write-Host $Message -ForegroundColor Yellow }
            'OK'    { Write-Host $Message -ForegroundColor Green }
            default { Write-Host $Message }
        }
    }
}

Write-UpdateLog "=== Update Check Started ==="
Write-UpdateLog "Repository: $GitHubRepo"

# ── Version Check Cache ────────────────────────────────────────────────────────
$UpdateCachePath = Join-Path $env:LOCALAPPDATA "FlyMigration\update-cache.json"
$UpdateCacheDir = Split-Path $UpdateCachePath

if (-not (Test-Path $UpdateCacheDir)) {
    New-Item -ItemType Directory -Path $UpdateCacheDir -Force | Out-Null
}

function Get-UpdateCache {
    if (Test-Path $UpdateCachePath) {
        try {
            return Get-Content $UpdateCachePath -Raw | ConvertFrom-Json
        } catch {
            return $null
        }
    }
    return $null
}

function Set-UpdateCache {
    param($Data)
    $Data | ConvertTo-Json -Depth 10 | Set-Content $UpdateCachePath -Encoding UTF8
}

# Check if we recently checked for updates (within last 24 hours)
if (-not $Force) {
    $cache = Get-UpdateCache
    if ($cache -and $cache.LastCheck) {
        $lastCheck = [DateTime]::Parse($cache.LastCheck)
        $hoursSinceCheck = ([DateTime]::Now - $lastCheck).TotalHours

        if ($hoursSinceCheck -lt 24) {
            Write-UpdateLog "Last update check was $([Math]::Round($hoursSinceCheck, 1)) hours ago. Skipping check (use -Force to override)"
            return
        }
    }
}

# ── Get Current Version ────────────────────────────────────────────────────────
$LocalVersionPath = Join-Path $ScriptRoot 'version.json'
if (-not (Test-Path $LocalVersionPath)) {
    Write-UpdateLog "Local version.json not found. Creating default version file." 'WARN'
    @{
        version = "1.0.0"
        releaseDate = (Get-Date -Format 'yyyy-MM-dd')
    } | ConvertTo-Json | Set-Content $LocalVersionPath -Encoding UTF8
}

$LocalVersion = (Get-Content $LocalVersionPath -Raw | ConvertFrom-Json).version
Write-UpdateLog "Current version: $LocalVersion"

# ── Check GitHub for Updates ───────────────────────────────────────────────────
Write-UpdateLog "Checking GitHub for updates..."

try {
    # Get latest version.json from GitHub
    $GitHubVersionUrl = "https://raw.githubusercontent.com/$GitHubRepo/main/VGMigrations/version.json"
    Write-UpdateLog "Contacting GitHub..."
    $RemoteVersionJson = Invoke-RestMethod -Uri $GitHubVersionUrl -TimeoutSec 20 -ErrorAction Stop
    $RemoteVersion = $RemoteVersionJson.version

    Write-UpdateLog "Remote version: $RemoteVersion"

    # Update cache
    Set-UpdateCache @{
        LastCheck = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        RemoteVersion = $RemoteVersion
        LocalVersion = $LocalVersion
    }

    # Compare versions
    $LocalVer = [Version]::Parse($LocalVersion)
    $RemoteVer = [Version]::Parse($RemoteVersion)

    if ($RemoteVer -le $LocalVer) {
        Write-UpdateLog "You have the latest version ($LocalVersion)" 'OK'
        return
    }

    Write-UpdateLog "Update available: $LocalVersion -> $RemoteVersion" 'OK'
    Write-Host "UPDATE_AVAILABLE"

} catch {
    Write-UpdateLog "Failed to check for updates: $($_.Exception.Message)" 'ERROR'
    Write-UpdateLog "  Make sure the GitHub repository is accessible and the URL is correct." 'WARN'
    return
}

# ── Show what's new and auto-install ──────────────────────────────────────────
$changelog = $RemoteVersionJson.changelog | Where-Object { $_.version -eq $RemoteVersion } | Select-Object -First 1

Write-Host "`n╔═══════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║    UPDATE AVAILABLE — INSTALLING          ║" -ForegroundColor Cyan
Write-Host "╚═══════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host "`nCurrent version: " -NoNewline
Write-Host $LocalVersion -ForegroundColor Yellow
Write-Host "Latest version:  " -NoNewline
Write-Host $RemoteVersion -ForegroundColor Green

if ($changelog) {
    Write-Host "`nWhat's new in $RemoteVersion`:" -ForegroundColor Cyan
    foreach ($change in $changelog.changes) {
        Write-Host "  • $change" -ForegroundColor Gray
    }
}

Write-Host ""

# ── Download and Install Update ───────────────────────────────────────────────
Write-UpdateLog "Downloading update from GitHub..."

try {
    # Create temp directory for download
    $TempDir = Join-Path $env:TEMP "MigrationToolkit-Update-$(Get-Date -Format 'yyyyMMddHHmmss')"
    New-Item -ItemType Directory -Path $TempDir -Force | Out-Null

    # Download ZIP from GitHub
    $ZipUrl = "https://github.com/$GitHubRepo/archive/refs/heads/main.zip"
    $ZipPath = Join-Path $TempDir "update.zip"

    Write-UpdateLog "Downloading from: $ZipUrl"
    Invoke-WebRequest -Uri $ZipUrl -OutFile $ZipPath -TimeoutSec 120 -ErrorAction Stop
    Write-UpdateLog "Download complete" 'OK'

    # Extract ZIP
    Write-UpdateLog "Extracting files..."
    Expand-Archive -Path $ZipPath -DestinationPath $TempDir -Force

    # Find extracted folder (GitHub adds repo name to folder)
    $ExtractedFolder = Get-ChildItem -Path $TempDir -Directory | Select-Object -First 1
    $RepoRoot  = $ExtractedFolder.FullName

    $VGSource  = Join-Path $RepoRoot 'VGMigrations'
    $WebSource = Join-Path $RepoRoot 'MigrationToolkit-Web'
    if (-not (Test-Path $VGSource)) {
        throw "Expected VGMigrations subfolder not found in extracted archive: $VGSource"
    }

    # Destination roots derived from script location (ScriptRoot = ...VGMigrations\)
    $ToolkitRoot = Split-Path $ScriptRoot -Parent
    $VGDest      = $ScriptRoot
    $WebDest     = Join-Path $ToolkitRoot 'MigrationToolkit-Web'

    # Backup user configuration files
    Write-UpdateLog "Backing up user configuration..."
    $BackupDir = Join-Path $ScriptRoot "backup-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
    New-Item -ItemType Directory -Path $BackupDir -Force | Out-Null

    # Script root config files
    $ConfigFiles = @(
        'domains.json'
        'workloads.json'
        'tenant-sites.json'
        'tenant-sites.meta.json'
        'shared-config.json'
    )

    foreach ($file in $ConfigFiles) {
        $filePath = Join-Path $ScriptRoot $file
        if (Test-Path $filePath) {
            Copy-Item -Path $filePath -Destination $BackupDir -Force
            Write-UpdateLog "  Backed up: $file"
        }
    }

    # APPDATA config files (Fly API credentials)
    $AppDataConfigPath = Join-Path $env:APPDATA 'FlyMigration\config.json'
    if (Test-Path $AppDataConfigPath) {
        Copy-Item -Path $AppDataConfigPath -Destination (Join-Path $BackupDir 'appdata-config.json') -Force
        Write-UpdateLog "  Backed up: Fly API config (from APPDATA)"
    }

    # LOCALAPPDATA config files
    $LocalAppDataConfigPath = Join-Path $env:LOCALAPPDATA 'FlyMigration\shared-config.json'
    if (Test-Path $LocalAppDataConfigPath) {
        Copy-Item -Path $LocalAppDataConfigPath -Destination (Join-Path $BackupDir 'localappdata-shared-config.json') -Force
        Write-UpdateLog "  Backed up: Shared config (from LOCALAPPDATA)"
    }

    # Helper: copy all files from a source folder to a destination folder
    function Copy-UpdateFolder {
        param([string]$Source, [string]$Dest, [string[]]$Exclude)
        Get-ChildItem -Path $Source -Recurse -File | ForEach-Object {
            $RelativePath = $_.FullName.Substring($Source.Length + 1)
            $DestPath = Join-Path $Dest $RelativePath

            $Skip = $false
            foreach ($pattern in $Exclude) {
                if ($RelativePath -like $pattern -or $_.Name -like $pattern -or
                    $RelativePath -like "$pattern\*" -or $RelativePath -like "$pattern/*") {
                    $Skip = $true; break
                }
            }

            if (-not $Skip) {
                $DestDir = Split-Path $DestPath
                if (-not (Test-Path $DestDir)) {
                    New-Item -ItemType Directory -Path $DestDir -Force | Out-Null
                }
                Copy-Item -Path $_.FullName -Destination $DestPath -Force
            }
        }
    }

    # Copy VGMigrations files
    Write-UpdateLog "Installing update (VGMigrations)..."
    $VGExclude = @('*.log', 'logs', 'reports', 'backup-*', 'auth') + $ConfigFiles
    Copy-UpdateFolder -Source $VGSource -Dest $VGDest -Exclude $VGExclude

    # Copy MigrationToolkit-Web files (skip node_modules and build artifacts)
    if (Test-Path $WebSource) {
        Write-UpdateLog "Installing update (MigrationToolkit-Web)..."
        $WebExclude = @('node_modules', 'dist', '*.log')
        Copy-UpdateFolder -Source $WebSource -Dest $WebDest -Exclude $WebExclude
    } else {
        Write-UpdateLog "MigrationToolkit-Web not found in archive — skipping web update." 'WARN'
    }

    # Restore user configuration
    Write-UpdateLog "Restoring user configuration..."
    foreach ($file in $ConfigFiles) {
        $backupPath = Join-Path $BackupDir $file
        if (Test-Path $backupPath) {
            Copy-Item -Path $backupPath -Destination (Join-Path $ScriptRoot $file) -Force
            Write-UpdateLog "  Restored: $file"
        }
    }

    # Restore APPDATA config (Fly API credentials)
    $appDataBackupPath = Join-Path $BackupDir 'appdata-config.json'
    if (Test-Path $appDataBackupPath) {
        $appDataDir = Join-Path $env:APPDATA 'FlyMigration'
        if (-not (Test-Path $appDataDir)) { New-Item -ItemType Directory -Path $appDataDir -Force | Out-Null }
        Copy-Item -Path $appDataBackupPath -Destination (Join-Path $appDataDir 'config.json') -Force
        Write-UpdateLog "  Restored: Fly API config (to APPDATA)"
    }

    # Restore LOCALAPPDATA config
    $localAppDataBackupPath = Join-Path $BackupDir 'localappdata-shared-config.json'
    if (Test-Path $localAppDataBackupPath) {
        $localAppDataDir = Join-Path $env:LOCALAPPDATA 'FlyMigration'
        if (-not (Test-Path $localAppDataDir)) { New-Item -ItemType Directory -Path $localAppDataDir -Force | Out-Null }
        Copy-Item -Path $localAppDataBackupPath -Destination (Join-Path $localAppDataDir 'shared-config.json') -Force
        Write-UpdateLog "  Restored: Shared config (to LOCALAPPDATA)"
    }

    Write-UpdateLog "Update installed successfully!" 'OK'
    Write-UpdateLog "Updated to version: $RemoteVersion" 'OK'

    # Cleanup
    Remove-Item -Path $TempDir -Recurse -Force -ErrorAction SilentlyContinue

    if (-not $Silent) {
        Write-Host "`n╔═══════════════════════════════════════════╗" -ForegroundColor Green
        Write-Host "║    UPDATE COMPLETE                        ║" -ForegroundColor Green
        Write-Host "╚═══════════════════════════════════════════╝" -ForegroundColor Green
        Write-Host "`nThe Migration Toolkit have been updated to version $RemoteVersion" -ForegroundColor Green
        Write-Host "Your configuration files have been preserved.`n" -ForegroundColor Gray
        Write-Host "Backup location: $BackupDir`n" -ForegroundColor Gray
    }

} catch {
    Write-UpdateLog "Update failed: $($_.Exception.Message)" 'ERROR'
    Write-UpdateLog "  Stack: $($_.ScriptStackTrace)" 'ERROR'

    if (-not $Silent) {
        Write-Host "`nUpdate failed. Please check the log file for details:" -ForegroundColor Red
        Write-Host $LogFile -ForegroundColor Yellow
    }
}

Write-UpdateLog "=== Update Check Complete ==="

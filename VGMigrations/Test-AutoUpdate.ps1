#Requires -Version 7.0
<#
.SYNOPSIS
    Test-AutoUpdate.ps1 - Diagnostic script for auto-update system

.DESCRIPTION
    Tests all components of the auto-update system to identify issues.
    Run this on your server to diagnose why updates aren't working.
#>

$ErrorActionPreference = 'Continue'

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  Auto-Update Diagnostic Test" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# ── Test 1: Check local version.json ───────────────────────────────────────────
Write-Host "[1] Checking local version.json..." -ForegroundColor Yellow
$versionPath = Join-Path $PSScriptRoot 'version.json'
if (Test-Path $versionPath) {
    try {
        $localVersion = (Get-Content $versionPath -Raw | ConvertFrom-Json).version
        Write-Host "    ✓ Local version: $localVersion" -ForegroundColor Green
    } catch {
        Write-Host "    ✗ Error reading version.json: $_" -ForegroundColor Red
    }
} else {
    Write-Host "    ✗ version.json not found!" -ForegroundColor Red
}

# ── Test 2: Check GitHub connectivity ──────────────────────────────────────────
Write-Host "`n[2] Testing GitHub connectivity..." -ForegroundColor Yellow
try {
    $timestamp = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    $url = "https://raw.githubusercontent.com/MoriteUK/AvepointFlyUtility/main/version.json?t=$timestamp"
    Write-Host "    URL: $url" -ForegroundColor Gray

    $remoteJson = Invoke-RestMethod -Uri $url -ErrorAction Stop
    $remoteVersion = $remoteJson.version
    Write-Host "    ✓ GitHub accessible" -ForegroundColor Green
    Write-Host "    ✓ Remote version: $remoteVersion" -ForegroundColor Green
} catch {
    Write-Host "    ✗ Cannot reach GitHub: $_" -ForegroundColor Red
    Write-Host "    Check internet connection and firewall" -ForegroundColor Yellow
}

# ── Test 3: Compare versions ───────────────────────────────────────────────────
Write-Host "`n[3] Comparing versions..." -ForegroundColor Yellow
if ($localVersion -and $remoteVersion) {
    try {
        $local = [Version]::Parse($localVersion)
        $remote = [Version]::Parse($remoteVersion)

        if ($remote -gt $local) {
            Write-Host "    ⚠ Update available: $localVersion → $remoteVersion" -ForegroundColor Yellow
        } elseif ($remote -eq $local) {
            Write-Host "    ✓ Already up to date ($localVersion)" -ForegroundColor Green
        } else {
            Write-Host "    ⚠ Local version is NEWER than remote? ($localVersion > $remoteVersion)" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "    ✗ Error comparing versions: $_" -ForegroundColor Red
    }
}

# ── Test 4: Check Check-Updates.ps1 exists ─────────────────────────────────────
Write-Host "`n[4] Checking for Check-Updates.ps1..." -ForegroundColor Yellow
$updateScript = Join-Path $PSScriptRoot 'Check-Updates.ps1'
if (Test-Path $updateScript) {
    Write-Host "    ✓ Check-Updates.ps1 found" -ForegroundColor Green
    $scriptSize = (Get-Item $updateScript).Length
    Write-Host "    Size: $scriptSize bytes" -ForegroundColor Gray
} else {
    Write-Host "    ✗ Check-Updates.ps1 NOT FOUND!" -ForegroundColor Red
    Write-Host "    This is required for auto-updates" -ForegroundColor Yellow
}

# ── Test 5: Check update cache ─────────────────────────────────────────────────
Write-Host "`n[5] Checking update cache..." -ForegroundColor Yellow
$cachePath = Join-Path $env:LOCALAPPDATA "FlyMigration\update-cache.json"
if (Test-Path $cachePath) {
    try {
        $cache = Get-Content $cachePath -Raw | ConvertFrom-Json
        Write-Host "    ✓ Cache found" -ForegroundColor Green
        Write-Host "    Last check: $($cache.LastCheck)" -ForegroundColor Gray
        Write-Host "    Cached remote version: $($cache.RemoteVersion)" -ForegroundColor Gray
        Write-Host "    Cached local version: $($cache.LocalVersion)" -ForegroundColor Gray
    } catch {
        Write-Host "    ⚠ Cache exists but cannot read: $_" -ForegroundColor Yellow
    }
} else {
    Write-Host "    ⚠ No cache file (first run?)" -ForegroundColor Yellow
}

# ── Test 6: Test Check-Updates.ps1 directly ────────────────────────────────────
Write-Host "`n[6] Running Check-Updates.ps1 with -Force..." -ForegroundColor Yellow
if (Test-Path $updateScript) {
    try {
        Write-Host "    Running..." -ForegroundColor Gray
        $result = & $updateScript -Force 2>&1 | Out-String

        if ($result -match 'Current version:\s*(\S+)') {
            Write-Host "    Current version: $($matches[1])" -ForegroundColor Gray
        }
        if ($result -match 'Remote version:\s*(\S+)') {
            Write-Host "    Remote version: $($matches[1])" -ForegroundColor Gray
        }
        if ($result -match 'Update available') {
            Write-Host "    ⚠ Update available!" -ForegroundColor Yellow
        } elseif ($result -match 'You have the latest version|Already up to date') {
            Write-Host "    ✓ Already up to date" -ForegroundColor Green
        }

        # Show last 5 lines of output
        Write-Host "`n    Last lines of output:" -ForegroundColor Gray
        $result.Split("`n") | Select-Object -Last 5 | ForEach-Object {
            Write-Host "      $_" -ForegroundColor Gray
        }
    } catch {
        Write-Host "    ✗ Error running Check-Updates.ps1: $_" -ForegroundColor Red
    }
}

# ── Test 7: Check main-menu.ps1 update check code ──────────────────────────────
Write-Host "`n[7] Checking main-menu.ps1 update integration..." -ForegroundColor Yellow
$mainMenuPath = Join-Path $PSScriptRoot 'main-menu.ps1'
if (Test-Path $mainMenuPath) {
    $content = Get-Content $mainMenuPath -Raw
    if ($content -match 'Check-Updates\.ps1') {
        Write-Host "    ✓ main-menu.ps1 references Check-Updates.ps1" -ForegroundColor Green
    } else {
        Write-Host "    ✗ main-menu.ps1 does NOT reference Check-Updates.ps1!" -ForegroundColor Red
    }

    if ($content -match '\$script:UpdateAvailable') {
        Write-Host "    ✓ UpdateAvailable variable exists" -ForegroundColor Green
    } else {
        Write-Host "    ⚠ UpdateAvailable variable not found" -ForegroundColor Yellow
    }

    if ($content -match 'updateBanner') {
        Write-Host "    ✓ Update banner code exists" -ForegroundColor Green
    } else {
        Write-Host "    ⚠ Update banner code not found" -ForegroundColor Yellow
    }
} else {
    Write-Host "    ✗ main-menu.ps1 not found!" -ForegroundColor Red
}

# ── Summary ────────────────────────────────────────────────────────────────────
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  Diagnostic Summary" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

Write-Host "Local version:  " -NoNewline
if ($localVersion) { Write-Host $localVersion -ForegroundColor Green } else { Write-Host "UNKNOWN" -ForegroundColor Red }

Write-Host "Remote version: " -NoNewline
if ($remoteVersion) { Write-Host $remoteVersion -ForegroundColor Green } else { Write-Host "UNKNOWN" -ForegroundColor Red }

Write-Host "`nNext steps:" -ForegroundColor Yellow
if ($localVersion -and $remoteVersion) {
    $local = [Version]::Parse($localVersion)
    $remote = [Version]::Parse($remoteVersion)
    if ($remote -gt $local) {
        Write-Host "  1. Run main-menu.ps1 - you should see yellow update banner" -ForegroundColor White
        Write-Host "  2. OR run: .\Check-Updates.ps1 -Force" -ForegroundColor White
        Write-Host "  3. If banner doesn't show, check logs in logs/ folder" -ForegroundColor White
    } else {
        Write-Host "  ✓ You're already on the latest version!" -ForegroundColor Green
    }
} else {
    Write-Host "  ✗ Cannot determine versions - check GitHub connectivity" -ForegroundColor Red
}

Write-Host ""

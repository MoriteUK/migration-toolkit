# Auto-Update System

## Overview
The Migration Tools now automatically check for updates from GitHub every time they run.

## How It Works

### 1. **Automatic Check on Launch**
- When `main-menu.ps1` starts, it runs `Check-Updates.ps1` in the background
- Checks GitHub repository: `MoriteUK/AvepointFlyUtility`
- Compares local `version.json` with remote `version.json`
- If update available, shows yellow banner at top of main menu

### 2. **Update Check Interval**
Currently configured to check **every time** you run the tools.

To change to weekly (7 days):
Edit `main-menu.ps1` line ~758:
```powershell
# Change from:
& $CheckUpdatesScript -Silent

# To:
& $CheckUpdatesScript -Silent -CheckIntervalHours 168  # 168 hours = 7 days
```

Common intervals:
- `0` = Every time (current setting)
- `24` = Daily
- `168` = Weekly (7 days)
- `336` = Bi-weekly (14 days)

### 3. **What Gets Backed Up**
The auto-update process preserves:

**Script Directory:**
- `domains.json` - Domain to VBU ID mappings
- `workloads.json` - Workload configuration
- `tenant-sites.json` - SharePoint site data
- `tenant-sites.meta.json` - Site metadata
- `shared-config.json` - Customer prefixes

**APPDATA (`$env:APPDATA\FlyMigration\`):**
- `config.json` - **Fly API credentials (Client ID & encrypted Secret)**

**LOCALAPPDATA (`$env:LOCALAPPDATA\FlyMigration\`):**
- `shared-config.json` - Shared configuration settings

### 4. **Update Process**
When you click "Install Update":

1. Downloads latest version from GitHub
2. Creates `backup-YYYYMMDD-HHMMSS` folder with all config files
3. Copies new files (excluding config)
4. Restores your configuration files
5. Shows success message

### 5. **Server Deployment**
On the server (or any remote copy):

1. **First Time Setup:**
   ```powershell
   # Copy the toolkit folder to the server
   # On the server, restore Fly API config:
   $destFolder = "$env:APPDATA\FlyMigration"
   New-Item -ItemType Directory -Path $destFolder -Force
   Copy-Item "fly-config.json" -Destination "$destFolder\config.json"
   ```

2. **Run the toolkit:**
   ```powershell
   .\main-menu.ps1
   ```

3. **It will automatically:**
   - Check for updates from GitHub
   - Show update notification if available
   - Allow one-click update while preserving credentials

## Version History

### v2.1.4 (2026-06-02) - **Current**
- ✅ Fixed auto-update preserving Fly API credentials
- ✅ Fixed update notification showing in main menu
- ✅ Fixed SECRET EXPIRY field layout
- ✅ Fixed Customer grid cell corruption
- ✅ Configurable check interval (default: every time)

### v2.1.3 (2026-06-01)
- Update notification banner
- One-click install button

### v2.1.2 (2026-06-01)
- Added version.json to package
- Complete auto-update system

## Troubleshooting

### Update notification not showing?
Check logs in `logs/` folder for error messages from update check.

### Credentials lost after update?
This should no longer happen in v2.1.4+. If it does:
1. Check `backup-*` folder for `appdata-config.json`
2. Restore manually:
   ```powershell
   Copy-Item "backup-YYYYMMDD-HHMMSS\appdata-config.json" `
            -Destination "$env:APPDATA\FlyMigration\config.json"
   ```

### Force update check?
From PowerShell:
```powershell
.\Check-Updates.ps1 -Force
```

### Check current version?
```powershell
Get-Content version.json | ConvertFrom-Json | Select-Object version, releaseDate
```

## GitHub Repository
https://github.com/MoriteUK/AvepointFlyUtility

All changes pushed to `main` branch are immediately available to all copies.

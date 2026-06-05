# Server Deployment Guide

This guide explains how to deploy the Migration Tools to your server.

## Quick Deployment Steps

### 1. Get the Latest Version

**Option A: Clone from GitHub (Recommended)**
```powershell
# Install git if needed
winget install Git.Git

# Clone the repository
cd C:\Scripts
git clone https://github.com/MoriteUK/AvepointFlyUtility.git MigrationToolkit
cd MigrationToolkit
```

**Option B: Download ZIP from GitHub**
1. Go to https://github.com/MoriteUK/AvepointFlyUtility
2. Click "Code" → "Download ZIP"
3. Extract to `C:\Scripts\MigrationToolkit\`

**Option C: Copy from existing installation**
Copy the entire `MigrationToolkit` folder to your server:
```
C:\Scripts\MigrationToolkit\
```

### 2. Set Up Fly API Configuration

The Fly API configuration needs to be placed in the server's AppData folder.

**Option A: Copy the config file (if same user account)**
```powershell
# On your LOCAL machine, the config is here:
C:\Users\[YourUsername]\AppData\Roaming\FlyMigration\config.json

# Copy it to the SERVER at:
C:\Users\[ServerUsername]\AppData\Roaming\FlyMigration\config.json
```

**Option B: Enter credentials via Settings (Recommended)**
1. Launch the toolkit: `.\main-menu.ps1`
2. Click the ⚙ Settings icon
3. Go to **Config** tab
4. Enter:
   - **Fly API URL**: `https://graph.avepointonlineservices.com/fly`
   - **Client ID**: `e982b3fe-ebb8-4926-aa2e-03065c6a8407`
   - **Client Secret**: (your secret - get from original machine if needed)
5. Click **Test Connection** to verify
6. Close Settings (auto-saves)

### 3. Transfer Configuration (Optional)

If you have an existing installation, you can transfer these config files:

**From Script Directory:**
- `domains.json` - Domain to VBU ID mappings
- `workloads.json` - Workload configuration  
- `tenant-sites.json` - SharePoint site data (if using)
- `shared-config.json` - Customer prefixes, portal URL, secret expiry

**From AppData (if copying from same user):**
- `%APPDATA%\FlyMigration\config.json` - Fly API credentials

> **Note:** You can also configure everything from scratch via Settings after launching the toolkit.

### 4. Install Prerequisites

**On the server, you need:**

1. **PowerShell 7+**
   ```powershell
   # Check version
   $PSVersionTable.PSVersion
   
   # Install if needed
   winget install Microsoft.PowerShell
   ```

2. **Microsoft Graph PowerShell Modules**
   ```powershell
   Install-Module Microsoft.Graph -Scope CurrentUser
   ```

3. **Exchange Online Management** (for Hide from Address Book)
   ```powershell
   Install-Module ExchangeOnlineManagement -Scope CurrentUser
   ```

4. **Active Directory Module** (for On-Premise UPN updates)
   - Already installed on Domain Controllers
   - Or install RSAT tools on regular servers

5. **Node.js and Playwright** (for AvePoint Fly automation)
   ```powershell
   # Run the setup script once
   .\Setup.ps1
   ```

### 5. Launch the Toolkit

```powershell
.\main-menu.ps1
```

### 6. Auto-Update System

The toolkit automatically checks for updates from GitHub **every time you run it**:

- **Repository:** https://github.com/MoriteUK/AvepointFlyUtility
- **Update Check:** Runs automatically on launch
- **Notification:** Yellow banner appears when updates are available
- **Installation:** Click "Install Update" button
- **Configuration:** Preserved automatically (including Fly API credentials)

**After first run, the server will:**
1. Check GitHub for updates automatically
2. Notify you of available updates
3. Allow one-click installation
4. Never lose your credentials during updates

See [AUTO-UPDATE-SYSTEM.md](AUTO-UPDATE-SYSTEM.md) for details.

## Configuration File Locations

| File | Purpose | Location |
|------|---------|----------|
| `config.json` | Fly API credentials | `%APPDATA%\FlyMigration\config.json` |
| `domains.json` | Domain mappings | `MigrationToolkit\domains.json` |
| `workloads.json` | Workload settings | `MigrationToolkit\workloads.json` |
| `shared-config.json` | Customer info | `MigrationToolkit\shared-config.json` |
| `version.json` | Current version | `MigrationToolkit\version.json` |

## Troubleshooting

### "Script not found" errors
- Ensure all .ps1 files are in the same folder
- Run: `.\Check-Updates.ps1 -Force` to download missing files

### Fly API connection fails
- Verify credentials in Settings > Config
- Click "Test Connection" button
- Check the secret hasn't expired

### Module errors
- Install required PowerShell modules (see Prerequisites)
- Restart PowerShell after installing modules

### Auto-update not working
1. Check internet connectivity to GitHub
2. Verify repository is accessible: https://github.com/MoriteUK/AvepointFlyUtility
3. Run manually: `.\Check-Updates.ps1 -Force`

## Files Excluded from Git (Created Locally)

These files/folders are created during use and should **not** be transferred:
- `logs\` - Log files
- `reports\` - Migration reports  
- `backup-*\` - Update backups
- `auth\storageState.json` - Playwright auth

## Support

Current Version: **2.1.4**  
Repository: https://github.com/MoriteUK/AvepointFlyUtility

For issues, check:
1. Log files in `logs\` folder
2. GitHub repository for latest updates
3. Settings > Config > Check for Updates
4. [AUTO-UPDATE-SYSTEM.md](AUTO-UPDATE-SYSTEM.md) for auto-update troubleshooting

## Deployment Methods Comparison

| Method | Best For | Keeps Updated? |
|--------|----------|----------------|
| **Git Clone** | Development/Admin | ✅ `git pull` |
| **Download ZIP** | One-time setup | ✅ Auto-update |
| **Copy Folder** | Quick transfer | ✅ Auto-update |

All methods support auto-update from GitHub!

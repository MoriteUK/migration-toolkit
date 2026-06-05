# Migration Toolkit - Installation & Setup Guide

## Overview
The Migration Toolkit is an Electron-based desktop application for managing M365 migrations using AvePoint Fly. It provides a modern dashboard interface for monitoring migration status, configuring connections, and running migration-related scripts.

## Prerequisites

### Required Software
1. **PowerShell 7+** (pwsh.exe)
   - Download from: https://github.com/PowerShell/PowerShell/releases
   - Verify installation: `pwsh --version`

2. **Node.js 18+** (includes npm)
   - Download from: https://nodejs.org/
   - Verify installation: `node --version` and `npm --version`

3. **AvePoint Fly.Client PowerShell Module**
   - Install from PowerShell Gallery: `Install-Module -Name Fly.Client`
   - Or contact AvePoint for installation instructions

### Optional (for development)
- Git (for version control)
- Visual Studio Code (recommended editor)

## Installation Steps

### 1. Extract Package
Extract the `MigrationToolkit-Package` folder to a location of your choice (e.g., `C:\MigrationToolkit`)

### 2. Install Dependencies
Open PowerShell (or Command Prompt) and navigate to the `MigrationToolkit-Web` folder:

```powershell
cd C:\MigrationToolkit\MigrationToolkit-Web
npm install
```

This will install all required Node.js packages including Electron.

### 3. Configure AvePoint Fly API
Before running the toolkit, you need to configure your AvePoint Fly API credentials:

1. Launch the application (see "Running the Application" below)
2. Click the **Settings** gear icon in the sidebar
3. Go to the **Config** tab
4. Fill in:
   - **FLY API URL**: `https://graph.avepointonlineservices.com/fly`
   - **AOS CLIENT ID**: Your AvePoint client ID
   - **CLIENT SECRET**: Your AvePoint client secret
   - **SECRET EXPIRY**: Expiry date of your secret
   - **FLY PORTAL URL**: `https://fly.avepointonlineservices.com/#/project/all-list`
   - **SHAREPOINT ADMIN URL**: Your default tenant admin URL
5. Click **Test Connection** to verify
6. Click **Save All**

### 4. Add Customer Information
In Settings > **Customer** tab:
1. Click **Add Customer**
2. Enter:
   - **Prefix**: Short identifier (e.g., "Expretio")
   - **Account Name**: Customer M365 account (e.g., "itvolaris@customer.onmicrosoft.com")
   - **SharePoint Admin URL**: Customer's SharePoint admin URL
3. Click **Save All**

## Running the Application

### Method 1: Using npm (Recommended)
From the `MigrationToolkit-Web` folder:

```powershell
npm start
```

### Method 2: Using Electron directly
```powershell
npx electron .
```

### Method 3: Create Desktop Shortcut
Create a batch file `start-migration-toolkit.bat`:

```batch
@echo off
cd /d C:\MigrationToolkit\MigrationToolkit-Web
npm start
```

Right-click the batch file → Send to → Desktop (create shortcut)

## Directory Structure

```
MigrationToolkit-Package/
├── MigrationToolkit-Web/          # Electron app
│   ├── index.html                 # Main UI
│   ├── main.js                    # Electron main process
│   ├── renderer.js                # UI logic
│   ├── preload.js                 # Security bridge
│   ├── package.json               # Dependencies
│   ├── src/
│   │   └── styles/
│   │       └── main.css          # Styles
│   └── public/
│       └── icon.ico              # App icon
│
└── VGMigrations/                  # PowerShell scripts
    ├── Get-MigrationData.ps1     # Fetch migration status
    ├── Test-FlyConnection.ps1    # Test API connection
    ├── Save-Config.ps1            # Save encrypted config
    ├── Encrypt-Secret.ps1         # Encrypt secrets
    ├── lib.ps1                    # Shared functions
    └── [other scripts...]
```

## Configuration Files

The application stores configuration in:
- **Windows**: `%APPDATA%\FlyMigration\config.json`

This file contains:
- API credentials (encrypted client secret)
- Customer information
- Tenant URLs

**Important**: This config file contains encrypted secrets. Do not share it publicly.

## Features

### Dashboard
- Real-time migration status overview
- View items in scope, completed, in progress, and needing attention
- Progress by workload (SharePoint, Exchange, OneDrive, Teams, Teams Chat)
- Click stat boxes for detailed information

### Settings
- Configure AvePoint Fly API credentials
- Test connection to validate credentials
- Manage customer/tenant information
- Check for software updates

### Tools
- **Discovery**: Run M365 discovery scripts
- **AvePoint Fly**: 
  - App Registration setup
  - AOS tenant configuration
  - Connection and mapping management
  - Migration reports
  - Real-time migration monitor
- **Misc Scripts**: OneDrive provisioning, Teams migration helpers
- **Domain Removal**: Guided workflows for domain cleanup

## Troubleshooting

### Application won't start
- Ensure Node.js 18+ is installed: `node --version`
- Reinstall dependencies: `npm install`
- Check for errors in terminal output

### "Fly.Client module not found"
- Install the module: `Install-Module -Name Fly.Client`
- Verify: `Get-Module -Name Fly.Client -ListAvailable`

### "Failed to connect to Fly API"
- Verify API credentials in Settings
- Click "Test Connection" to diagnose
- Check that Client Secret hasn't expired
- Ensure network connectivity

### "Client secret not being saved"
- Make sure to click "Save All" after entering the secret
- The field will show dots (••••••••) when a secret is configured
- Check `%APPDATA%\FlyMigration\config.json` exists

### Dashboard shows no data
- Ensure Fly.Client module is installed
- Configure API credentials in Settings
- Add customers in Settings > Customer tab
- Projects must exist in AvePoint Fly with naming: `{Prefix} - {Workload}`
  - Example: "Expretio - Exchange", "Expretio - SharePoint"

### Layout/UI issues
- Press `Ctrl+Shift+R` to hard refresh
- Restart the application
- Clear Electron cache and restart

## Building for Distribution (Optional)

To create an executable installer:

1. Install electron-builder:
```powershell
npm install --save-dev electron-builder
```

2. Add to package.json:
```json
"scripts": {
  "build": "electron-builder"
},
"build": {
  "appId": "com.volaris.migrationtoolkit",
  "productName": "Migration Toolkit",
  "directories": {
    "output": "dist"
  },
  "win": {
    "target": "nsis",
    "icon": "public/icon.ico"
  }
}
```

3. Build:
```powershell
npm run build
```

The installer will be created in the `dist/` folder.

## Support & Issues

For issues or feature requests:
1. Check the console logs (Ctrl+Shift+I in the app)
2. Check PowerShell script logs in `VGMigrations/logs/`
3. Verify prerequisites are installed correctly

## Security Notes

- Client secrets are encrypted using Windows DPAPI
- Secrets are user-specific and machine-specific
- Config files should not be shared between users/machines
- Always use HTTPS for API connections

## Version Information

- Electron App Version: Check in Settings > Config
- PowerShell Scripts: See `VGMigrations/version.json`

## License

© 2024 Volaris Group. All rights reserved.

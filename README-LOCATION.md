# Migration Toolkit - Latest Working Version

**Location**: `C:\toolkit`  
**Version**: 2.9.34  
**Date**: 2026-07-28

## What's Here

This is the **latest consolidated version** of the Migration Toolkit, containing:

### PowerShell Scripts (VGMigrations folder)
- 73 PowerShell scripts including all utilities
- **NEW**: Check-TenantLicenses.ps1 - License report with auto-app registration
- **NEW**: Check-DomainMigrationReadiness.ps1 - Domain verification checker
- All discovery, AvePoint Fly, and domain removal tools

### Web Application (MigrationToolkit-Web folder)
- Electron-based web UI (version 2.9.34)
- **NEW**: Discovery section with License Report and Domain Readiness buttons
- Modern interface with sidebar navigation
- Dashboard with live migration stats

## How to Use

### PowerShell Version
```powershell
cd C:\toolkit\VGMigrations
.\main-menu.ps1
```

### Web Version
```powershell
cd C:\toolkit\MigrationToolkit-Web
npm start
```

## Key Features in v2.9.34

1. **License Report** - Auto-creates app registration per tenant, counts mailbox-granting licenses
2. **Domain Readiness** - Checks if domains exist in target tenant and DNS verification status
3. **Web UI** - Discovery tools accessible via buttons in the Discovery section

## Git Repository

The main git repository is at: `C:\migration-toolkit-main`
- GitHub: https://github.com/MoriteUK/migration-toolkit
- This location (C:\toolkit) is a working copy synced from the git repo

## Other Locations Cleaned Up

✅ Removed: `C:\Users\Andy White\Downloads\migration-toolkit-main`  
✅ Removed: `C:\Users\Andy White\Downloads\migration-toolkit-main (2)`  
✅ Kept: `C:\migration-toolkit-main` (git repository for version control)

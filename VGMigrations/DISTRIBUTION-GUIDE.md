# Distribution Guide - Migration Tools

This guide explains how to package and distribute the Migration Tools to customers.

## Distribution Strategies

You have **two options** for distributing the toolkit to customers:

### Option 1: GitHub Auto-Update (Recommended for Internal/Tech-Savvy Users)

**Advantages:**
- Always up-to-date automatically
- No manual packaging needed
- Customers get bug fixes and new features instantly
- Smaller initial download

**How to Deploy:**

1. **Give customers the GitHub repository link:**
   ```
   https://github.com/MoriteUK/AvepointFlyUtility
   ```

2. **Customers clone or download:**
   ```powershell
   # Option A: Clone with git
   git clone https://github.com/MoriteUK/AvepointFlyUtility.git MigrationToolkit
   
   # Option B: Download ZIP
   # Download from GitHub → Extract to desired location
   ```

3. **Run setup:**
   ```powershell
   cd MigrationToolkit
   .\Setup.ps1
   ```

4. **Configure credentials:**
   - Run `.\main-menu.ps1`
   - Click ⚙ Settings
   - Enter Fly API credentials
   - Enter customer prefix

5. **Auto-updates work automatically:**
   - Toolkit checks GitHub on every launch
   - Yellow banner appears when updates available
   - One-click to install

**Best for:**
- Internal team deployments
- Tech-savvy customers
- Long-term engagements where you want customers to get updates

---

### Option 2: Standalone Package (Best for Customer Distribution)

**Advantages:**
- Self-contained .exe launcher
- No git required
- Looks more professional
- Can include pre-configured settings
- Works offline (after initial setup)

**How to Package:**

#### Step 1: Create the Package

```powershell
# In your MigrationToolkit directory
.\New-Package.ps1
```

This creates `MigrationToolkit.zip` containing:
- All scripts and tools
- Pre-built `MigrationTools.exe` launcher
- Documentation
- Node modules (for Playwright)

#### Step 2: Customize for Customer (Optional)

Before packaging, you can pre-configure:

1. **Add customer prefix:**
   Edit `shared-config.json`:
   ```json
   {
     "Customers": [
       { "Prefix": "CUSTOMER", "AccountName": "", "SharePointAdminUrl": "" }
     ]
   }
   ```

2. **Pre-configure workloads:**
   Edit `workloads.json` with default policies

3. **Include Fly credentials** (if same for all customers):
   - Export from `$env:APPDATA\FlyMigration\config.json`
   - Include as `fly-config.json` in package
   - Instruct customer to run installation script

#### Step 3: Distribute to Customer

**Send them:**
1. `MigrationToolkit.zip`
2. Installation instructions (see below)

---

## Customer Installation Instructions

### For Standalone Package Distribution:

Create a simple instruction document for customers:

```
Migration Tools - Installation Guide

Prerequisites:
- Windows 10/11 or Windows Server 2019+
- PowerShell 7.0 or later
- Internet connection (for initial setup)

Installation Steps:

1. Extract MigrationToolkit.zip to:
   C:\MigrationToolkit\

2. Right-click PowerShell 7 and "Run as Administrator"

3. Navigate to the folder:
   cd C:\MigrationToolkit

4. Run the setup:
   .\Setup.ps1

5. Launch the toolkit:
   - Double-click: MigrationTools.exe
   - OR run: .\main-menu.ps1

6. Configure your credentials:
   - Click the ⚙ Settings icon
   - Enter your Fly API URL and credentials
   - Save

7. You're ready to go!

Support:
- Check logs in the logs\ folder for errors
- README.md contains detailed documentation
```

---

## Hybrid Approach (Best of Both Worlds)

**Recommended for most customers:**

1. **Initial deployment:** Send them the standalone package
   - Professional .exe launcher
   - Pre-configured for their environment
   - Works immediately

2. **Enable auto-updates:** Already built-in!
   - The package includes `Check-Updates.ps1`
   - Auto-update checks GitHub on every launch
   - Customers get updates without you doing anything

3. **Update process:**
   - You push fixes to GitHub
   - Customer launches toolkit
   - Yellow banner: "Update available"
   - Customer clicks "Install Update"
   - Done!

This way:
- ✅ Professional initial installation
- ✅ Automatic updates from GitHub
- ✅ No re-packaging needed
- ✅ Customers always have latest version

---

## Building Custom Packages

### For Different Customer Environments:

If you need customer-specific packages (different Fly instances, branding, etc.):

1. **Create customer branch:**
   ```bash
   git checkout -b customer-acme
   ```

2. **Customize:**
   - Update branding (icons, titles)
   - Pre-configure customer settings
   - Set their Fly API endpoint

3. **Package:**
   ```powershell
   .\New-Package.ps1
   ```

4. **Distribute:**
   Send `MigrationToolkit.zip` to customer

5. **Updates:**
   - Merge main branch updates into customer branch
   - Push to GitHub
   - Customer gets auto-updates from their branch

---

## Quick Reference

| Distribution Method | Use Case | Auto-Updates | Setup Complexity |
|---------------------|----------|--------------|------------------|
| **GitHub Clone** | Internal team | ✅ Yes | Low |
| **Standalone Package** | External customers | ✅ Yes | Very Low |
| **Custom Package** | Multi-tenant customers | ✅ Yes (from custom branch) | Medium |

---

## Files Required for Distribution

### Minimum Files (GitHub approach):
Just share the GitHub URL - they get everything via clone/download.

### Standalone Package Includes:
- All `.ps1` scripts
- `MigrationTools.exe` launcher
- `version.json` (for auto-update)
- `node_modules/` (Playwright)
- Documentation (README.md, etc.)
- Icon files
- Configuration templates

### NOT Included in Packages:
- User data (`domains.json`, `workloads.json` - customer creates)
- Logs and reports
- Credentials (`config.json` - customer enters via Settings)
- `.git` folder
- Backup folders

---

## Updating the Master Package

When you make changes and want to distribute to new customers:

1. **Test locally:**
   ```powershell
   .\main-menu.ps1
   # Test all features
   ```

2. **Commit to GitHub:**
   ```bash
   git add .
   git commit -m "Your changes"
   git push origin main
   ```

3. **Existing customers:** Auto-update picks it up automatically

4. **New customers:** 
   ```powershell
   .\New-Package.ps1
   ```
   Send them the new `MigrationToolkit.zip`

---

## Support Model

### For GitHub-Distributed Customers:
- They pull updates automatically
- You just push fixes to GitHub
- Zero deployment effort

### For Package-Distributed Customers:
- They still get auto-updates from GitHub
- No need to send new packages
- Only send new package if they need to re-install

---

## Current Version

Check `version.json` for current version number.

Latest: **2.1.11** (2026-06-02)

GitHub Repository: https://github.com/MoriteUK/AvepointFlyUtility

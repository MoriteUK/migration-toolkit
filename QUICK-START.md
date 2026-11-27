# Migration Toolkit - Quick Start Guide

## 🚀 First Time Setup

### 1. Install Required Modules

Run the module installation checker:

```powershell
# From the migration-toolkit-main directory
.\Install-RequiredModules.ps1
```

**What it does:**
- ✅ Checks PowerShell version (7+ recommended)
- ✅ Verifies NuGet package provider
- ✅ Checks PowerShell Gallery configuration
- ✅ Scans all required modules
- ✅ Offers to install missing modules
- ✅ Checks Node.js and browser for FLY reports

**Auto-install mode:**
```powershell
.\Install-RequiredModules.ps1 -AutoInstall -Scope CurrentUser
```

**Admin installation (system-wide):**
```powershell
# Run PowerShell as Administrator
.\Install-RequiredModules.ps1 -Scope AllUsers
```

### 2. Required Modules

**Essential Modules:**
- ✅ Microsoft.Graph (Azure AD, Users, Groups)
- ✅ Microsoft.Online.SharePoint.PowerShell
- ✅ ExchangeOnlineManagement
- ✅ MicrosoftTeams
- ✅ Fly.Client (AvePoint FLY)
- ✅ ImportExcel

**Optional but Recommended:**
- AzureAD (legacy support)
- MSOnline (legacy support)
- PSWriteColor (enhanced output)

### 3. Additional Requirements

**Node.js** (for FLY Reports)
- Download: https://nodejs.org/
- Required for FLY Executive Brief PDF generation

**Chrome or Edge** (for PDF generation)
- Microsoft Edge (recommended): Included with Windows
- Google Chrome: https://www.google.com/chrome/

## 📊 Using FLY Executive Brief

### Access the Feature

1. Launch Migration Toolkit
2. Click **Reports** in sidebar (📊 icon)
3. Select **FLY Executive Brief**

### Generate Your First Brief

**Quick Steps:**
1. **Select Client**: Choose domain from dropdown
2. **Browse Export**: Select FLY export folder
3. **Click Generate**: Creates HTML + PDF report

**Detailed Steps:**

```
┌─ Client Selection ──────────────────┐
│ Domain: Select from dropdown        │
│ Workload: Auto-detect (or specify)  │
└─────────────────────────────────────┘

┌─ Export Data ───────────────────────┐
│ Export Folder: Browse to FLY export │
│   Must contain: *Mappings.csv       │
│   Optional: *Site security.csv      │
│   Optional: *SharedLink.csv         │
└─────────────────────────────────────┘

┌─ Options ───────────────────────────┐
│ ✅ Generate PDF                     │
│ ✅ Auto-open PDF                    │
│ 📧 Email (optional)                 │
└─────────────────────────────────────┘

[📄 Generate Brief]
```

### What You Get

**Files Created:**
1. `[Client]_[Workload]_Executive_Brief.html`
2. `[Client]_[Workload]_Executive_Brief.pdf`
3. `[Client]_[Workload]_Executive_Brief.summary.json`

**Report Contents:**
- 📈 Executive verdict
- 📊 4 key metrics
- 🏢 Delivery status (CIO)
- 🔒 Security analysis (CISO)
- ✅ Decision requests

## 🔧 Troubleshooting

### Module Not Found

**Error:**
```
Module 'Microsoft.Online.SharePoint.PowerShell' was not loaded
```

**Fix:**
```powershell
.\Install-RequiredModules.ps1
```

### FLY Brief Script Not Found

**Error:**
```
FLY brief script not found at: C:\migration-toolkit-main\FLY\Template\fly-brief.js
```

**Fix:**
Verify files exist:
```powershell
dir C:\migration-toolkit-main\FLY\Template\
```

Should show:
- ✅ FLY_Executive_Brief.html
- ✅ fly-brief.js
- ✅ Invoke-FlyBriefs.ps1
- ✅ README.md

### Node.js Not Found

**Error:**
```
'node' is not recognized as an internal or external command
```

**Fix:**
1. Install Node.js: https://nodejs.org/
2. Restart Migration Toolkit
3. Verify: `node --version`

### PDF Not Generated

**Possible Causes:**
- Chrome/Edge not installed
- Edge path not found

**Fix:**
1. Install Microsoft Edge (included with Windows 11)
2. Or install Google Chrome
3. Set EDGE_PATH environment variable if needed:
   ```powershell
   $env:EDGE_PATH = "C:\Program Files\Microsoft\Edge\Application\msedge.exe"
   ```

### Export Folder Error

**Error:**
```
No *Mappings.csv found
```

**Fix:**
Ensure export folder contains FLY CSV exports:
- Required: `SomeName_Mappings.csv`
- Optional: `SomeName_Site security.csv`
- Optional: `SomeName_SharedLink.csv`

## 📁 Directory Structure

```
migration-toolkit-main/
├── MigrationToolkit-Web/      # Main application
│   ├── index.html             # UI with Reports section
│   ├── renderer.js            # Frontend logic
│   ├── main.js                # Backend handlers
│   └── preload.js             # IPC bridge
│
├── FLY/
│   └── Template/              # FLY Brief templates
│       ├── FLY_Executive_Brief.html
│       ├── fly-brief.js       # Generation script
│       ├── Invoke-FlyBriefs.ps1
│       ├── Test-FlyPrereqs.ps1
│       └── README.md
│
├── VGMigrations/              # PowerShell scripts
│
└── Install-RequiredModules.ps1  # Module installer
```

## 🎯 Common Workflows

### 1. First Time Setup
```powershell
# 1. Install modules
.\Install-RequiredModules.ps1 -AutoInstall

# 2. Launch toolkit
.\MigrationToolkit-Web\START.cmd

# 3. Navigate to Reports > FLY Executive Brief
```

### 2. Generate Report
```
1. Reports > FLY Executive Brief
2. Select domain
3. Browse to export folder
4. Click Generate Brief
5. Review PDF
```

### 3. Email Report
```
1. ✅ Check "Email PDF report"
2. Enter recipient emails
3. Customize subject/body
4. Generate Brief
```
*(Note: Email requires additional setup - see INTEGRATION-SUMMARY.md)*

### 4. Batch Processing
Use PowerShell automation:
```powershell
cd C:\migration-toolkit-main\FLY\Template

# Single domain
.\Invoke-FlyBriefs.ps1 Contoso

# All domains
.\Invoke-FlyBriefs.ps1

# Dry run
.\Invoke-FlyBriefs.ps1 -ListOnly
```

## 📚 Documentation

| Document | Purpose |
|----------|---------|
| **INTEGRATION-SUMMARY.md** | Complete integration details |
| **FLY/Template/README.md** | FLY Brief standalone usage |
| **QUICK-START.md** | This guide |

## 🆘 Getting Help

### Module Issues
1. Run `.\Install-RequiredModules.ps1`
2. Check output for errors
3. Review `module-check-results.txt`

### FLY Brief Issues
1. Check Generation Log in toolkit
2. Review `FLY/Template/README.md`
3. Run `.\FLY\Template\Test-FlyPrereqs.ps1`

### General Issues
1. Check Migration Toolkit logs
2. Verify PowerShell version (7+ recommended)
3. Ensure running latest version

## 🔄 Updates

### Update Modules
```powershell
# Update all modules
Update-Module

# Update specific module
Update-Module -Name Microsoft.Graph
```

### Update Toolkit
```powershell
# Pull latest from repository
git pull

# Or download latest release
```

## ⚙️ Advanced Configuration

### Custom Thresholds
Edit `FLY/Template/FLY_Executive_Brief.html`:
```javascript
// Line 97-103
var CFG = {
  contentGood: 99.5,   // Green threshold
  contentWarn: 97,     // Amber threshold
  permGood: 95,        // Permission threshold
  anyoneRed: 1000,     // Anonymous link risk
  staleDays: 14        // Stalled project days
};
```

### Email Integration
See `INTEGRATION-SUMMARY.md` section "Email Functionality"

### Custom Report Templates
1. Copy `FLY_Executive_Brief.html`
2. Modify styling/content
3. Update `fly-brief.js` TEMPLATE path

## 📞 Support Resources

- **Module Installation**: Run `.\Install-RequiredModules.ps1`
- **FLY Prerequisites**: Run `.\FLY\Template\Test-FlyPrereqs.ps1`
- **Documentation**: See `INTEGRATION-SUMMARY.md`
- **AvePoint FLY**: Contact AvePoint support

## ✅ Pre-Flight Checklist

Before generating your first report:

- [ ] PowerShell 7+ installed
- [ ] All modules installed (`.\Install-RequiredModules.ps1`)
- [ ] Node.js installed (`node --version`)
- [ ] Chrome or Edge installed
- [ ] FLY export folder ready
- [ ] Domain selected in toolkit
- [ ] Migration Toolkit launched

## 🎉 Success Indicators

Your setup is complete when:
- ✅ All modules show green in module check
- ✅ Reports section appears in sidebar
- ✅ Domain dropdown populates
- ✅ FLY brief generates successfully
- ✅ PDF opens automatically
- ✅ No errors in Generation Log

---

**Version**: 1.0  
**Last Updated**: November 2026  
**Integration**: Migration Toolkit + FLY Executive Brief

# FLY Executive Brief - Automated Generation Guide

## 🚀 What's New: Automatic Report Download

The FLY Executive Brief now **automatically connects to AvePoint FLY**, downloads the latest migration reports, and generates professional PDF briefs—all in one click!

**No more manual export steps!** ✨

---

## 📋 How It Works

```
┌──────────────────────────────────────────────┐
│  1. Enter Client Name: "Contoso"            │
├──────────────────────────────────────────────┤
│  2. Click "Generate Brief"                   │
├──────────────────────────────────────────────┤
│  3. Script automatically:                    │
│     ✅ Connects to FLY API                  │
│     ✅ Finds all Contoso projects           │
│     ✅ Downloads latest exports             │
│     ✅ Generates HTML + PDF briefs          │
│     ✅ Opens PDFs for review                │
└──────────────────────────────────────────────┘
```

**Previous Workflow (5+ steps):**
1. ❌ Login to FLY web portal
2. ❌ Find each project manually
3. ❌ Click Export → Wait → Download
4. ❌ Extract ZIP files
5. ❌ Browse to folder in toolkit
6. ✅ Generate brief

**New Workflow (1 step):**
1. ✅ Type client name → Click Generate → Done! 🎉

---

## 🎯 Quick Start

### Step 1: Ensure FLY is Configured

The toolkit uses the same FLY API credentials from Settings.

**Check Settings:**
1. Click ⚙️ **Settings** in toolkit
2. Verify **FLY API URL** (e.g., `https://graph-uk.avepointonlineservices.com/fly`)
3. Verify **Client ID** and **Client Secret**
4. Click **Test Connection** to verify

### Step 2: Generate Briefs

1. Click **📊 Reports** in sidebar
2. Select **FLY Executive Brief**
3. Enter **Client Name** (e.g., "Contoso")
4. *Optional:* Select specific workloads (or leave unchecked for all)
5. Click **📄 Generate Brief**

**That's it!** The system will:
- Find all FLY projects for that client
- Download the latest migration data
- Generate a professional brief for each workload
- Auto-open the PDFs

---

## 🔧 Configuration

### FLY API Settings

Located in: **Settings → Config Tab**

| Field | Example | Description |
|-------|---------|-------------|
| **FLY API URL** | `https://graph-uk.avepointonlineservices.com/fly` | Your data center endpoint |
| **Client ID** | `9ebef03c-...` | From AvePoint Online Services |
| **Client Secret** | `xxxxx` | Your app registration secret |

**Get Credentials:**
1. Login to AvePoint Online Services
2. Go to **Administration → App registrations**
3. Create or use existing app with FLY permissions
4. Copy Client ID and Secret

### Project Configuration

The automation uses `fly-config.json` located at:
```
C:\migration-toolkit-main\FLY\Template\fly-config.json
```

**Auto-Discovery Mode** (Recommended):
```json
{
  "discover": {
    "enabled": true,
    "workloads": ["OneDrive", "SharePoint", "Exchange", "Teams", "M365Group", "TeamChat"],
    "exclude": ["TestProject*"]
  }
}
```

This automatically finds all FLY projects and matches them by client name!

---

## 📝 Usage Examples

### Example 1: All Workloads for a Client

```
Client Name: Contoso
Workloads: [All unchecked = include all]
↓
Generates briefs for:
✅ Contoso - OneDrive
✅ Contoso - SharePoint  
✅ Contoso - Exchange
✅ Contoso - Teams
```

### Example 2: Specific Workloads Only

```
Client Name: AEP Italia
Workloads: ✅ OneDrive  ✅ SharePoint
↓
Generates briefs for:
✅ AEP Italia - OneDrive
✅ AEP Italia - SharePoint
❌ AEP Italia - Exchange (skipped)
```

### Example 3: Custom Output Location

```
Client Name: Fabrikam
Output Folder: D:\Reports\Migration
↓
Saves briefs to:
D:\Reports\Migration\Fabrikam\{Date}\
```

---

## 📊 Output Structure

**Default Output Location:**
```
C:\Temp\Scripts\Avepoint\FLY\Reports\
├── Contoso\
│   ├── Contoso_-_OneDrive\
│   │   └── 20261127-1430\
│   │       ├── Contoso_OneDrive_Executive_Brief.html
│   │       ├── Contoso_OneDrive_Executive_Brief.pdf ⭐
│   │       ├── Contoso_OneDrive_Executive_Brief.summary.json
│   │       └── [FLY CSV exports]
│   │
│   └── Contoso_-_SharePoint\
│       └── 20261127-1432\
│           ├── Contoso_SharePoint_Executive_Brief.html
│           ├── Contoso_SharePoint_Executive_Brief.pdf ⭐
│           └── ...
```

**What You Get:**
- 📄 **HTML** - Interactive report (editable)
- 📄 **PDF** - Print-ready executive brief
- 📄 **JSON** - Summary data for analysis
- 📄 **CSV Files** - Original FLY exports (for reference)

---

## 🎨 UI Features

### New Simplified Interface

**Client Selection:**
- Dropdown auto-populated from discovery
- Or type any client name manually
- Tooltip shows how matching works

**Workload Selection:**
- Checkboxes for each workload type
- Leave all unchecked = include all
- Useful for testing or specific reports

**Progress Tracking:**
- Real-time log shows each step
- Lists all projects being processed
- Shows export and generation progress
- Displays final file paths

**Results Preview:**
- Summary statistics for each workload
- Quick overview of migration status
- Links to open PDFs or folders

### Removed Fields

**Gone (automated now):**
- ~~Export Folder Browse~~ - Auto-downloaded
- ~~Internal Domains~~ - Stored per FLY project
- ~~Workload Dropdown~~ - Now multi-select checkboxes

**Simplified from 7 fields to 3!**

---

## 🔍 How Projects Are Found

The system searches FLY for projects matching your client name:

### Pattern Matching

**Standard Format:**
```
{Client} - {Workload}
```

Examples:
- "Contoso - OneDrive"
- "Contoso - SharePoint"
- "AEP Italia - Exchange"

**The script matches:**
- Exact prefix: "Contoso" finds all "Contoso - *" projects
- Case insensitive
- Filters by workload if specified

### Auto-Discovery

Enabled in `fly-config.json`:
```json
{
  "discover": {
    "enabled": true,
    "workloads": ["OneDrive", "SharePoint", ...],
    "exclude": ["TestProject*", "Sandbox*"]
  }
}
```

**What it does:**
1. Queries all FLY projects
2. Filters by workload type
3. Excludes test projects (via patterns)
4. Matches client name from project name

---

## 🚨 Troubleshooting

### "No FLY projects found for this client"

**Cause:** No matching projects in FLY

**Solutions:**
1. Check exact client name spelling
2. Verify projects exist in FLY portal
3. Check project naming format (`{Client} - {Workload}`)
4. Try without workload filters (include all)

### "FLY API connection failed"

**Cause:** Invalid credentials or network issue

**Solutions:**
1. Settings → Test Connection
2. Verify Client ID/Secret
3. Check FLY API URL (correct region?)
4. Verify app has FLY permissions
5. Check network/firewall

### "PowerShell script failed"

**Cause:** Fly.Client module not installed

**Solutions:**
```powershell
# Run module installer
.\Install-RequiredModules.ps1

# Or install manually
Install-Module -Name Fly.Client -Scope CurrentUser
```

### "No briefs were generated"

**Causes & Solutions:**

| Issue | Solution |
|-------|----------|
| No matching projects | Check client name spelling |
| Projects have no data | Wait for migration to run |
| Export failed | Check FLY API permissions |
| Node.js not found | Install Node.js |
| Browser not found | Install Edge/Chrome |

**Check Generation Log** for specific error messages!

---

## 📧 Email Functionality

Email options are available but require setup (see original INTEGRATION-SUMMARY.md).

**Current Status:**
- ✅ UI controls present
- ⚠️ Backend integration required

**To Enable:**
1. Choose email service (SendGrid, Mailgun, Graph API)
2. Update `main.js` with API integration
3. Install required npm packages

**Workaround:**
- Generate PDFs
- Manually attach to email
- Or share via OneDrive/SharePoint link

---

## ⚙️ Advanced Configuration

### Custom Report Thresholds

Edit `C:\migration-toolkit-main\FLY\Template\FLY_Executive_Brief.html`:

```javascript
// Lines 97-103
var CFG = {
  contentGood: 99.5,   // Green if ≥99.5% migrated
  contentWarn: 97,     // Amber if ≥97%
  permGood: 95,        // Green if ≥95% permissions
  permWarn: 70,        // Amber if ≥70%
  anyoneRed: 1000,     // High risk if >1000 anon links
  staleDays: 14        // Stalled if no activity for 14 days
};
```

### Filter Projects

Edit `fly-config.json`:

```json
{
  "discover": {
    "exclude": [
      "TestProject*",
      "Sandbox*",
      "Demo*"
    ]
  }
}
```

### Override Project Settings

```json
{
  "projects": [
    {
      "project": "AEP Italia - OneDrive",
      "client": "AEP Italia",
      "workload": "OneDrive",
      "internalDomains": ["aep-italia.it", "easyticketing.it"]
    }
  ]
}
```

This overrides auto-discovery for specific projects.

---

## 🔄 Comparison: Old vs New

| Feature | Manual Export | Automated 🌟 |
|---------|---------------|--------------|
| **Login Required** | ✅ Yes | ❌ No |
| **Find Projects** | Manual | Auto |
| **Download Exports** | Manual per project | Auto all projects |
| **Extract ZIPs** | Manual | Auto |
| **Browse Folders** | Manual | Auto |
| **Multi-Workload** | Repeat 5x for each | One click |
| **Time Required** | ~5-10 min per project | ~1 min total |
| **Steps** | 6+ per workload | 1 total |
| **Errors** | Easy to miss files | Validated |

---

## 📚 Related Documentation

| Document | Purpose |
|----------|---------|
| **QUICK-START.md** | General toolkit guide |
| **INTEGRATION-SUMMARY.md** | Technical implementation details |
| **AUTOMATED-BRIEF-GUIDE.md** | This document |
| **FLY/Template/README.md** | Standalone CLI usage |

---

## 🎓 Best Practices

### 1. **Regular Schedule**
Run briefs weekly or at migration milestones:
```
Week 1: Generate baseline briefs
Week 2: Track progress
Week 3: Review before go-live
```

### 2. **Naming Convention**
Keep FLY projects named consistently:
```
✅ ClientName - Workload
✅ Contoso - OneDrive
❌ onedrive_contoso
❌ Contoso OneDrive Migration
```

### 3. **Output Organization**
Use default output structure or organize by date:
```
Reports/
  2026-Q4/
    Contoso/
  2027-Q1/
    AEP/
```

### 4. **Review Workflow**
1. Generate all briefs
2. Review PDFs for red/amber status
3. Address issues
4. Re-generate updated briefs
5. Present to stakeholders

---

## ✅ Success Checklist

Before presenting briefs:

- [ ] All FLY projects have recent data
- [ ] API credentials tested and working
- [ ] Node.js and browser installed
- [ ] All workloads processed successfully
- [ ] No red flags in Generation Log
- [ ] PDFs open correctly
- [ ] Statistics look accurate
- [ ] Ready to share/email

---

## 🆘 Support

**Common Issues:**
1. Module errors → Run `Install-RequiredModules.ps1`
2. API errors → Check Settings → Test Connection
3. No briefs → Check Generation Log for details
4. PDF issues → Verify Edge/Chrome installed

**Still Stuck?**
1. Check Generation Log (detailed errors)
2. Review fly-config.json (project discovery)
3. Test FLY connection manually
4. Verify Fly.Client module installed

---

## 🎉 Summary

**The new automated workflow:**
1. ✅ One-click brief generation
2. ✅ No manual FLY portal login
3. ✅ Auto-downloads latest data
4. ✅ Processes multiple workloads
5. ✅ Professional PDF output
6. ✅ Real-time progress tracking

**Time saved:** ~90% reduction in manual steps! 🚀

---

**Version**: 2.0 (Automated)  
**Last Updated**: November 2026  
**Feature**: Auto-download from FLY API

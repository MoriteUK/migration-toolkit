# FLY Executive Brief - Integration Summary

## Overview
The FLY Executive Brief generator has been successfully integrated into the Migration Toolkit Web application with a new **Reports** section in the sidebar.

## What Was Added

### 1. **User Interface (index.html)**

#### Sidebar Section
- Added new "Reports" expandable section with 📊 icon
- Positioned between "Post Migration" and "Misc Scripts"
- Contains "FLY Executive Brief" submenu item

#### New View: FLY Executive Brief
Located at the bottom of the main content area, includes:

**Input Fields:**
- **Client/Migration Name**: Domain dropdown (auto-populated from discovery CSV)
- **Workload Type**: Dropdown (OneDrive, SharePoint, Exchange, Teams, etc.) with auto-detect option
- **FLY Export Folder**: Browse button to select folder containing CSV exports
- **Internal Email Domains**: Optional comma-separated list (auto-detects if blank)
- **Output Folder**: Optional (defaults to export folder)

**Output Options:**
- ✅ Generate PDF (checked by default)
- ✅ Auto-open PDF when complete
- 📧 Email PDF report (with expandable email options)
  - To: Email addresses (comma-separated)
  - Subject: Customizable subject line
  - Body: Optional email message

**Action Buttons:**
- 📄 Generate Brief (primary action)
- ⬇️ Download PDF (appears after generation)
- 📂 Open Output Folder (appears after generation)

**Output Sections:**
- Generation Log: Real-time progress and results
- Preview: Summary statistics displayed after generation

### 2. **Frontend Logic (renderer.js)**

**View Management:**
- Added `'reports-fly-brief': 'flyBriefView'` to viewMap
- Added `loadFlyBriefDomains()` initialization on view switch

**Domain Dropdown:**
- `loadFlyBriefDomains()`: Loads domains from discovery CSV
- Auto-completes with existing domain data
- Click-to-select dropdown interface

**Email Options:**
- Toggle checkbox shows/hides email configuration fields
- Pre-fills subject line with client name

**Brief Generation Handler:**
- Validates all required fields
- Shows generation log in real-time
- Displays summary statistics after completion
- Shows preview with key metrics
- Enables download/open buttons on success
- Handles errors gracefully

### 3. **Backend (main.js)**

**New IPC Handler: `generate-fly-brief`**
- Validates export folder exists
- Locates fly-brief.js script at `migration-toolkit-main/FLY/Template/`
- Builds node command with appropriate arguments
- Spawns node process to run fly-brief.js
- Captures stdout/stderr for logging
- Parses output to extract HTML/PDF paths
- Reads summary.json for statistics
- Auto-opens PDF if requested
- Placeholder for email functionality (requires email service integration)

**Helper Handler: `open-path`**
- Opens files or folders using system default application

### 4. **IPC Bridge (preload.js)**

**Exposed API Methods:**
- `generateFlyBrief(options)`: Generates the executive brief
- `openPath(filePath)`: Opens files/folders
- `getDiscoveryDomains()`: Gets domains from shared config

### 5. **Template Files**

**Location:** `C:\migration-toolkit-main\FLY\Template\`

**Files Copied:**
- ✅ FLY_Executive_Brief.html (core template)
- ✅ fly-brief.js (Node.js generator script)
- ✅ Invoke-FlyBriefs.ps1 (PowerShell orchestrator)
- ✅ Test-FlyPrereqs.ps1 (prerequisite checker)
- ✅ fly-config.json (configuration example)
- ✅ fly-config.example.json
- ✅ README.md

## How to Use

### 1. **Access the Feature**
1. Launch Migration Toolkit
2. Click **Reports** in the sidebar (📊 icon)
3. Click **FLY Executive Brief**

### 2. **Generate a Brief**
1. **Select Client**: Choose from dropdown or type a domain name
2. **Choose Workload** (optional): Select the migration type or leave as "Auto-detect"
3. **Browse Export Folder**: Click Browse and select the folder containing:
   - `*Mappings.csv` (required)
   - `*Site security.csv` (optional, for security section)
   - `*SharedLink.csv` (optional, for sharing analysis)
4. **Internal Domains** (optional): Enter comma-separated domains or leave blank
5. **Output Folder** (optional): Choose where to save, or leave blank to use export folder

### 3. **Configure Output**
- ✅ Check "Generate PDF" (recommended)
- ✅ Check "Auto-open PDF" to view immediately
- For email delivery:
  - ✅ Check "Email PDF report"
  - Enter recipient email(s)
  - Customize subject/body (optional)

### 4. **Generate**
Click **📄 Generate Brief**

**What Happens:**
1. Progress appears in "Generation Log"
2. Node.js runs fly-brief.js script
3. HTML brief is created
4. PDF is rendered (if enabled)
5. Summary statistics appear in preview
6. Download/Open buttons become available

### 5. **View Results**
- Click **⬇️ Download PDF** to open the PDF
- Click **📂 Open Output Folder** to view all files
- Review summary statistics in the preview panel

## Output Files

Generated in the output folder (or export folder):

1. **[Client]_[Workload]_Executive_Brief.html**
   - Self-contained HTML report
   - Can be opened in any browser
   - Text is editable (click to edit)

2. **[Client]_[Workload]_Executive_Brief.pdf**
   - PDF version for distribution
   - A4 format, print-ready
   - Generated via headless Chrome/Edge

3. **[Client]_[Workload]_Executive_Brief.summary.json**
   - JSON data with all metrics
   - For programmatic access

## Report Contents

The generated one-page executive brief includes:

### Header
- Client name and workload
- Migration date range
- Report generation date

### Verdict Section
- Executive summary
- Completion status
- Key risks
- Decisions needed

### Four Key Figures
1. Content migration percentage
2. Permissions restored percentage
3. Anonymous edit links count
4. Failed accounts count

### Delivery Column (CIO)
- Content completeness status
- User experience impact
- Open issues count
- Activity timeline

### Security Column (CISO)
- Anonymous edit link exposure
- External sharing analysis
- Identity hygiene (disabled accounts)
- Migration integrity

### Decisions Requested
- Sharing policy approval
- Leavers/guests handling
- Close-out authorization
- Each with assigned owner

### Footer
- Source data details
- Record counts
- Technical report reference

## Technical Details

### Prerequisites
- **Node.js**: Required to run fly-brief.js
- **Chrome or Edge**: For PDF generation
- **FLY CSV Exports**: Migration data source

### File Processing
- Uses streaming CSV parser (handles large files)
- Skips "Site objects" files (not needed)
- Auto-detects internal domains from mappings
- Workload-specific wording and analysis

### Thresholds (Configurable in FLY_Executive_Brief.html)
```javascript
contentGood: 99.5%   // Green if above this
contentWarn: 97%     // Amber if above this, red if below
permGood: 95%        // Permission restoration threshold
permWarn: 70%        // Permission warning threshold
anyoneRed: 1000      // High risk threshold for anonymous links
staleDays: 14        // Days before flagging as stalled
```

### Browser Compatibility
- HTML report works in all modern browsers
- PDF requires Chrome/Edge for rendering
- Fallback: Use browser "Print to PDF" feature

## Email Functionality (Future Enhancement)

The UI includes email options, but email delivery requires additional setup:

**To Enable Email:**
1. Choose an email service:
   - SendGrid API
   - Mailgun API
   - Microsoft Graph API (Outlook)
   - Nodemailer with SMTP
   
2. Update main.js email section with:
   - API credentials
   - Email sending logic
   - Error handling
   
3. Install required npm packages:
   ```bash
   npm install @sendgrid/mail
   # or
   npm install nodemailer
   ```

## Customization Options

### 1. **Adjust Thresholds**
Edit `C:\migration-toolkit-main\FLY\Template\FLY_Executive_Brief.html`
- Line 97-103: CFG block

### 2. **Add Workload Types**
Edit index.html workload dropdown:
- Line ~2345: Add `<option>` elements

### 3. **Modify Report Template**
Edit `FLY_Executive_Brief.html`:
- Lines 230-378: Rendering functions
- Lines 97-103: Thresholds
- Lines 10-68: Styling

### 4. **Change Output Format**
Modify fly-brief.js:
- Line 56-66: PDF generation options
- Add custom post-processing

## Troubleshooting

### "FLY brief script not found"
**Solution:** Verify files exist at `C:\migration-toolkit-main\FLY\Template\`

### "Export folder not found"
**Solution:** Check the path contains CSV files from FLY export

### "No Mappings.csv found"
**Solution:** Ensure the export folder contains a file ending in "Mappings.csv"

### "Process exited with code 1"
**Check:**
1. Node.js is installed: `node --version`
2. Edge/Chrome is installed (for PDF)
3. Export folder contains valid CSV files
4. Check Generation Log for details

### PDF not generated
**Solutions:**
1. Verify Edge or Chrome is installed
2. Check EDGE_PATH environment variable
3. Use "Download HTML" and open in browser, then Print to PDF

### Domain dropdown empty
**Solution:** 
1. Ensure discovery has been run
2. Check that domains.csv exists
3. Verify domains are in shared config

## Files Modified

### Added/Modified Files:
1. ✅ `MigrationToolkit-Web/index.html` (sidebar + view)
2. ✅ `MigrationToolkit-Web/renderer.js` (frontend logic)
3. ✅ `MigrationToolkit-Web/main.js` (backend handler)
4. ✅ `MigrationToolkit-Web/preload.js` (IPC bridge)

### New Directory:
5. ✅ `FLY/Template/` (all template files)

## Next Steps

### Immediate
1. ✅ Test the integration
2. ✅ Run a brief generation
3. ✅ Verify PDF output

### Optional Enhancements
1. 📧 Implement email delivery
2. 📊 Add more report types
3. 📅 Schedule automated generation
4. 📁 Batch processing multiple domains
5. 🔗 Direct FLY API integration

## Support

For issues or questions:
1. Check the generation log for error details
2. Review `FLY/Template/README.md` for standalone usage
3. Run `Test-FlyPrereqs.ps1` to verify setup
4. Check Node.js and browser installation

## License & Credits

**Original FLY Executive Brief Tool:**
- Created for AvePoint FLY migration reporting
- Self-contained HTML/JS architecture
- Streaming CSV analysis engine

**Integration:**
- Added to Migration Toolkit November 2026
- Maintains compatibility with standalone version
- Uses same core analysis engine

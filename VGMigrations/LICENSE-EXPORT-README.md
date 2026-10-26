# Microsoft 365 License Export Guide

This guide explains how to export user license assignments with friendly names using the Migration Toolkit.

## Overview

The toolkit provides **two ways** to export license data:

1. **Run-Assessment.ps1** (Recommended) - Full tenant assessment including licenses
2. **Export-UserLicenses.ps1** - Standalone license export script

Both methods collect:
- User license assignments with friendly names (e.g., "Microsoft 365 E3" instead of "SPE_E3")
- Aggregate license counts per SKU
- Full license details (SkuPartNumber, SkuId, License Name)

## Method 1: Run-Assessment.ps1 (Full Assessment)

### Location
```
C:\toolkit\VGMigrations\Assessment\Run-Assessment.ps1
```

### What It Collects
The assessment collects comprehensive tenant data including:
- Active Directory users, groups, devices
- Exchange Online mailboxes, groups, transport rules
- Microsoft Graph data (Teams, M365 Groups, **Licenses**, Intune devices, etc.)
- SharePoint sites and OneDrive accounts

### License Output
- **Excel Workbook** - `<VBUName>-Assessment.xlsx`
  - Sheet: "User Licenses" - One row per user with all assigned licenses
  - Sheet: "License Counts" - Aggregate counts per SKU
- **Raw JSON** (optional) - `Raw/UserLicenses.json`, `Raw/LicenseCounts.json`

### Usage
```powershell
cd C:\toolkit\VGMigrations\Assessment
.\Run-Assessment.ps1
```

### Interactive Prompts
1. **Mode**: Choose "Run Assessment" (not "Generate Mapping Files")
2. **VBU Domain**: e.g., `contoso.com`
3. **VBU Search Term**: e.g., `Contoso`
4. **VBU ID**: extensionAttribute7 value
5. **SPO Admin URL**: e.g., `https://contoso-admin.sharepoint.com`

### Authentication
The script will prompt for authentication to:
1. **SharePoint Online** (if SPO module is available)
2. **Exchange Online** (via Connect-ExchangeOnline)
3. **Microsoft Graph** (consent to scopes: User.Read.All, Organization.Read.All, etc.)

### Troubleshooting Assessment License Collection

#### Issue: "No data collected" in User Licenses sheet

**Possible Causes:**

1. **Graph connection failed**
   - Look for error message: `[FAIL] Graph connection failed:`
   - **Solution**: Re-run the assessment and ensure you complete the Graph authentication

2. **ADUsers.json not found**
   - Look for warning: `[WARN] ADUsers.json not found - user license lookup skipped`
   - **Cause**: Active Directory collection failed before Graph collection
   - **Solution**: Ensure you have network access to your AD domain controller

3. **Graph API permissions denied**
   - Look for error: `Authorization_RequestDenied` or `Forbidden`
   - **Solution**: Ensure the account running the assessment has:
     - User.Read.All (Application or Delegated)
     - Organization.Read.All (Application or Delegated)
   - Check Azure AD > App registrations > Microsoft Graph PowerShell > API permissions

4. **Users not found in Entra ID**
   - Look for warning: `User not found in Entra ID: 'user@domain.com'`
   - **Cause**: On-prem AD user not synced to Entra ID (cloud-only or sync issues)
   - **Expected**: This is normal for on-prem only accounts

#### Diagnostic Output

With the enhanced logging, you'll now see:
```
[INFO] Building SKU lookup table from subscribed SKUs...
[INFO] SKU lookup table: 5 SKU(s)
[INFO] Processing licenses for 150 user(s)...
[INFO] Progress: 50 of 150 users processed (Success: 48, Errors: 2)
[INFO] Progress: 100 of 150 users processed (Success: 95, Errors: 5)
[INFO] License lookup complete:
[INFO]   Total users processed: 150
[INFO]   Successful lookups: 145
[INFO]   Failed lookups: 5
[INFO]   Users with licenses: 120
[INFO]   Users without licenses: 25
```

#### Check Raw JSON

If the Excel shows "No data collected" but collection ran:
```powershell
# Check if the JSON file exists and has data
Get-Content "C:\toolkit\VGMigrations\Assessment\<VBUName>-<timestamp>\Raw\UserLicenses.json" | ConvertFrom-Json | Measure-Object
```

---

## Method 2: Export-UserLicenses.ps1 (Standalone)

### Location
```
C:\toolkit\VGMigrations\Export-UserLicenses.ps1
```

### When to Use
- You only need license data (not a full assessment)
- You want to export licenses on-demand
- You want to filter by domain or search term
- Quick license audit for reporting

### Usage

#### Export all users
```powershell
cd C:\toolkit\VGMigrations
.\Export-UserLicenses.ps1
```

#### Filter by domain
```powershell
.\Export-UserLicenses.ps1 -Domain "contoso.com"
```

#### Filter by name
```powershell
.\Export-UserLicenses.ps1 -SearchTerm "Contoso"
```

#### Specify output location
```powershell
.\Export-UserLicenses.ps1 -OutputPath "C:\Reports" -Domain "contoso.com"
```

### Output Files
```
UserLicenses_20261026_143022.csv
LicenseCounts_20261026_143022.csv
```

### Prerequisites
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

### Sample Output

**UserLicenses_*.csv**
```csv
DisplayName,UserPrincipalName,LicenseCount,SkuPartNumbers,LicenseNames,SkuIds
John Smith,john.smith@contoso.com,2,SPE_E3; FLOW_P1,Microsoft 365 E3; Microsoft Power Automate Plan 1,6fd2c87f-b296-42f0-b197-1e91e994b900; ...
Jane Doe,jane.doe@contoso.com,1,SPE_E5,Microsoft 365 E5,c7df2760-2c81-4ef7-b578-5b5392b571df
```

**LicenseCounts_*.csv**
```csv
SkuPartNumber,LicenseName,AssignedCount
SPE_E3,Microsoft 365 E3,450
SPE_E5,Microsoft 365 E5,125
FLOW_P1,Microsoft Power Automate Plan 1,75
POWER_BI_PRO,Power BI Pro,50
```

---

## Supported License SKUs

The toolkit includes friendly names for **140+ Microsoft 365 license SKUs**, including:

### Microsoft 365 Enterprise
- Microsoft 365 E3, E5, F1, F3
- Microsoft 365 A3/A5 (Education)

### Office 365
- Office 365 E3, E5
- Office 365 Business, Business Premium, Business Essentials
- Microsoft 365 Apps for Enterprise/Business

### Microsoft 365 Business
- Microsoft 365 Business Premium
- Microsoft 365 Business Standard
- Microsoft 365 Business Basic

### Exchange Online
- Exchange Online Plan 1, Plan 2
- Exchange Online Archiving
- Exchange Online Protection

### SharePoint
- SharePoint Online Plan 1, Plan 2
- Office Online

### Teams
- Microsoft Teams
- Teams Phone Standard
- Calling Plans (Domestic, International)
- Audio Conferencing

### Power Platform
- Power Automate (Free, Plan 1, Plan 2)
- Power BI (Free, Pro)
- Power Apps (Per User, Per App)

### Project & Visio
- Project Plan 1, 3, 5
- Visio Plan 1, Plan 2

### Security & Compliance
- Enterprise Mobility + Security E3, E5
- Azure Active Directory Premium P1, P2
- Microsoft Defender for Office 365 (Plan 1, Plan 2)
- Microsoft Defender for Cloud Apps
- Microsoft 365 E5 Compliance/Security
- Azure Rights Management

### Intune & Device Management
- Microsoft Intune
- Windows 10/11 Enterprise E3, E5

### Other Services
- Microsoft Stream
- Yammer Enterprise
- Microsoft Dynamics CRM Online
- Microsoft Defender for Endpoint

**Note**: If a SKU is not in the friendly names list, the raw `SkuPartNumber` is shown instead.

---

## Adding Custom License Friendly Names

If you encounter a license SKU that's not mapped, you can add it to the dictionary.

### For Run-Assessment.ps1

Edit: `C:\toolkit\VGMigrations\Assessment\Modules\Graph.psm1`

Find the `$script:LicenseFriendlyNames` hashtable and add:
```powershell
$script:LicenseFriendlyNames = @{
    # ... existing entries ...
    'YOUR_SKU_PARTNUMBER' = 'Your Friendly License Name'
}
```

### For Export-UserLicenses.ps1

Edit: `C:\toolkit\VGMigrations\Export-UserLicenses.ps1`

Find the `$script:LicenseFriendlyNames` hashtable and add the same entry.

### Finding Unknown SKU PartNumbers

1. Run the export and check the output CSV
2. Look in the `SkuPartNumbers` column
3. Any unfamiliar codes are unmapped SKUs
4. Look them up at: https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference

---

## Common Issues & Solutions

### Issue: "Microsoft.Graph.Authentication module not found"
**Solution:**
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

### Issue: Graph connection fails with "Insufficient privileges"
**Solution:**
The account needs these Graph API permissions:
- `User.Read.All`
- `Organization.Read.All`

Grant them in Azure AD > App registrations > Microsoft Graph PowerShell > API permissions

### Issue: "ADUsers.json not found" (Run-Assessment.ps1 only)
**Solution:**
This means the Active Directory collection failed before Graph collection ran. Ensure:
1. You're running on a domain-joined machine
2. You have network access to a domain controller
3. Your account has AD read permissions

### Issue: License export is slow
**Expected behavior**: The script queries Graph API per-user, which can take time for large tenants.
- ~2-3 seconds per user on average
- 1000 users = ~40 minutes

**Workaround**: Use domain or name filters to reduce the user count:
```powershell
.\Export-UserLicenses.ps1 -Domain "contoso.com"
```

### Issue: Some users show 0 licenses but they have licenses assigned
**Possible causes:**
1. Licenses assigned via group membership (Graph API shows direct assignments only)
2. User account in different geo location (multi-geo tenants)
3. Licensing delay (recently assigned licenses may take time to reflect)

---

## Post-Migration License Replication

After moving a domain to a new tenant, you can use the exported license data to:

1. **Purchase matching licenses** in the new tenant
   - Reference `LicenseCounts_*.csv` to see how many of each SKU you need

2. **Assign licenses to users** after migration
   - Use `UserLicenses_*.csv` to map which licenses each user had
   - PowerShell example:
     ```powershell
     $licenses = Import-Csv "UserLicenses_20261026_143022.csv"
     
     foreach ($lic in $licenses) {
         $upn = $lic.UserPrincipalName
         $skus = $lic.SkuPartNumbers -split '; '
         
         # Assign each SKU
         foreach ($sku in $skus) {
             # Add your license assignment logic here
             # Set-MgUserLicense -UserId $upn -AddLicenses @{SkuId = $skuId} -RemoveLicenses @()
         }
     }
     ```

3. **Compare before/after** licensing
   - Export licenses from the new tenant after migration
   - Compare CSVs to verify all users have equivalent licenses

---

## Questions or Issues?

If you encounter issues with license collection:

1. **Check the diagnostic output** - The enhanced logging shows exactly where failures occur
2. **Verify Graph permissions** - Ensure `User.Read.All` and `Organization.Read.All` are granted
3. **Review the Raw JSON** - Check if data was collected but not appearing in the Excel workbook
4. **Run the standalone script** - Try `Export-UserLicenses.ps1` to isolate license collection from the full assessment

---

## Version History

### 2.9.193+ (Current)
- ✅ Added 140+ license SKU friendly names
- ✅ Created standalone `Export-UserLicenses.ps1` script
- ✅ Enhanced diagnostic logging in `Graph.psm1`
- ✅ Added detailed error messages and troubleshooting guidance
- ✅ Progress reporting for large user sets
- ✅ Statistics: success/error counts, users with/without licenses

### Previous Versions
- Basic license collection via `Get-UserLicenseData`
- Limited SKU friendly names (~11 SKUs)
- Minimal error reporting

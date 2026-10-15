# Migration Toolkit - AvePoint Fly Edition

## What is This?

This is the **AvePoint Fly** edition of the Migration Toolkit. It provides a complete toolkit for migrating Microsoft 365 tenants using the AvePoint Fly migration platform.

## Platform

🔵 **AvePoint Fly Migration Platform**

This environment is configured specifically for AvePoint Fly migrations. If you need to use **Quest On Demand** instead, use the `VGMigrations-Quest` folder.

## Quick Start

1. **Launch the tool**:
   ```powershell
   .\main-menu.ps1
   ```

2. **Run M365 Discovery**:
   - Click "Discovery"
   - Select "M365 Discovery"
   - Connect to the source tenant
   - Export discovery results

3. **Configure AvePoint Fly**:
   - Click "AvePoint Fly"
   - "1. Create App Registration"
   - "2. Setup AOS Tenant & App"
   - "3. Connections & Migration Mappings"

4. **Start Migration**:
   - "4. Run Migrations"
   - "5. Monitor Jobs"
   - "6. Generate Reports"

## Main Features

### Discovery Tools
- **M365 Discovery**: Comprehensive tenant assessment

### AvePoint Fly Migration
- **App Registration**: Create Entra ID app with required permissions
- **AOS Setup**: Configure AvePoint Online Services tenant
- **Connections**: Create source/destination connections
- **Mappings**: Map workloads (SharePoint, Exchange, OneDrive, Teams, Groups)
- **Migration Runner**: Execute AvePoint Fly migrations
- **Monitoring**: Track migration job progress
- **Reporting**: Generate migration reports

### Domain Removal
- **Remove Devices**: Clean up Entra ID / Intune devices
- **Remove Domain**: Remove verified domains
- **Update UPNs**: Bulk change user principal names
- **Hide from Address Book**: Hide objects from GAL

### Miscellaneous Tools
- **Add Site Label**: Bulk label SharePoint sites
- **Set Teams Owners**: Manage Teams ownership
- **Provision OneDrive**: Pre-provision OneDrive accounts
- **Import Domains from Excel**: Batch operations

## Configuration Files

- **`workloads.json`**: Workload mappings and project suffixes
- **`domains.json`**: Domain configuration
- **`%APPDATA%\FlyMigration\config.json`**: Fly API settings (URL, Client ID, Secret)
- **`shared-config.json`**: Tenant and customer settings

## Platform-Specific Scripts

- **`aossetup.ps1`**: AvePoint Online Services tenant setup
- **`fly-migrator.ps1`**: AvePoint migration orchestration
- **`fly-reporter.ps1`**: AvePoint report generation
- **`fly-connector.js`**: Node.js helper for Fly API connections

## Common Tools (Platform-Independent)

These tools work with both AvePoint Fly and Quest On Demand:
- `Update-UPN.ps1`
- `Remove-devices.ps1`
- `remove-domain.ps1`
- `Hide-AddressBook.ps1`
- `discovery-menu.ps1`
- `Add-SiteLabel.ps1`
- `Set-TeamsOwners.ps1`
- `provision-onedrives.ps1`

## Requirements

- PowerShell 7.0+
- Microsoft.Graph PowerShell modules
- Node.js 18+ (for AOS setup automation)
- Chrome browser (for AOS authentication)
- AvePoint Fly tenant access
- Entra ID Global Administrator role (for app registration)

## Settings

Access settings via the ⚙ gear icon:

### Config Tab
- **Fly API URL**: Your AvePoint Online Services URL
- **Client ID**: Entra ID app client ID
- **Client Secret**: Entra ID app secret

### Customer Tab
- **Tenant Prefixes**: Customer identifiers
- **Account Names**: Microsoft 365 account names
- **SharePoint Admin URLs**: Admin center URLs
- **Secret Expiry**: App secret expiration date

### Workloads Tab
- **Project Suffix**: Custom suffix for each workload
- Auto-populated Policy/Source/Destination from connections

### Discovery Tab
- **Discovery Output Path**: Base folder for discovery CSVs

## Logs

All operations are logged:
- **`logs\main-menu-*.log`**: Main menu activity
- **`logs\update-upn-*.log`**: UPN update operations
- **`logs\hide-addressbook-*.log`**: Address book operations
- **`logs\FlyRunner_*.log`**: Migration execution logs

## Support

For AvePoint Fly specific issues:
- Check `logs\` directory for error messages
- Review AOS tenant configuration
- Verify Fly API credentials in Settings
- Ensure app registration has required permissions

## Related Documentation

- **`README-TWO-ENVIRONMENTS.md`**: Overview of both AvePoint and Quest environments (in parent folder: `c:\Temp\Scripts\`)

---

**Platform**: AvePoint Fly  
**Version**: 2.0  
**Last Updated**: 2026-05-31

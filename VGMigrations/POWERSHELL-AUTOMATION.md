# AvePoint Fly PowerShell Automation

## Overview

The AvePoint Fly.Client PowerShell module provides comprehensive cmdlets for automating migrations **without needing Playwright or browser automation**. All operations can be performed via PowerShell scripts.

## Why PowerShell Instead of Playwright?

✅ **More Reliable** - Direct API calls, no browser dependencies  
✅ **Faster** - No browser overhead or UI rendering  
✅ **Easier to Debug** - Clear error messages from API  
✅ **Better Logging** - Native PowerShell logging and transcripts  
✅ **Scriptable** - Easy to integrate into CI/CD pipelines  
✅ **No UI Changes** - API remains stable even if UI changes  

## Available Scripts

### 1. New-FlyProject.ps1
Creates a new migration project for a workload.

```powershell
.\New-FlyProject.ps1 -ProjectName "Contoso - SharePoint" -Workload SharePoint -Description "Q2 2026 Migration"
```

**Parameters:**
- `ProjectName` - Name of the project
- `Workload` - SharePoint, Exchange, OneDrive, Teams, TeamChat, Groups
- `Description` - Optional description

### 2. Import-FlyMappings.ps1
Imports user/mailbox/site mappings from CSV file.

```powershell
.\Import-FlyMappings.ps1 -ProjectName "Contoso - Exchange" -Workload Exchange -MappingFile "C:\mappings\exchange.csv"
```

**Parameters:**
- `ProjectName` - Name of the Fly project
- `Workload` - Workload type
- `MappingFile` - Path to CSV file
- `SkipValidation` - Skip validation (optional)

### 3. Start-FlyMigrationWorkflow.ps1
Complete end-to-end migration workflow automation.

```powershell
.\Start-FlyMigrationWorkflow.ps1 -CustomerPrefix "Contoso" -Workload SharePoint -MappingFile "C:\mappings\spo.csv"
```

**Parameters:**
- `CustomerPrefix` - Customer name
- `Workload` - Workload type
- `MappingFile` - Path to CSV file
- `StopAt` - Stop at specific stage (optional)
- `SkipPreScan` - Skip pre-scan (optional)
- `SkipVerification` - Skip verification (optional)

**What it does:**
1. Creates project (if not exists)
2. Imports mappings
3. Runs pre-scan
4. Runs verification
5. Starts migration
6. Logs everything

### 4. Get-MigrationData.ps1 (Already exists)
Gets real-time migration status and progress.

```powershell
.\Get-MigrationData.ps1 -ProjectPrefix "Contoso"
```

## Complete Migration Workflow Example

### Scenario: Migrate Contoso SharePoint

```powershell
# Step 1: Create project
.\New-FlyProject.ps1 `
    -ProjectName "Contoso - SharePoint" `
    -Workload SharePoint `
    -Description "Q2 2026 SharePoint Migration"

# Step 2: Import mappings
.\Import-FlyMappings.ps1 `
    -ProjectName "Contoso - SharePoint" `
    -Workload SharePoint `
    -MappingFile "C:\migrations\contoso\sharepoint-mappings.csv"

# Step 3: Start pre-scan
Start-FlySharePointPreScan -Project "Contoso - SharePoint"

# Step 4: Check pre-scan results
Export-FlySharePointMappingStatus `
    -Project "Contoso - SharePoint" `
    -OutFile "C:\reports\prescan-results.csv"

# Step 5: Start verification (optional)
Start-FlySharePointVerification -Project "Contoso - SharePoint"

# Step 6: Start migration
Start-FlySharePointMigration -Project "Contoso - SharePoint"

# Step 7: Monitor progress
.\Get-MigrationData.ps1 -ProjectPrefix "Contoso"
```

### Or use the workflow script (all in one):

```powershell
.\Start-FlyMigrationWorkflow.ps1 `
    -CustomerPrefix "Contoso" `
    -Workload SharePoint `
    -MappingFile "C:\migrations\contoso\sharepoint-mappings.csv"
```

## Available Fly.Client Cmdlets

### Connection Management
```powershell
Connect-Fly -Url $apiUrl -ClientId $clientId -ClientSecret $clientSecret
Disconnect-Fly
```

### Project Management
```powershell
New-FlyMigrationProject -Name "Project Name"
Get-FlyMigrationProject -Name "Project Name"
Import-FlyMigrationProjects -Path "projects.csv"
```

### Policy Management (per workload)
```powershell
New-FlySharePointPolicy -Name "Policy Name" -Project "Project Name"
New-FlyExchangePolicy -Name "Policy Name" -Project "Project Name"
New-FlyOneDrivePolicy -Name "Policy Name" -Project "Project Name"
New-FlyTeamsPolicy -Name "Policy Name" -Project "Project Name"
New-FlyTeamChatPolicy -Name "Policy Name" -Project "Project Name"
New-FlyM365GroupPolicy -Name "Policy Name" -Project "Project Name"
```

### Import Mappings (per workload)
```powershell
Import-FlySharePointMappings -Project "Project" -Path "mappings.csv"
Import-FlyExchangeMappings -Project "Project" -Path "mappings.csv"
Import-FlyOneDriveMappings -Project "Project" -Path "mappings.csv"
Import-FlyTeamsMappings -Project "Project" -Path "mappings.csv"
Import-FlyTeamChatMappings -Project "Project" -Path "mappings.csv"
Import-FlyM365GroupMappings -Project "Project" -Path "mappings.csv"
```

### Pre-Scan
```powershell
Start-FlySharePointPreScan -Project "Project"
Start-FlyExchangePreScan -Project "Project"
Start-FlyOneDrivePreScan -Project "Project"
Start-FlyTeamsPreScan -Project "Project"
Start-FlyM365GroupPreScan -Project "Project"
# Note: TeamChat doesn't have pre-scan
```

### Verification
```powershell
Start-FlySharePointVerification -Project "Project"
Start-FlyExchangeVerification -Project "Project"
Start-FlyOneDriveVerification -Project "Project"
Start-FlyTeamsVerification -Project "Project"
Start-FlyTeamChatVerification -Project "Project"
Start-FlyM365GroupVerification -Project "Project"
```

### Start Migration
```powershell
Start-FlySharePointMigration -Project "Project"
Start-FlyExchangeMigration -Project "Project"
Start-FlyOneDriveMigration -Project "Project"
Start-FlyTeamsMigration -Project "Project"
Start-FlyTeamChatMigration -Project "Project"
Start-FlyM365GroupMigration -Project "Project"
```

### Export Status & Reports
```powershell
Export-FlySharePointMappingStatus -Project "Project" -OutFile "status.csv"
Export-FlySharePointMigrationReport -Project "Project" -OutFile "report.csv"
# Similar cmdlets for Exchange, OneDrive, Teams, TeamChat, Groups
```

## CSV Mapping File Format

### SharePoint
```csv
SourceSiteUrl,DestinationSiteUrl
https://source.sharepoint.com/sites/site1,https://dest.sharepoint.com/sites/site1
https://source.sharepoint.com/sites/site2,https://dest.sharepoint.com/sites/site2
```

### Exchange
```csv
SourceUserPrincipalName,DestinationUserPrincipalName
user1@source.com,user1@destination.com
user2@source.com,user2@destination.com
```

### OneDrive
```csv
SourceUserPrincipalName,DestinationUserPrincipalName
user1@source.com,user1@destination.com
user2@source.com,user2@destination.com
```

### Teams
```csv
SourceTeamName,DestinationTeamName
Sales Team,Sales Team
Marketing Team,Marketing Team
```

## Error Handling

All scripts include comprehensive error handling:

```powershell
try {
    Start-FlySharePointMigration -Project "Contoso - SharePoint" -ErrorAction Stop
} catch {
    Write-Error "Migration failed: $_"
    # Send notification, log to system, etc.
}
```

## Logging

The workflow script creates detailed logs:

```
logs/migration-Contoso-SharePoint-20260604-143022.log
```

Contains:
- Timestamps for all actions
- Success/failure status
- API responses
- Error details

## Integration with Migration Toolkit

The Migration Toolkit UI can call these PowerShell scripts instead of using Playwright:

**Before (Playwright):**
- Launch browser
- Navigate to pages
- Fill forms
- Click buttons
- Wait for responses

**After (PowerShell):**
```javascript
await window.electronAPI.executeScript('New-FlyProject.ps1', [
  '-ProjectName', 'Contoso - SharePoint',
  '-Workload', 'SharePoint'
]);
```

## Advantages Summary

| Feature | Playwright | PowerShell |
|---------|-----------|------------|
| Speed | Slow (browser) | Fast (API) |
| Reliability | UI-dependent | API-stable |
| Error Messages | Generic | Specific |
| Logging | Limited | Comprehensive |
| Debugging | Difficult | Easy |
| Maintenance | High | Low |
| Automation | Complex | Simple |

## Next Steps

1. **Replace Playwright automation** in the toolkit with PowerShell scripts
2. **Add UI buttons** that call these scripts directly
3. **Show progress** in the toolkit from script output
4. **Store logs** in the logs folder for review
5. **Handle errors** gracefully with user-friendly messages

## Support

For Fly.Client cmdlet help:
```powershell
Get-Help <CmdletName> -Detailed
Get-Help New-FlyMigrationProject -Examples
```

For script help:
```powershell
Get-Help .\New-FlyProject.ps1 -Detailed
```

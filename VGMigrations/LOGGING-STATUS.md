# Migration Toolkit - Logging Status

## Overview
This document tracks which scripts have comprehensive logging and which need it added.

## Logging Standard
All scripts should:
1. Import the `logging.ps1` module
2. Call `Initialize-Logging` at the start
3. Use `Write-Log` for all output
4. Use `Write-LogException` for error handling
5. Call `Complete-Logging` at the end

## Scripts with Full Logging ✅

### Discovery & Reporting
- [x] Check-TenantLicenses.ps1
- [x] Check-DomainMigrationReadiness.ps1
- [x] search-domain.ps1
- [x] discovery-menu.ps1
- [x] reports.ps1

### Setup & Configuration
- [x] Setup-LicenseReportApp.ps1
- [x] Setup-DomainReadinessApp.ps1
- [x] Invoke-TenantBaseline.ps1

### Domain Management
- [x] Domain-Removal-Workflow.ps1
- [x] remove-domain.ps1

### Migration Workflows
- [x] Start-FlyMigrationWorkflow.ps1
- [x] Start-FlyMigrationStage.ps1
- [x] Stop-FlyMigrationStage.ps1
- [x] Import-FlyMappings.ps1
- [x] New-FlyProject.ps1

### User Management
- [x] Update-UPN.ps1
- [x] Update-OnPremUPN.ps1
- [x] Hide-AddressBook.ps1

### Core Libraries
- [x] lib.ps1 (has Write-Log for GUI)
- [x] main-menu.ps1
- [x] menu.ps1
- [x] monitor.ps1

## Scripts Needing Logging ❌

### High Priority (User-Facing)
- [ ] **Get-TenantLicenseReport.ps1** - License reporting (similar to Check-TenantLicenses)
- [ ] **Get-UserMailboxLicenses.ps1** - User license details
- [ ] **New-AzureAppRegistration.ps1** - App registration creation
- [ ] **Set-DestinationLicenses.ps1** - License assignment
- [ ] **Clear-EmployeeId.ps1** - Employee ID management
- [ ] **Provision-OneDrive-User.ps1** - OneDrive provisioning
- [ ] **Check-OneDriveStatus.ps1** - OneDrive status checking
- [ ] **New-SharePointSites.ps1** - SharePoint site creation
- [ ] **Remove-devices.ps1** - Device removal
- [ ] **Retire-Devices.ps1** - Device retirement
- [ ] **Remove-EntraUsers.ps1** - User deletion
- [ ] **Add-ExchangeAdmin.ps1** - Exchange admin rights

### Medium Priority
- [ ] Import-DomainsFromExcel.ps1
- [ ] Remove-AliasAddresses.ps1
- [ ] Remove-DeletedSharePointSites.ps1
- [ ] Rename-DomainObjects.ps1
- [ ] Set-TeamsOwners.ps1
- [ ] Set-TeamsOwners-Run.ps1
- [ ] Update-SIPDomain.ps1
- [ ] Add-SiteLabel.ps1
- [ ] Deduplicate-Inventory.ps1
- [ ] Get-DomainDevices.ps1
- [ ] Get-MappingError.ps1

### Low Priority (Utilities/Internal)
- [ ] Archive-OldLogs.ps1
- [ ] Build-Exe.ps1
- [ ] Check-Updates.ps1
- [ ] Deploy-ToServer.ps1
- [ ] Encrypt-Secret.ps1
- [ ] Fix-ConfigBOM.ps1
- [ ] Install-Prerequisites.ps1
- [ ] New-Package.ps1
- [ ] Test-AutoUpdate.ps1
- [ ] Test-FlyConnection.ps1
- [ ] Save-Config.ps1
- [ ] Setup.ps1

### Special Cases
- [ ] Get-TenantLicenseReport-InteractiveWorker.ps1 (spawned subprocess)
- [ ] Connect-MicrosoftGraph.ps1 (connection helper)
- [ ] Ensure-GraphModules.ps1 (module installer)
- [ ] fly-migrator.ps1
- [ ] fly-reporter.ps1
- [ ] provision-onedrives.ps1
- [ ] run-multiple-domains.ps1
- [ ] runner.ps1
- [ ] connections.ps1
- [ ] settings.ps1

## Log File Locations

All log files are saved to:
```
C:\Users\<username>\Volaris Group\GRP Data Security (Volaris Consolidated) - M365 Migrations\Logs\
```

Log file naming convention:
```
<ScriptName>_<yyyy-MM-dd_HHmmss>.log
```

Example:
```
Check-TenantLicenses_2026-07-29_183045.log
```

## Next Steps

1. Copy `logging.ps1` to C:\toolkit\VGMigrations\
2. Manually add logging to high-priority scripts
3. Test logging with Check-TenantLicenses.ps1
4. Roll out to remaining scripts gradually
5. Update version.json with logging additions

## Usage Example

```powershell
# At the start of a script
$loggingModule = Join-Path $PSScriptRoot 'logging.ps1'
if (Test-Path $loggingModule) {
    Import-Module $loggingModule -Force
    $logFile = Initialize-Logging -ScriptName "MyScript"
}

# Throughout the script
Write-Log "Processing tenant: $tenantName"
Write-Log "Operation completed successfully" -Level SUCCESS
Write-Log "Warning: Skipping item" -Level WARN
Write-Log "Error occurred" -Level ERROR

# In error handling
try {
    # ... code ...
} catch {
    Write-LogException -Exception $_ -Context "ProcessTenant"
    Complete-Logging -Success $false -Summary $_.Exception.Message
    throw
}

# At the end
Complete-Logging -Success $true -Summary "Processed 5 tenants"
```

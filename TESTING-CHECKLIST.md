# Migration Toolkit - Testing Checklist

## Pre-Testing Setup
- [ ] Node.js 18+ installed and verified
- [ ] PowerShell 7+ installed and verified
- [ ] AvePoint Fly.Client module installed
- [ ] Have valid AvePoint Fly API credentials
- [ ] Have at least one active migration in AvePoint Fly

## Installation Testing
- [ ] Run INSTALL.bat successfully
- [ ] No errors during npm install
- [ ] All dependencies installed

## Application Launch
- [ ] START.bat launches the application
- [ ] Application window opens and maximizes
- [ ] Sidebar is visible on the left
- [ ] Dashboard view loads by default
- [ ] No JavaScript errors in console (Ctrl+Shift+I)

## Settings Configuration
### Config Tab
- [ ] Settings dialog opens when clicking ⚙ icon
- [ ] Can enter API URL
- [ ] Can enter Client ID
- [ ] Can enter Client Secret (field shows as password •••)
- [ ] Can select Secret Expiry date
- [ ] Can enter Portal URL
- [ ] Can enter SharePoint Admin URL
- [ ] "Test Connection" button works
- [ ] Connection test shows success or specific error
- [ ] "Save All" saves configuration
- [ ] After closing and reopening, fields are populated (except secret shows as ••••••••)
- [ ] "Check for Updates" button responds

### Customer Tab
- [ ] Can add new customer
- [ ] Can enter Prefix, Account Name, and SharePoint URL
- [ ] "Add Customer" button creates new row
- [ ] "Remove Customer" button removes last row
- [ ] "Save All" saves customer data
- [ ] After closing and reopening Settings, customers are loaded correctly

## Dashboard Functionality
- [ ] Domain dropdown populates with customer prefixes
- [ ] Can select different domains from dropdown
- [ ] "Refresh" button updates dashboard data
- [ ] "Open AvePoint Fly" button opens web portal
- [ ] Stat boxes show correct numbers:
  - [ ] Items in Scope
  - [ ] Completed
  - [ ] In Progress
  - [ ] Needs Attention (red border)
- [ ] Clicking stat boxes opens detail dialog
- [ ] Detail dialogs show relevant data
- [ ] Workload progress bars display
- [ ] Workload bars show correct status (on track/warnings/failed)
- [ ] Progress percentages are accurate

## Real Data Integration
- [ ] Dashboard loads real data from AvePoint Fly (not sample data)
- [ ] Workload bars show actual migration progress
- [ ] Clicking "Needs Attention" shows actual failed/warning items
- [ ] Clicking "In Progress" shows current migrations
- [ ] Clicking "Completed" shows finished items
- [ ] Data matches what's shown in AvePoint Fly portal

## Navigation & Views
### Discovery
- [ ] Discovery view loads
- [ ] Can enter domain information
- [ ] Radio buttons work (Single/Multiple domains)
- [ ] Options checkboxes work
- [ ] Browse buttons work
- [ ] "Start Discovery" button is visible

### AvePoint Fly Views
- [ ] App Registration view displays correctly
- [ ] AOS Setup view displays correctly
- [ ] Connections & Mappings view displays correctly
- [ ] Migration Reports view displays correctly
- [ ] Monitor view displays correctly
- [ ] All views show content without covering sidebar

### Misc Scripts
- [ ] Provision OneDrives view loads
- [ ] Tenant URL dropdown populates with customer tenants
- [ ] Dropdown shows format: "Prefix - URL"
- [ ] Teams Migration view loads

### Domain Removal
- [ ] Workflow view loads
- [ ] All sub-views accessible

## Layout & UI
- [ ] Sidebar always visible (240px wide)
- [ ] Content area doesn't overlap sidebar
- [ ] All views have proper padding
- [ ] Text is readable and properly aligned
- [ ] Buttons are clickable
- [ ] Dropdowns work correctly
- [ ] Modal dialogs display centered
- [ ] Modal dialogs close properly

## Data Persistence
- [ ] API credentials persist after app restart
- [ ] Customer data persists after app restart
- [ ] Client secret remains encrypted in config file
- [ ] Dashboard remembers last selected domain
- [ ] Settings changes are saved correctly

## Error Handling
- [ ] Appropriate error messages for invalid API credentials
- [ ] Graceful handling when Fly.Client module not installed
- [ ] Clear error messages when migration data unavailable
- [ ] Validation for required fields
- [ ] User-friendly error messages (no raw stack traces)

## Performance
- [ ] Application launches in < 5 seconds
- [ ] Dashboard data loads in < 10 seconds
- [ ] Switching views is smooth
- [ ] No lag when clicking buttons
- [ ] Refresh updates data without freezing UI

## Security
- [ ] Client secret is encrypted in config file
- [ ] Passwords display as •••• in UI
- [ ] No sensitive data in console logs
- [ ] Config file location is in %APPDATA%

## Edge Cases
- [ ] Works with no customers configured
- [ ] Works with empty migration (0 items)
- [ ] Handles migrations with only 1 workload
- [ ] Handles large datasets (1000+ items)
- [ ] Works when AvePoint Fly is offline/unavailable
- [ ] Handles expired API credentials gracefully

## Documentation
- [ ] README.md is clear and comprehensive
- [ ] QUICK-START.txt provides easy onboarding
- [ ] INSTALL.bat has clear output
- [ ] Error messages point to documentation

## Issues Found
Document any issues below with:
- Steps to reproduce
- Expected behavior
- Actual behavior
- Screenshots if applicable
- Console errors (if any)

---

### Issue 1:
**Description**: 
**Steps**: 
**Expected**: 
**Actual**: 
**Priority**: High/Medium/Low

### Issue 2:
**Description**: 
**Steps**: 
**Expected**: 
**Actual**: 
**Priority**: High/Medium/Low

---

## Sign-off
- [ ] All critical functionality tested
- [ ] All issues documented
- [ ] Ready for production use

**Tester Name**: ___________________
**Date**: ___________________
**Version Tested**: ___________________

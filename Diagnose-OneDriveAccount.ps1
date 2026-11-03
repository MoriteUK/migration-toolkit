<#
.SYNOPSIS
    Diagnose OneDrive access issues for a user account

.DESCRIPTION
    This script helps diagnose why a user's OneDrive cannot be accessed.
    It checks user status, OneDrive provisioning, and permissions.

.PARAMETER UserEmail
    The email address or GUID-based identifier to investigate
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$UserEmail
)

Write-Host "=== OneDrive Access Diagnostic Tool ===" -ForegroundColor Cyan
Write-Host "Target User: $UserEmail" -ForegroundColor Yellow
Write-Host ""

# Check if Microsoft.Graph modules are installed
$requiredModules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Users', 'Microsoft.Graph.Sites')
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "Installing $module module..." -ForegroundColor Yellow
        Install-Module -Name $module -Force -AllowClobber -Scope CurrentUser
    }
}

Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Users
Import-Module Microsoft.Graph.Sites

try {
    # Connect to Microsoft Graph
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    $scopes = @("User.Read.All", "Sites.Read.All", "Files.Read.All", "Directory.Read.All")
    Connect-MgGraph -Scopes $scopes
    Write-Host "Connected successfully" -ForegroundColor Green
    Write-Host ""

    # Get current context
    $context = Get-MgContext
    Write-Host "Connected as: $($context.Account)" -ForegroundColor Green
    Write-Host "Tenant ID: $($context.TenantId)" -ForegroundColor Green
    Write-Host ""

    # Attempt to find the user
    Write-Host "=== USER LOOKUP ===" -ForegroundColor Cyan
    $user = $null

    # Try 1: Direct lookup
    Write-Host "Method 1: Direct lookup by UserPrincipalName..." -ForegroundColor Yellow
    try {
        $user = Get-MgUser -UserId $UserEmail -Property Id,DisplayName,Mail,UserPrincipalName,AccountEnabled,UserType,AssignedLicenses -ErrorAction Stop
        Write-Host "  SUCCESS: Found user" -ForegroundColor Green
    }
    catch {
        Write-Host "  FAILED: $($_.Exception.Message)" -ForegroundColor Red
    }

    # Try 2: GUID extraction
    if (-not $user -and $UserEmail -match '^([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})@') {
        $guid = $matches[1]
        Write-Host "Method 2: Lookup by extracted GUID ($guid)..." -ForegroundColor Yellow
        try {
            $user = Get-MgUser -UserId $guid -Property Id,DisplayName,Mail,UserPrincipalName,AccountEnabled,UserType,AssignedLicenses -ErrorAction Stop
            Write-Host "  SUCCESS: Found user" -ForegroundColor Green
        }
        catch {
            Write-Host "  FAILED: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Try 3: Filter search
    if (-not $user) {
        Write-Host "Method 3: Search with filter..." -ForegroundColor Yellow
        try {
            $users = Get-MgUser -Filter "userPrincipalName eq '$UserEmail' or mail eq '$UserEmail'" -Property Id,DisplayName,Mail,UserPrincipalName,AccountEnabled,UserType,AssignedLicenses
            if ($users -and $users.Count -gt 0) {
                $user = $users[0]
                Write-Host "  SUCCESS: Found user" -ForegroundColor Green
            }
            else {
                Write-Host "  FAILED: No users matched the filter" -ForegroundColor Red
            }
        }
        catch {
            Write-Host "  FAILED: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    if ($user) {
        Write-Host ""
        Write-Host "=== USER DETAILS ===" -ForegroundColor Cyan
        Write-Host "Display Name: $($user.DisplayName)" -ForegroundColor White
        Write-Host "User Principal Name: $($user.UserPrincipalName)" -ForegroundColor White
        Write-Host "Mail: $($user.Mail)" -ForegroundColor White
        Write-Host "User ID: $($user.Id)" -ForegroundColor White
        Write-Host "Account Enabled: $($user.AccountEnabled)" -ForegroundColor $(if($user.AccountEnabled){'Green'}else{'Red'})
        Write-Host "User Type: $($user.UserType)" -ForegroundColor White
        Write-Host "Has Licenses: $($user.AssignedLicenses.Count -gt 0)" -ForegroundColor $(if($user.AssignedLicenses.Count -gt 0){'Green'}else{'Red'})
        Write-Host ""

        # Check OneDrive
        Write-Host "=== ONEDRIVE CHECK ===" -ForegroundColor Cyan
        try {
            $drive = Get-MgUserDrive -UserId $user.Id -ErrorAction Stop
            if ($drive) {
                Write-Host "OneDrive Status: PROVISIONED" -ForegroundColor Green
                Write-Host "Drive ID: $($drive.Id)" -ForegroundColor White
                Write-Host "Drive Name: $($drive.Name)" -ForegroundColor White
                Write-Host "Drive Type: $($drive.DriveType)" -ForegroundColor White
                Write-Host "Web URL: $($drive.WebUrl)" -ForegroundColor White
                Write-Host ""

                # Try to get root folder
                Write-Host "Testing access to OneDrive root..." -ForegroundColor Yellow
                try {
                    $root = Get-MgUserDriveRoot -UserId $user.Id -ErrorAction Stop
                    Write-Host "  SUCCESS: Can access OneDrive root" -ForegroundColor Green
                    Write-Host "  Root ID: $($root.Id)" -ForegroundColor White

                    # Try to list items
                    Write-Host "Testing file listing..." -ForegroundColor Yellow
                    try {
                        $items = Get-MgUserDriveRootChild -UserId $user.Id -DriveId $drive.Id -Top 5
                        Write-Host "  SUCCESS: Can list files" -ForegroundColor Green
                        Write-Host "  Sample items found: $($items.Count)" -ForegroundColor White
                        if ($items.Count -gt 0) {
                            Write-Host "  First few items:" -ForegroundColor White
                            foreach ($item in $items) {
                                $type = if ($item.Folder) { "Folder" } else { "File" }
                                Write-Host "    - [$type] $($item.Name)" -ForegroundColor Gray
                            }
                        }
                    }
                    catch {
                        Write-Host "  FAILED: Cannot list files - $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
                catch {
                    Write-Host "  FAILED: Cannot access root - $($_.Exception.Message)" -ForegroundColor Red
                }
            }
            else {
                Write-Host "OneDrive Status: NOT PROVISIONED" -ForegroundColor Red
                Write-Host "This user may not have OneDrive enabled or it hasn't been initialized." -ForegroundColor Yellow
            }
        }
        catch {
            Write-Host "OneDrive Status: ERROR" -ForegroundColor Red
            Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host ""
            Write-Host "This could mean:" -ForegroundColor Yellow
            Write-Host "  - OneDrive is not provisioned for this user" -ForegroundColor Yellow
            Write-Host "  - User doesn't have required license" -ForegroundColor Yellow
            Write-Host "  - OneDrive is disabled in tenant settings" -ForegroundColor Yellow
        }
    }
    else {
        Write-Host ""
        Write-Host "=== USER NOT FOUND ===" -ForegroundColor Red
        Write-Host "Could not locate user: $UserEmail" -ForegroundColor Red
        Write-Host ""
        Write-Host "Possible reasons:" -ForegroundColor Yellow
        Write-Host "  1. This is not a real user account (might be a shared mailbox or resource)" -ForegroundColor Yellow
        Write-Host "  2. The account has been deleted" -ForegroundColor Yellow
        Write-Host "  3. The email format is incorrect" -ForegroundColor Yellow
        Write-Host "  4. You don't have permission to see this user" -ForegroundColor Yellow
        Write-Host ""

        # Show a sample of users
        Write-Host "Showing first 5 users in your tenant for reference:" -ForegroundColor Cyan
        try {
            $sampleUsers = Get-MgUser -Top 5 -Property Id,DisplayName,UserPrincipalName,Mail,UserType
            foreach ($sampleUser in $sampleUsers) {
                Write-Host "  - $($sampleUser.DisplayName)" -ForegroundColor White
                Write-Host "    UPN: $($sampleUser.UserPrincipalName)" -ForegroundColor Gray
                Write-Host "    Mail: $($sampleUser.Mail)" -ForegroundColor Gray
                Write-Host "    Type: $($sampleUser.UserType)" -ForegroundColor Gray
                Write-Host ""
            }
        }
        catch {
            Write-Host "Could not retrieve sample users: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    Write-Host "=== DIAGNOSTIC COMPLETE ===" -ForegroundColor Cyan
    Disconnect-MgGraph
}
catch {
    Write-Host ""
    Write-Host "=== FATAL ERROR ===" -ForegroundColor Red
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}

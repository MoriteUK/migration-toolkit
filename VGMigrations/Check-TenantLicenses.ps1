#Requires -Version 7.0
<#
.SYNOPSIS
    Generates a report of mailbox-granting licenses by domain with automatic app registration
.DESCRIPTION
    Connects to Microsoft 365 tenants, auto-creates app registration if needed, and counts licenses that grant Exchange Online mailboxes
.PARAMETER TenantIDsPath
    Path to the Tenant IDs.xlsx file. Defaults to the standard location.
.PARAMETER OutputPath
    Path where the report will be saved
.PARAMETER AppName
    Name for the app registration (default: "VG-License-Reporter")
.PARAMETER CredentialStorePath
    Path to store app credentials (default: script directory)
#>
[CmdletBinding()]
param(
    [string]$TenantIDsPath = "C:\Users\Andy White\Volaris Group\GRP Data Security (Volaris Consolidated) - M365 Migrations\Tenant IDs.xlsx",
    [string]$OutputPath = "C:\Users\Andy White\Volaris Group\GRP Data Security (Volaris Consolidated) - M365 Migrations\Logs",
    [string]$AppName = "VG-License-Reporter",
    [string]$CredentialStorePath = $PSScriptRoot
)

$ErrorActionPreference = 'Stop'

# Load library
$libPath = Join-Path $PSScriptRoot 'lib.ps1'
if (Test-Path $libPath) {
    try { . $libPath }
    catch { Write-Warning "Failed to load lib.ps1: $_" }
}

function Write-Log {
    param([string]$Msg, [string]$Level = 'INFO')
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
    $line = "$timestamp  [$($Level.PadRight(5))]  $Msg"
    Write-Host $line
}

function Get-AppCredentials {
    param(
        [string]$TenantId,
        [string]$TenantName
    )

    # Check if we have stored credentials
    $credFile = Join-Path $CredentialStorePath "appcreds_$($TenantId).json"
    if (Test-Path $credFile) {
        $creds = Get-Content $credFile | ConvertFrom-Json
        Write-Log "Using stored credentials (App: $($creds.AppId))"
        return $creds
    }

    # No credentials found - need setup
    Write-Log "No app registration found for $TenantName" "ERROR"
    Write-Log "" "ERROR"
    Write-Log "FIRST TIME SETUP REQUIRED:" "ERROR"
    Write-Log "Run this command in PowerShell 7:" "ERROR"
    Write-Log "  .\Setup-LicenseReportApp.ps1 -TenantId '$TenantId' -TenantName '$TenantName'" "ERROR"
    Write-Log "" "ERROR"
    throw "App registration not found - run Setup-LicenseReportApp.ps1 first"
}

function Get-OrCreateAppRegistration-OLD {
    param(
        [string]$TenantId,
        [string]$TenantName,
        [string]$AppName
    )

    Write-Log "Checking for existing app registration in $TenantName..."

    try {
        # Check if we have a stored secret for this tenant
        $credFile = Join-Path $CredentialStorePath "appcreds_$($TenantId).json"
        if (Test-Path $credFile) {
            Write-Log "Using existing credentials"
            $creds = Get-Content $credFile | ConvertFrom-Json
            return $creds
        }

        # No credentials - need setup
        Write-Log "No existing credentials found, setting up app registration..."

        # First, connect interactively to check/create the app
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Connect-MgGraph -TenantId $TenantId -Scopes "Application.ReadWrite.All", "Directory.Read.All" -NoWelcome

        # Check if app exists
        $existingApp = Get-MgApplication -Filter "displayName eq '$AppName'" -ErrorAction SilentlyContinue

        if ($existingApp) {
            Write-Log "App '$AppName' already exists (AppId: $($existingApp.AppId))"
            $appId = $existingApp.AppId
            $objectId = $existingApp.Id
        } else {
            Write-Log "Creating new app registration '$AppName'..."

            # Required resource access for Microsoft Graph
            $graphResourceId = "00000003-0000-0000-c000-000000000000" # Microsoft Graph

            # Application permissions we need
            $requiredPermissions = @(
                @{
                    Id = "7ab1d382-f21e-4acd-a863-ba3e13f7da61" # Directory.Read.All
                    Type = "Role"
                }
                @{
                    Id = "9a5d68dd-52b0-4cc2-bd40-abcf44ac3a30" # Application.Read.All
                    Type = "Role"
                }
                @{
                    Id = "b0afded3-3588-46d8-8b3d-9842eff778da" # Organization.Read.All
                    Type = "Role"
                }
            )

            # Build the required resource access object
            $resourceAccess = $requiredPermissions | ForEach-Object {
                @{
                    Id = $_.Id
                    Type = $_.Type
                }
            }

            $requiredResourceAccess = @(
                @{
                    ResourceAppId = $graphResourceId
                    ResourceAccess = $resourceAccess
                }
            )

            # Create the application
            $appParams = @{
                DisplayName = $AppName
                SignInAudience = "AzureADMyOrg"
                RequiredResourceAccess = $requiredResourceAccess
            }

            $newApp = New-MgApplication @appParams
            $appId = $newApp.AppId
            $objectId = $newApp.Id

            Write-Log "App created with AppId: $appId"

            # Wait a moment for app to propagate
            Start-Sleep -Seconds 3
        }

        # Create new client secret (valid for 2 years)
        Write-Log "Creating new client secret..."
        $secretName = "AutoGenerated-$(Get-Date -Format 'yyyyMMdd-HHmmss')"

        $passwordCred = @{
            DisplayName = $secretName
            EndDateTime = (Get-Date).AddYears(2)
        }

        $secret = Add-MgApplicationPassword -ApplicationId $objectId -PasswordCredential $passwordCred
        $clientSecret = $secret.SecretText

        Write-Log "Client secret created (expires: $($secret.EndDateTime))"

        # Check if service principal exists, create if not
        $sp = Get-MgServicePrincipal -Filter "appId eq '$appId'" -ErrorAction SilentlyContinue

        if (-not $sp) {
            Write-Log "Creating service principal..."
            $sp = New-MgServicePrincipal -AppId $appId
            Start-Sleep -Seconds 2
        }

        # Grant admin consent for the required permissions
        Write-Log "Granting admin consent for permissions..."

        # Get the Microsoft Graph service principal
        $graphSP = Get-MgServicePrincipal -Filter "appId eq '$graphResourceId'"

        foreach ($permission in $requiredPermissions) {
            try {
                $grant = @{
                    ClientId = $sp.Id
                    ConsentType = "AllPrincipals"
                    ResourceId = $graphSP.Id
                    Scope = $null
                    PrincipalId = $null
                }

                # For application permissions, we need to create app role assignments
                $appRoleAssignment = @{
                    PrincipalId = $sp.Id
                    ResourceId = $graphSP.Id
                    AppRoleId = $permission.Id
                }

                New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -BodyParameter $appRoleAssignment -ErrorAction SilentlyContinue | Out-Null
            } catch {
                # May already exist, that's okay
                Write-Log "Permission already granted or error: $($_.Exception.Message)" "WARN"
            }
        }

        Write-Log "Admin consent granted" "SUCCESS"

        # Store credentials securely
        $credFile = Join-Path $CredentialStorePath "appcreds_$($TenantId).json"
        $credObject = @{
            TenantId = $TenantId
            TenantName = $TenantName
            AppId = $appId
            ClientSecret = $clientSecret
            SecretExpires = $secret.EndDateTime
            CreatedDate = Get-Date
        }

        # Convert to JSON and save with restricted permissions
        $credObject | ConvertTo-Json | Out-File -FilePath $credFile -Encoding UTF8 -Force

        # Set file to read-only for current user
        $acl = Get-Acl $credFile
        $acl.SetAccessRuleProtection($true, $false)
        $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            [System.Security.Principal.WindowsIdentity]::GetCurrent().Name,
            "Read",
            "Allow"
        )
        $acl.SetAccessRule($rule)
        Set-Acl $credFile $acl

        Write-Log "Credentials saved to: $credFile" "SUCCESS"

        return [PSCustomObject]$credObject

    } catch {
        Write-Log "ERROR: Failed to create/configure app registration: $_" "ERROR"
        throw
    } finally {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
}

function Connect-WithAppCredentials {
    param(
        [string]$TenantId,
        [string]$AppId,
        [string]$ClientSecret
    )

    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

        # Create secure string for client secret
        $secureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential($AppId, $secureSecret)

        # Connect using app credentials
        Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $credential -NoWelcome

        return $true
    } catch {
        Write-Log "ERROR connecting with app credentials: $_" "ERROR"
        return $false
    }
}

# Function to check if a SKU grants Exchange Online mailbox or OneDrive
# Same logic as Get-TenantLicenseReport.ps1 - checks service plans instead of hardcoded SKU list
function Test-SkuGrantsMailboxOrOneDrive($Sku) {
    foreach ($plan in $Sku.ServicePlans) {
        $name = $plan.ServicePlanName
        # Skip EXCHANGE_S_FOUNDATION (internal plumbing, not a real mailbox)
        if ($name -match '_FOUNDATION$') { continue }
        # Skip SHAREPOINTWAC (Office-for-web viewing, not storage)
        if ($name -eq 'SHAREPOINTWAC' -or $name -eq 'SHAREPOINTWAC_EDU') { continue }
        # Check for Exchange or SharePoint service plans
        if ($name -match '^EXCHANGE_S_' -or $name -match '^SHAREPOINT') { return $true }
    }
    return $false
}

Write-Log "=== Mailbox License Report (Auto-App) Started ==="
Write-Log "Tenant IDs file: $TenantIDsPath"
Write-Log "App Name: $AppName"

# Check if file exists
if (-not (Test-Path $TenantIDsPath)) {
    Write-Log "ERROR: Tenant IDs file not found: $TenantIDsPath" "ERROR"
    exit 1
}

# Check if Microsoft Graph module is available
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    Write-Log "ERROR: Microsoft.Graph module not installed. Install with: Install-Module Microsoft.Graph -Scope CurrentUser" "ERROR"
    exit 1
}

# Verify file exists
if (-not (Test-Path $TenantIDsPath)) {
    Write-Log "ERROR: Tenant IDs file not found: $TenantIDsPath" "ERROR"
    Write-Log "Please ensure the file exists at the specified path." "ERROR"
    exit 1
}

# Read Excel file using COM
Write-Log "Opening Excel file..."
$excel = $null
$workbook = $null
$tenants = [System.Collections.Generic.List[object]]::new()

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    Write-Log "Opening workbook: $TenantIDsPath"
    $workbook = $excel.Workbooks.Open($TenantIDsPath, $null, $true)  # Open read-only

    if (-not $workbook) {
        throw "Failed to open workbook - workbook object is null"
    }

    $worksheet = $workbook.Worksheets.Item(1)
    $usedRange = $worksheet.UsedRange

    $rowCount = $usedRange.Rows.Count
    $colCount = $usedRange.Columns.Count

    Write-Log "Found $rowCount rows and $colCount columns"

    # Process data rows (starting from row 2)
    # Column layout: A=FriendlyName, B=TenantId, C=AppId, D=AppSecret, J=licenses checked, N=cutover done, P=Domains (comma-separated)
    for ($row = 2; $row -le $rowCount; $row++) {
        $friendlyName = $usedRange.Cells.Item($row, 1).Text.Trim()  # Column A = Friendly Name

        # Skip empty rows
        if ([string]::IsNullOrWhiteSpace($friendlyName)) {
            continue
        }

        # Check column J (column 10) and column N (column 14) for "Yes"
        $columnJValue = $usedRange.Cells.Item($row, 10).Text.Trim()
        $columnNValue = $usedRange.Cells.Item($row, 14).Text.Trim()

        # Skip if column J OR column N has "Yes" (case-insensitive)
        if ($columnJValue.ToLower() -eq "yes" -or $columnNValue.ToLower() -eq "yes") {
            $whichCol = if ($columnNValue.ToLower() -eq "yes") { "N" } elseif ($columnJValue.ToLower() -eq "yes") { "J" } else { "" }
            Write-Log "Row $row : [$friendlyName] - Skipping (Column $whichCol = Yes)"
            continue
        }

        # Get Tenant ID from column B (column 2)
        $tenantIdValue = $usedRange.Cells.Item($row, 2).Text.Trim()

        # Validate it's a proper GUID
        if (-not ($tenantIdValue -match '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')) {
            Write-Log "Row $row : [$friendlyName] - Skipping (Column B does not contain a valid Tenant ID GUID)" "WARN"
            continue
        }

        # Get domains from column P (column 16) - comma-separated list
        $domainsValue = $usedRange.Cells.Item($row, 16).Text.Trim()

        $tenantData = @{
            Name = $friendlyName
            TenantId = $tenantIdValue
            Domain = $domainsValue
        }

        $tenants.Add([PSCustomObject]$tenantData)
        Write-Log "Found tenant: $friendlyName (Domain: $($tenantData.Domain), TenantId: $($tenantData.TenantId))"
    }

} catch {
    Write-Log "ERROR reading Excel file: $_" "ERROR"
    throw
} finally {
    # Aggressive cleanup of Excel COM objects
    if ($workbook) {
        try {
            $workbook.Close($false)
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        } catch { Write-Log "Error closing workbook: $_" "WARN" }
    }
    if ($excel) {
        try {
            $excel.Workbooks.Close()
            $excel.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        } catch { Write-Log "Error quitting Excel: $_" "WARN" }
    }

    # Force garbage collection
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()

    # Kill any lingering Excel processes as last resort
    Start-Sleep -Milliseconds 500
    Get-Process excel -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq "" } | Stop-Process -Force -ErrorAction SilentlyContinue
}

Write-Log "Found $($tenants.Count) tenants to process"

# Results collection
$results = [System.Collections.Generic.List[object]]::new()

foreach ($tenant in $tenants) {
    Write-Log ""
    Write-Log "=== Processing: $($tenant.Name) ===" "INFO"

    try {
        # Determine tenant ID to use
        $tenantIdToUse = if ($tenant.TenantId) { $tenant.TenantId } else { $tenant.Domain }

        if (-not $tenantIdToUse) {
            Write-Log "ERROR: No tenant ID or domain found" "ERROR"
            throw "No tenant ID or domain available"
        }

        # Check for existing credentials
        $credFile = Join-Path $CredentialStorePath "appcreds_$tenantIdToUse.json"

        if (Test-Path $credFile) {
            Write-Log "Loading existing app credentials..."
            $creds = Get-Content $credFile | ConvertFrom-Json

            # Check if secret is expired
            if ($creds.SecretExpires) {
                $expiryDate = [DateTime]$creds.SecretExpires
                if ($expiryDate -lt (Get-Date).AddDays(30)) {
                    Write-Log "Client secret expires soon ($expiryDate)" "WARN"
                    Write-Log "Run Setup-LicenseReportApp.ps1 to regenerate" "WARN"
                }
            }
        } else {
            # No credentials - get them (this will throw with setup instructions)
            $creds = Get-AppCredentials -TenantId $tenantIdToUse -TenantName $tenant.Name
        }

        # Connect using app credentials
        Write-Log "Connecting to tenant with app credentials..."
        $connected = Connect-WithAppCredentials -TenantId $creds.TenantId -AppId $creds.AppId -ClientSecret $creds.ClientSecret

        if (-not $connected) {
            throw "Failed to connect with app credentials"
        }

        # Get organization details
        $org = Get-MgOrganization
        $primaryDomain = $org.VerifiedDomains | Where-Object { $_.IsDefault } | Select-Object -ExpandProperty Name

        Write-Log "Connected to: $primaryDomain"

        # Get all subscribed SKUs
        Write-Log "Querying licenses..."
        $skus = Get-MgSubscribedSku -All

        $mailboxLicenseCount = 0
        $mailboxSKUDetails = @()

        foreach ($sku in $skus) {
            # Use the service plan checking function
            if (Test-SkuGrantsMailboxOrOneDrive $sku) {
                $consumed = $sku.ConsumedUnits
                $mailboxLicenseCount += $consumed

                $mailboxSKUDetails += [PSCustomObject]@{
                    SKU = $skuPartNumber
                    Consumed = $consumed
                    Total = $sku.PrepaidUnits.Enabled
                }

                Write-Log "  $skuPartNumber : $consumed licenses consumed"
            }
        }

        # Add to results
        $results.Add([PSCustomObject]@{
            Domain = $primaryDomain
            TenantName = $tenant.Name
            TotalMailboxLicenses = $mailboxLicenseCount
            SKUDetails = ($mailboxSKUDetails | ForEach-Object { "$($_.SKU):$($_.Consumed)" }) -join '; '
            Status = 'Success'
        })

        Write-Log "Total mailbox licenses: $mailboxLicenseCount" "SUCCESS"

    } catch {
        Write-Log "ERROR processing tenant: $_" "ERROR"
        $results.Add([PSCustomObject]@{
            Domain = if ($tenant.Domain) { $tenant.Domain } else { "Unknown" }
            TenantName = $tenant.Name
            TotalMailboxLicenses = 0
            SKUDetails = "Error: $_"
            Status = 'Failed'
        })
    } finally {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
}

# Generate report
Write-Log ""
Write-Log "=== Generating Report ==="

# Ensure output directory exists
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    Write-Log "Created output directory: $OutputPath"
}

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$reportFile = Join-Path $OutputPath "MailboxLicenseReport_$timestamp.csv"
$summaryFile = Join-Path $OutputPath "MailboxLicenseSummary_$timestamp.txt"

# Export detailed CSV
$results | Export-Csv -Path $reportFile -NoTypeInformation -Encoding UTF8
Write-Log "Detailed report saved: $reportFile"

# Generate summary
$summary = @"
=== MAILBOX LICENSE REPORT SUMMARY ===
Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

TOTAL DOMAINS PROCESSED: $($results.Count)
SUCCESSFUL: $($results | Where-Object { $_.Status -eq 'Success' } | Measure-Object | Select-Object -ExpandProperty Count)
FAILED: $($results | Where-Object { $_.Status -eq 'Failed' } | Measure-Object | Select-Object -ExpandProperty Count)

=== MAILBOX LICENSE COUNT BY DOMAIN ===

"@

foreach ($result in $results | Sort-Object Domain) {
    $summary += "Domain: $($result.Domain)`n"
    $summary += "  Tenant Name: $($result.TenantName)`n"
    $summary += "  Mailbox Licenses: $($result.TotalMailboxLicenses)`n"
    $summary += "  Status: $($result.Status)`n"
    if ($result.SKUDetails -and $result.Status -eq 'Success') {
        $summary += "  SKU Breakdown: $($result.SKUDetails)`n"
    }
    $summary += "`n"
}

$totalLicenses = ($results | Where-Object { $_.Status -eq 'Success' } | Measure-Object -Property TotalMailboxLicenses -Sum).Sum
$summary += "=== GRAND TOTAL: $totalLicenses MAILBOX LICENSES ===`n"

# Save summary
$summary | Out-File -FilePath $summaryFile -Encoding UTF8
Write-Log "Summary saved: $summaryFile"

# Display summary to console
Write-Log ""
Write-Host $summary

Write-Log ""
Write-Log "=== Mailbox License Report Complete ===" "SUCCESS"
Write-Log "Detailed CSV: $reportFile"
Write-Log "Summary TXT: $summaryFile"
Write-Log ""
Write-Log "App credentials stored in: $CredentialStorePath\appcreds_*.json"
Write-Log "These will be reused on subsequent runs for automatic authentication."

# Signal to Electron to open the CSV file
Write-Output "##OPEN_FILE:$reportFile##"

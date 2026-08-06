#Requires -Version 7.0
<#
.SYNOPSIS
    Checks if domains are added to target tenant and if DNS verification records are ready
.DESCRIPTION
    Reads source and destination tenants from Tenant IDs.xlsx:
    - Column P: Source tenant(s) - can be comma-separated for multiple sources
    - Column Q: Destination tenant

    For each domain in the source tenant(s):
    1. Checks if domain exists in target tenant
    2. If not, checks if MS-verify TXT record exists in public DNS
    3. Reports readiness status for each domain
.PARAMETER TenantIDsPath
    Path to the Tenant IDs.xlsx file. Defaults to the standard location.
.PARAMETER OutputPath
    Path where the report will be saved
.PARAMETER ProcessAll
    Process all rows. If false (default), skips rows with "Yes" in column J or N
.EXAMPLE
    .\Check-DomainMigrationReadiness.ps1
.EXAMPLE
    .\Check-DomainMigrationReadiness.ps1 -ProcessAll
#>
[CmdletBinding()]
param(
    [string]$TenantIDsPath = "C:\Users\Andy White\Volaris Group\GRP Data Security (Volaris Consolidated) - M365 Migrations\Tenant IDs.xlsx",

    [string]$OutputPath = "C:\Users\Andy White\Volaris Group\GRP Data Security (Volaris Consolidated) - M365 Migrations",

    [string]$AppName = "VG-DomainReadiness-Reporter",

    [string]$CredentialStorePath = $PSScriptRoot,

    [switch]$ProcessAll
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
    $credFile = Join-Path $CredentialStorePath "domaincheck_$($TenantId).json"
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
    Write-Log "  .\Setup-DomainReadinessApp.ps1 -TenantId '$TenantId' -TenantName '$TenantName'" "ERROR"
    Write-Log "" "ERROR"
    throw "App registration not found - run Setup-DomainReadinessApp.ps1 first"
}

function Connect-WithAppCredentials {
    param(
        [string]$TenantId,
        [string]$AppId,
        [string]$ClientSecret
    )

    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        $secureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential($AppId, $secureSecret)
        Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $credential -NoWelcome
    } catch {
        Write-Log "Failed to connect with app credentials: $_" "ERROR"
        throw
    }
}

function Check-DNSForRecord {
    param(
        [string]$Domain,
        [string]$ExpectedValue,
        [switch]$CheckForAnyMSRecord
    )

    try {
        # Query TXT records
        $txtRecords = Resolve-DnsName -Name $Domain -Type TXT -ErrorAction SilentlyContinue

        if (-not $txtRecords) {
            Write-Log "    DNS: No TXT records found" "WARN"
            return @{
                Found = $false
                Message = "No TXT records found in DNS"
                ActualValue = $null
            }
        }

        Write-Log "    DNS: Found $($txtRecords.Count) TXT records in public DNS"

        # If checking for any MS= record (domain in another tenant scenario)
        if ($CheckForAnyMSRecord) {
            $msRecords = @()
            foreach ($record in $txtRecords) {
                $recordString = $record.Strings -join ''
                Write-Log "    DNS: Checking record: $recordString"

                if ($recordString -match '^MS=ms\d+') {
                    Write-Log "    DNS: MS= verification record FOUND!" "SUCCESS"
                    $msRecords += $recordString
                }
            }

            if ($msRecords.Count -gt 0) {
                return @{
                    Found = $true
                    Message = "MS= verification record present in DNS (domain currently in another tenant - cannot validate exact value until transfer)"
                    ActualValue = $msRecords[0]
                }
            } else {
                return @{
                    Found = $false
                    Message = "No MS= verification record found in DNS"
                    ActualValue = $null
                }
            }
        }

        # Normal check - validate exact expected value
        $matchFound = $false
        $foundMSRecords = @()
        foreach ($record in $txtRecords) {
            $recordString = $record.Strings -join ''
            Write-Log "    DNS: Checking record: $recordString"

            if ($recordString -match '^MS=ms\d+') {
                $foundMSRecords += $recordString
            }

            if ($recordString -eq $ExpectedValue) {
                Write-Log "    DNS: MATCH FOUND!" "SUCCESS"
                $matchFound = $true
                break
            }
        }

        if ($matchFound) {
            return @{
                Found = $true
                Message = "Correct verification record found in DNS"
                ActualValue = $ExpectedValue
            }
        } else {
            if ($foundMSRecords.Count -gt 0) {
                Write-Log "    DNS: Found MS= record(s) but none match expected value" "WARN"
                Write-Log "    DNS: Expected: $ExpectedValue" "WARN"
                Write-Log "    DNS: Found: $($foundMSRecords -join ', ')" "WARN"
                return @{
                    Found = $false
                    Message = "MS= record exists but does not match expected value (Found: $($foundMSRecords[0]))"
                    ActualValue = $foundMSRecords[0]
                }
            } else {
                Write-Log "    DNS: No MS= verification record found" "WARN"
                return @{
                    Found = $false
                    Message = "No MS= verification record found in DNS"
                    ActualValue = $null
                }
            }
        }

    } catch {
        return @{
            Found = $false
            Message = "DNS query error: $_"
            ActualValue = $null
        }
    }
}

Write-Log "=== Domain Migration Readiness Check Started ==="
Write-Log "Tenant IDs file: $TenantIDsPath"

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
$migrationPairs = [System.Collections.Generic.List[object]]::new()

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

    # Read headers from row 1
    $headers = @{}
    for ($col = 1; $col -le $colCount; $col++) {
        $headerValue = $usedRange.Cells.Item(1, $col).Text
        if ($headerValue) {
            $headers[$col] = $headerValue
        }
    }

    # Process data rows (starting from row 2)
    for ($row = 2; $row -le $rowCount; $row++) {
        $tenantName = $usedRange.Cells.Item($row, 1).Text.Trim()

        # Skip empty rows
        if ([string]::IsNullOrWhiteSpace($tenantName)) {
            continue
        }

        # Check skip conditions
        if (-not $ProcessAll) {
            # Column N (column 14) = "Yes" indicates migration complete
            $columnNValue = $usedRange.Cells.Item($row, 14).Text.Trim()
            if ($columnNValue.ToLower() -eq "yes") {
                Write-Log "Row $row : [$tenantName] - Skipping (Migration Complete - Column N = Yes)"
                continue
            }

            # Column L (column 12) = "Yes" indicates DNS already requested/added
            $columnLValue = $usedRange.Cells.Item($row, 12).Text.Trim()
            if ($columnLValue.ToLower() -eq "yes") {
                Write-Log "Row $row : [$tenantName] - Skipping (DNS Already Requested - Column L = Yes)"
                continue
            }
        }

        # Read columns:
        # Column B = Primary domain name (column 2) - but might be a tenant ID GUID, skip if so
        # Column P = Domain names (comma-separated) (column 16)
        # Column Q = Target tenant ID (column 17)
        # Column R = Login credentials (optional) (column 18)
        $primaryDomain = $usedRange.Cells.Item($row, 2).Text.Trim()
        $additionalDomains = $usedRange.Cells.Item($row, 16).Text.Trim()
        $destinationTenant = $usedRange.Cells.Item($row, 17).Text.Trim()
        $loginCredential = $usedRange.Cells.Item($row, 18).Text.Trim()

        # Build domain list - only use Column P (not Column B if it's a GUID)
        $domainNames = ""

        # Check if primary domain is a GUID pattern (skip if so)
        if ($primaryDomain -match '^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$') {
            # It's a GUID, skip it
            $domainNames = $additionalDomains
        } else {
            # It's a domain name, include it
            $domainNames = $primaryDomain
            if (-not [string]::IsNullOrWhiteSpace($additionalDomains)) {
                $domainNames += ",$additionalDomains"
            }
        }

        # Skip if no domain or destination
        if ([string]::IsNullOrWhiteSpace($domainNames) -or [string]::IsNullOrWhiteSpace($destinationTenant)) {
            Write-Log "Row $row : [$tenantName] - Skipping (No domain/destination in columns P/Q)"
            continue
        }

        # Split domains by comma if multiple
        $domainList = $domainNames -split ',' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

        foreach ($domain in $domainList) {
            $migrationPairs.Add([PSCustomObject]@{
                TenantName = $tenantName
                DomainName = $domain
                DestinationTenant = $destinationTenant
                LoginCredential = $loginCredential
                Row = $row
            })

            Write-Log "Row $row : [$tenantName] - Domain: $domain -> Destination: $destinationTenant"
        }
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

Write-Log "Found $($migrationPairs.Count) source->destination migration pairs to process"

if ($migrationPairs.Count -eq 0) {
    Write-Log "No migration pairs found to process. Exiting." "WARN"
    exit 0
}

# Results collection
$allResults = [System.Collections.Generic.List[object]]::new()

# Process each migration pair
foreach ($pair in $migrationPairs) {
    Write-Log ""
    Write-Log "=== Processing: $($pair.TenantName) ===" "INFO"
    Write-Log "Domain: $($pair.DomainName)"
    Write-Log "Destination: $($pair.DestinationTenant)"

    $domainToCheck = $pair.DomainName

    $result = [PSCustomObject]@{
        TenantName = $pair.TenantName
        DomainName = $domainToCheck
        DestinationTenant = $pair.DestinationTenant
        DNS_RecordType = ""
        DNS_Label = ""
        DNS_Value = ""
        DNSVerifyRecord = ""
        ReadinessStatus = ""
        Notes = ""
    }

    # STEP 1: Get app credentials and connect
    Write-Log "Connecting to destination tenant: $($pair.DestinationTenant)"
    try {
        # Get app credentials (must be set up via Setup-DomainReadinessApp.ps1 first)
        $appCreds = Get-AppCredentials -TenantId $pair.DestinationTenant -TenantName $pair.TenantName

        # Connect using app credentials
        Connect-WithAppCredentials -TenantId $appCreds.TenantId -AppId $appCreds.AppId -ClientSecret $appCreds.ClientSecret

        # Check if domain is already added to target
        Write-Log "Checking if domain exists in target tenant..."
        $targetDomains = Get-MgDomain
        $targetDomain = $targetDomains | Where-Object { $_.Id -eq $domainToCheck }

        $domainInAnotherTenant = $false
        $cannotGetVerificationRecord = $false

        if (-not $targetDomain) {
            Write-Log "  ℹ Domain NOT yet added to target tenant - attempting to add..."
            try {
                # Add domain to tenant
                $newDomain = New-MgDomain -Id $domainToCheck
                Write-Log "  ✓ Domain added to target tenant successfully" "SUCCESS"
                $targetDomain = $newDomain
            } catch {
                # Check if error is because domain is in another tenant
                if ($_ -match "domain is already added to a different|domain.*already.*organization|domain.*cannot be added") {
                    Write-Log "  ℹ Domain is currently in another tenant (expected pre-cutover)" "INFO"
                    Write-Log "  Will check DNS for MS= verification record presence"
                    $domainInAnotherTenant = $true
                    $cannotGetVerificationRecord = $true
                } else {
                    Write-Log "  ✗ Failed to add domain: $_" "ERROR"
                    $result.ReadinessStatus = "Error"
                    $result.Notes = "Failed to add domain to target tenant: $_"
                    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
                    $allResults.Add($result)
                    continue
                }
            }
        } else {
            Write-Log "  ℹ Domain is already in target tenant (Verified: $($targetDomain.IsVerified))"

            # If domain is already verified, mark as ready and skip DNS check
            if ($targetDomain.IsVerified) {
                Write-Log "  ✓ Domain is already verified in target tenant" "SUCCESS"
                $result.ReadinessStatus = "Ready"
                $result.Notes = "Domain already verified in target tenant"
                $result.DNS_RecordType = "N/A"
                $result.DNS_Label = "N/A"
                $result.DNS_Value = "Already verified"
                $result.DNSVerifyRecord = "Already verified"
                Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
                $allResults.Add($result)
                continue
            }
        }

        # Get the verification DNS record from the target tenant (if possible)
        if ($cannotGetVerificationRecord) {
            Write-Log "  ⚠ Cannot retrieve exact verification record - domain in another tenant"
            Write-Log "  Will check DNS for any MS= verification record instead"
            $result.DNS_RecordType = "TXT"
            $result.DNS_Label = "@"
            $result.DNS_Value = "Unknown - domain in another tenant"
            # We'll check DNS for ANY MS= record below
        } else {
            Write-Log "  Retrieving DNS verification record..."
            try {
                # Try the verification DNS records endpoint instead of service configuration
                $verificationRecords = Get-MgDomainVerificationDnsRecord -DomainId $domainToCheck

                Write-Log "  Found $($verificationRecords.Count) verification DNS records"

                # Find the TXT record for MS= verification
                $txtRecords = $verificationRecords | Where-Object {
                    $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.domainDnsTxtRecord'
                }

                Write-Log "  Found $($txtRecords.Count) TXT records"

                # Debug: Log each TXT record value
                foreach ($rec in $txtRecords) {
                    $txtValue = $rec.AdditionalProperties.text
                    Write-Log "    TXT Record: $txtValue"
                }

                # Filter for the MS= verification record (not SPF or other records)
                $verifyRecord = $txtRecords | Where-Object {
                    $_.AdditionalProperties.text -match '^MS=ms\d+'
                } | Select-Object -First 1

                if ($verifyRecord) {
                    $recordType = "TXT"
                    $recordLabel = if ($verifyRecord.AdditionalProperties.label) { $verifyRecord.AdditionalProperties.label } else { "@" }
                    $recordValue = $verifyRecord.AdditionalProperties.text

                    Write-Log "  DNS Record to configure:"
                    Write-Log "    Type: $recordType"
                    Write-Log "    Label: $recordLabel"
                    Write-Log "    Value: $recordValue"

                    $result.DNS_RecordType = $recordType
                    $result.DNS_Label = $recordLabel
                    $result.DNS_Value = $recordValue
                    $result.DNSVerifyRecord = $recordValue
                } else {
                    Write-Log "  ⚠ Could not retrieve MS= verification record from target" "WARN"
                    $result.DNS_RecordType = "TXT"
                    $result.DNS_Label = "Unable to retrieve"
                    $result.DNS_Value = "Unable to retrieve MS= record from target"
                    $result.ReadinessStatus = "Error"
                    $result.Notes = "Domain added but could not retrieve MS= verification record"
                    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
                    $allResults.Add($result)
                    continue
                }
            } catch {
                Write-Log "  ✗ Error retrieving verification record: $_" "ERROR"
                $result.DNS_RecordType = "TXT"
                $result.DNS_Label = "Error"
                $result.DNS_Value = "Error: $_"
                $result.ReadinessStatus = "Error"
                $result.Notes = "Failed to retrieve DNS verification record: $_"
                Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
                $allResults.Add($result)
                continue
            }
        }

        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

    } catch {
        Write-Log "ERROR connecting to target tenant: $_" "ERROR"

        # Check if it's a permissions issue
        if ($_ -match "Insufficient privileges|Authorization_RequestDenied") {
            Write-Log "" "ERROR"
            Write-Log "APP PERMISSIONS ISSUE DETECTED:" "ERROR"
            Write-Log "The app registration exists but lacks admin consent." "ERROR"
            Write-Log "Re-run setup to refresh permissions:" "ERROR"
            Write-Log "  .\Setup-DomainReadinessApp.ps1 -TenantId '$($pair.DestinationTenant)' -TenantName '$($pair.TenantName)'" "ERROR"
            Write-Log "" "ERROR"
            $result.Notes = "Insufficient privileges to complete the operation - re-run Setup-DomainReadinessApp.ps1"
        } else {
            $result.Notes = "Failed to connect to target tenant: $_"
        }

        $result.ReadinessStatus = "Error"
        $allResults.Add($result)
        continue
    }

    # STEP 2: Check if the expected verification record is in DNS
    Write-Log "Checking public DNS for verification record..."

    if ($domainInAnotherTenant) {
        # Domain is in another tenant - check for ANY MS= record
        Write-Log "  Checking for any MS= verification record (exact value unknown)"
        $dnsCheck = Check-DNSForRecord -Domain $domainToCheck -ExpectedValue "" -CheckForAnyMSRecord

        if ($dnsCheck.Found) {
            Write-Log "  ✓ $($dnsCheck.Message)" "SUCCESS"
            $result.ReadinessStatus = "Ready - Pre-Cutover"
            $result.Notes = "MS= verification record found in DNS. Domain currently in source tenant - record will be validated on cutover."
            $result.DNS_Value = $dnsCheck.ActualValue
            $result.DNSVerifyRecord = $dnsCheck.ActualValue
        } else {
            Write-Log "  ✗ $($dnsCheck.Message)" "WARN"
            $result.ReadinessStatus = "Not Ready"
            $result.Notes = "No MS= verification record found in DNS. Add verification record before cutover. Transfer domain to target tenant first to get exact value, or use existing verification record from source tenant."
        }
    } else {
        # Normal flow - validate exact expected value
        Write-Log "  Expected value: $($result.DNS_Value)"
        $dnsCheck = Check-DNSForRecord -Domain $domainToCheck -ExpectedValue $result.DNS_Value

        if ($dnsCheck.Found) {
            Write-Log "  ✓ $($dnsCheck.Message)" "SUCCESS"
            $result.ReadinessStatus = "Ready"
            $result.Notes = "DNS verification record is configured correctly"
        } else {
            Write-Log "  ✗ $($dnsCheck.Message)" "WARN"
            $result.ReadinessStatus = "Not Ready"
            if ($dnsCheck.ActualValue) {
                $result.Notes = "Wrong MS= record in DNS. Expected: $($result.DNS_Value), Found: $($dnsCheck.ActualValue)"
            } else {
                $result.Notes = "DNS record missing - Request DNS team to add: Type=TXT, Host=$($result.DNS_Label), Value=$($result.DNS_Value)"
            }
        }
    }

    $allResults.Add($result)
}

# Generate report
Write-Log ""
Write-Log "=== Generating Report ==="

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$reportFile = Join-Path $OutputPath "DomainMigrationReadiness_$timestamp.csv"
$summaryFile = Join-Path $OutputPath "DomainMigrationReadiness_$timestamp.txt"

# Export detailed CSV
$allResults | Export-Csv -Path $reportFile -NoTypeInformation -Encoding UTF8
Write-Log "Detailed report saved: $reportFile"

# Open the CSV file automatically
try {
    Write-Log "Opening CSV report in Excel..."
    Invoke-Item $reportFile
} catch {
    Write-Log "Could not auto-open CSV: $_" "WARN"
}

# Generate summary
$totalDomains = $allResults.Count
$ready = ($allResults | Where-Object { $_.ReadinessStatus -eq "Ready" }).Count
$notReady = ($allResults | Where-Object { $_.ReadinessStatus -eq "Not Ready" }).Count
$errors = ($allResults | Where-Object { $_.ReadinessStatus -eq "Error" }).Count

$summary = @"
=== DOMAIN MIGRATION READINESS SUMMARY ===
Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

TOTAL DOMAINS CHECKED: $totalDomains

STATUS BREAKDOWN:
  Ready: $ready (Domain added to target, correct DNS record present)
  Not Ready: $notReady (Domain not added OR DNS record missing/incorrect)
  Error: $errors (Connection or other errors)

=== DETAILED STATUS BY TENANT ===

"@

# Group by tenant name
$groupedResults = $allResults | Group-Object -Property TenantName

foreach ($group in $groupedResults) {
    $summary += "`n--- $($group.Name) ---`n"

    # Show destination mapping
    $firstResult = $group.Group[0]
    $summary += "  Destination: $($firstResult.DestinationTenant)`n"
    $summary += "  Domains: $($group.Count)`n`n"

    # Ready
    $readyDomains = $group.Group | Where-Object { $_.ReadinessStatus -eq "Ready" }
    if ($readyDomains) {
        $summary += "  READY ($($readyDomains.Count)):`n"
        foreach ($r in $readyDomains) {
            $summary += "    ✓ $($r.DomainName)`n"
        }
        $summary += "`n"
    }

    # Not Ready
    $notReadyDomains = $group.Group | Where-Object { $_.ReadinessStatus -eq "Not Ready" }
    if ($notReadyDomains) {
        $summary += "  NOT READY ($($notReadyDomains.Count)) - DNS Records Needed:`n"
        foreach ($r in $notReadyDomains) {
            $summary += "    Domain: $($r.DomainName)`n"
            $summary += "      Type: $($r.DNS_RecordType)`n"
            $summary += "      Label: $($r.DNS_Label)`n"
            $summary += "      Value: $($r.DNS_Value)`n"
            $summary += "`n"
        }
    }

    # Errors
    $errorDomains = $group.Group | Where-Object { $_.ReadinessStatus -eq "Error" }
    if ($errorDomains) {
        $summary += "  ERROR ($($errorDomains.Count)):`n"
        foreach ($r in $errorDomains) {
            $summary += "    $($r.DomainName) - $($r.Notes)`n"
        }
        $summary += "`n"
    }
}

$summary += "`n=== NEXT STEPS ===`n"
$summary += "Ready: DNS verification record matches target tenant - ready for cutover`n"
$summary += "Not Ready: Either domain not added to target yet, or DNS record missing/incorrect`n"

# Save summary
$summary | Out-File -FilePath $summaryFile -Encoding UTF8
Write-Log "Summary saved: $summaryFile"

# Display summary to console
Write-Log ""
Write-Host $summary

Write-Log ""
Write-Log "=== Domain Migration Readiness Check Complete ===" "SUCCESS"
Write-Log "Detailed CSV: $reportFile"
Write-Log "Summary TXT: $summaryFile"

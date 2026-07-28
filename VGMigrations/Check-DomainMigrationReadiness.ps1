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

function Get-DNSVerificationRecord {
    param([string]$Domain)

    try {
        # Query TXT records for MS verification
        $txtRecords = Resolve-DnsName -Name $Domain -Type TXT -ErrorAction SilentlyContinue

        if ($txtRecords) {
            $msVerifyRecords = $txtRecords | Where-Object {
                $_.Strings -match '^MS=ms\d+'
            }

            if ($msVerifyRecords) {
                return @{
                    Found = $true
                    Record = ($msVerifyRecords[0].Strings -match '^MS=ms\d+')[0]
                    AllRecords = $msVerifyRecords | ForEach-Object { $_.Strings -join '; ' }
                }
            }
        }

        return @{
            Found = $false
            Record = $null
            AllRecords = $null
        }
    } catch {
        Write-Log "DNS query failed for $Domain : $_" "WARN"
        return @{
            Found = $false
            Record = "DNS query failed: $_"
            AllRecords = $null
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

        # Check column J (column 10) and column N (column 14) for "Yes"
        if (-not $ProcessAll) {
            $columnJValue = $usedRange.Cells.Item($row, 10).Text.Trim()
            $columnNValue = $usedRange.Cells.Item($row, 14).Text.Trim()

            # Skip if column J OR column N has "Yes" (case-insensitive)
            if ($columnJValue.ToLower() -eq "yes" -or $columnNValue.ToLower() -eq "yes") {
                $whichCol = if ($columnNValue.ToLower() -eq "yes") { "N" } elseif ($columnJValue.ToLower() -eq "yes") { "J" } else { "" }
                Write-Log "Row $row : [$tenantName] - Skipping (Column $whichCol = Yes)"
                continue
            }
        }

        # Read column P (source tenant - column 16) and Q (destination tenant - column 17)
        $sourceTenants = $usedRange.Cells.Item($row, 16).Text.Trim()
        $destinationTenant = $usedRange.Cells.Item($row, 17).Text.Trim()

        # Skip if no source or destination
        if ([string]::IsNullOrWhiteSpace($sourceTenants) -or [string]::IsNullOrWhiteSpace($destinationTenant)) {
            Write-Log "Row $row : [$tenantName] - Skipping (No source/destination in columns P/Q)"
            continue
        }

        # Split source tenants by comma if multiple
        $sourceList = $sourceTenants -split ',' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

        foreach ($source in $sourceList) {
            $migrationPairs.Add([PSCustomObject]@{
                TenantName = $tenantName
                SourceTenant = $source
                DestinationTenant = $destinationTenant
                Row = $row
            })

            Write-Log "Row $row : [$tenantName] - Source: $source -> Destination: $destinationTenant"
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
    Write-Log "Source: $($pair.SourceTenant)"
    Write-Log "Destination: $($pair.DestinationTenant)"

    # Connect to source tenant to get domains
    Write-Log "Connecting to source tenant..."
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Connect-MgGraph -TenantId $pair.SourceTenant -Scopes "Domain.Read.All" -NoWelcome

        # Get source tenant info
        $sourceOrg = Get-MgOrganization
        $sourceTenantName = $sourceOrg.DisplayName
        Write-Log "Source tenant: $sourceTenantName"

        # Get domains from source tenant
        Write-Log "Retrieving domains from source tenant..."
        $sourceDomains = Get-MgDomain | Where-Object {
            # Skip .onmicrosoft.com domains
            $_.Id -notmatch '\.onmicrosoft\.com$'
        }
        Write-Log "Found $($sourceDomains.Count) custom domains in source tenant"

        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

    } catch {
        Write-Log "ERROR connecting to source tenant: $_" "ERROR"

        # Add error record and continue to next pair
        $allResults.Add([PSCustomObject]@{
            TenantName = $pair.TenantName
            SourceTenant = $pair.SourceTenant
            DestinationTenant = $pair.DestinationTenant
            Domain = "N/A"
            InSourceTenant = $false
            InTargetTenant = $false
            DNSVerifyRecordFound = $false
            DNSVerifyRecord = $null
            ReadinessStatus = "Error"
            Recommendation = "Failed to connect to source tenant"
            Notes = "Error: $_"
        })
        continue
    }

    # Connect to target tenant
    Write-Log "Connecting to destination tenant..."
    try {
        Connect-MgGraph -TenantId $pair.DestinationTenant -Scopes "Domain.Read.All" -NoWelcome

        # Get target tenant info
        $targetOrg = Get-MgOrganization
        $targetTenantName = $targetOrg.DisplayName
        Write-Log "Destination tenant: $targetTenantName"

        # Get domains from target tenant
        Write-Log "Retrieving domains from destination tenant..."
        $targetDomains = Get-MgDomain
        $targetDomainNames = $targetDomains | Select-Object -ExpandProperty Id
        Write-Log "Found $($targetDomains.Count) total domains in destination tenant"

        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

    } catch {
        Write-Log "ERROR connecting to destination tenant: $_" "ERROR"

        # Add error record and continue to next pair
        $allResults.Add([PSCustomObject]@{
            TenantName = $pair.TenantName
            SourceTenant = $pair.SourceTenant
            DestinationTenant = $pair.DestinationTenant
            Domain = "N/A"
            InSourceTenant = $false
            InTargetTenant = $false
            DNSVerifyRecordFound = $false
            DNSVerifyRecord = $null
            ReadinessStatus = "Error"
            Recommendation = "Failed to connect to destination tenant"
            Notes = "Error: $_"
        })
        continue
    }

    # Check each domain
    Write-Log "Checking domain readiness..."
    foreach ($domain in $sourceDomains) {
        $domainName = $domain.Id
        Write-Log "  Checking: $domainName"

        $result = [PSCustomObject]@{
            TenantName = $pair.TenantName
            SourceTenant = $pair.SourceTenant
            DestinationTenant = $pair.DestinationTenant
            Domain = $domainName
            InSourceTenant = $true
            InTargetTenant = $false
            DNSVerifyRecordFound = $false
            DNSVerifyRecord = $null
            ReadinessStatus = ""
            Recommendation = ""
            Notes = ""
        }

        # Check if domain exists in target tenant
        if ($targetDomainNames -contains $domainName) {
            Write-Log "    ✓ Domain exists in destination tenant" "SUCCESS"
            $result.InTargetTenant = $true
            $result.ReadinessStatus = "Already Added"
            $result.Recommendation = "Domain is already in destination tenant"

            # Check verification status in target
            $targetDomain = $targetDomains | Where-Object { $_.Id -eq $domainName }
            if ($targetDomain.IsVerified) {
                $result.Notes = "Verified in destination tenant"
            } else {
                $result.Notes = "Added but NOT verified in destination tenant"
                $result.ReadinessStatus = "Added - Needs Verification"
                $result.Recommendation = "Complete domain verification in destination tenant"
            }
        } else {
            Write-Log "    ✗ Domain NOT in destination tenant" "WARN"
            $result.InTargetTenant = $false

            # Check DNS for verification record
            Write-Log "    Checking DNS for MS verification record..."
            $dnsCheck = Get-DNSVerificationRecord -Domain $domainName

            if ($dnsCheck.Found) {
                Write-Log "    ✓ MS verification record found in DNS: $($dnsCheck.Record)" "SUCCESS"
                $result.DNSVerifyRecordFound = $true
                $result.DNSVerifyRecord = $dnsCheck.Record
                $result.ReadinessStatus = "Ready to Add"
                $result.Recommendation = "Add domain to destination tenant and verify"
                $result.Notes = "MS verification TXT record present in DNS"
            } else {
                Write-Log "    ✗ No MS verification record found in DNS" "WARN"
                $result.DNSVerifyRecordFound = $false
                $result.DNSVerifyRecord = $dnsCheck.Record
                $result.ReadinessStatus = "Not Ready"
                $result.Recommendation = "Add domain to destination tenant to get verification record, then add TXT record to DNS"

                if ($dnsCheck.Record) {
                    $result.Notes = "DNS issue: $($dnsCheck.Record)"
                } else {
                    $result.Notes = "No MS verification TXT record found in public DNS"
                }
            }
        }

        $allResults.Add($result)
    }
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

# Generate summary
$totalDomains = $allResults.Count
$alreadyAdded = ($allResults | Where-Object { $_.ReadinessStatus -eq "Already Added" }).Count
$readyToAdd = ($allResults | Where-Object { $_.ReadinessStatus -eq "Ready to Add" }).Count
$notReady = ($allResults | Where-Object { $_.ReadinessStatus -eq "Not Ready" }).Count
$needsVerification = ($allResults | Where-Object { $_.ReadinessStatus -eq "Added - Needs Verification" }).Count
$errors = ($allResults | Where-Object { $_.ReadinessStatus -eq "Error" }).Count

$summary = @"
=== DOMAIN MIGRATION READINESS SUMMARY ===
Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

TOTAL MIGRATION PAIRS PROCESSED: $($migrationPairs.Count)
TOTAL DOMAINS CHECKED: $totalDomains

STATUS BREAKDOWN:
  ✓ Already Added to Destination: $alreadyAdded
  ⚠ Added - Needs Verification: $needsVerification
  ✓ Ready to Add (DNS verified): $readyToAdd
  ✗ Not Ready (DNS not configured): $notReady
  ⚠ Errors: $errors

=== DETAILED STATUS BY TENANT ===

"@

# Group by tenant name
$groupedResults = $allResults | Group-Object -Property TenantName

foreach ($group in $groupedResults) {
    $summary += "`n--- $($group.Name) ---`n"

    # Show source -> destination mapping
    $firstResult = $group.Group[0]
    $summary += "  Source: $($firstResult.SourceTenant)`n"
    $summary += "  Destination: $($firstResult.DestinationTenant)`n"
    $summary += "  Domains: $($group.Count)`n`n"

    # Already Added
    $addedDomains = $group.Group | Where-Object { $_.ReadinessStatus -eq "Already Added" }
    if ($addedDomains) {
        $summary += "  ALREADY IN DESTINATION ($($addedDomains.Count)):`n"
        foreach ($r in $addedDomains) {
            $summary += "    ✓ $($r.Domain) - $($r.Notes)`n"
        }
        $summary += "`n"
    }

    # Needs Verification
    $needsVerifyDomains = $group.Group | Where-Object { $_.ReadinessStatus -eq "Added - Needs Verification" }
    if ($needsVerifyDomains) {
        $summary += "  ADDED - NEEDS VERIFICATION ($($needsVerifyDomains.Count)):`n"
        foreach ($r in $needsVerifyDomains) {
            $summary += "    ⚠ $($r.Domain)`n"
        }
        $summary += "`n"
    }

    # Ready to Add
    $readyDomains = $group.Group | Where-Object { $_.ReadinessStatus -eq "Ready to Add" }
    if ($readyDomains) {
        $summary += "  READY TO ADD ($($readyDomains.Count)):`n"
        foreach ($r in $readyDomains) {
            $summary += "    ✓ $($r.Domain) - DNS: $($r.DNSVerifyRecord)`n"
        }
        $summary += "`n"
    }

    # Not Ready
    $notReadyDomains = $group.Group | Where-Object { $_.ReadinessStatus -eq "Not Ready" }
    if ($notReadyDomains) {
        $summary += "  NOT READY ($($notReadyDomains.Count)):`n"
        foreach ($r in $notReadyDomains) {
            $summary += "    ✗ $($r.Domain) - $($r.Notes)`n"
        }
        $summary += "`n"
    }

    # Errors
    $errorDomains = $group.Group | Where-Object { $_.ReadinessStatus -eq "Error" }
    if ($errorDomains) {
        $summary += "  ERRORS ($($errorDomains.Count)):`n"
        foreach ($r in $errorDomains) {
            $summary += "    ⚠ $($r.Notes)`n"
        }
        $summary += "`n"
    }
}

$summary += "`n=== NEXT STEPS ===`n"
$summary += "1. For 'Ready to Add' domains: Add to destination tenant and verify`n"
$summary += "2. For 'Not Ready' domains: Add to destination tenant, get TXT record, update DNS`n"
$summary += "3. For 'Added - Needs Verification': Complete verification process`n"
$summary += "4. For 'Already Added': No action needed`n"

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

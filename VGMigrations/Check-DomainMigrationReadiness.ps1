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

        # Check column N (column 14) for "Yes" - indicates migration complete
        if (-not $ProcessAll) {
            $columnNValue = $usedRange.Cells.Item($row, 14).Text.Trim()

            # Skip if column N has "Yes" (migration complete)
            if ($columnNValue.ToLower() -eq "yes") {
                Write-Log "Row $row : [$tenantName] - Skipping (Migration Complete - Column N = Yes)"
                continue
            }
        }

        # Read column P (source tenant - column 16) and Q (destination tenant - column 17)
        # Column P = Domain names (comma-separated)
        # Column Q = Target tenant ID
        # Column R = Login credentials (optional)
        $domainNames = $usedRange.Cells.Item($row, 16).Text.Trim()
        $destinationTenant = $usedRange.Cells.Item($row, 17).Text.Trim()
        $loginCredential = $usedRange.Cells.Item($row, 18).Text.Trim()

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
        ExpectedVerifyRecord = ""
        DNSVerifyRecord = ""
        ReadinessStatus = ""
        Notes = ""
    }

    # STEP 1: Connect to target tenant to get the expected verification record
    Write-Log "Connecting to destination tenant: $($pair.DestinationTenant)"
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

        if (-not [string]::IsNullOrWhiteSpace($pair.LoginCredential)) {
            Write-Log "Using account: $($pair.LoginCredential)"
            Connect-MgGraph -TenantId $pair.DestinationTenant -Scopes "Domain.Read.All" -UseDeviceCode -NoWelcome -AccountId $pair.LoginCredential
        } else {
            Connect-MgGraph -TenantId $pair.DestinationTenant -Scopes "Domain.Read.All" -UseDeviceCode -NoWelcome
        }

        # Check if domain is already added to target
        Write-Log "Checking if domain exists in target tenant..."
        $targetDomains = Get-MgDomain
        $targetDomain = $targetDomains | Where-Object { $_.Id -eq $domainToCheck }

        if ($targetDomain) {
            Write-Log "  ℹ Domain is added to target tenant (Status: $($targetDomain.AuthenticationType))"

            # Get the verification DNS record from the target tenant
            $verificationRecords = Get-MgDomainServiceConfigurationRecord -DomainId $domainToCheck | Where-Object { $_.RecordType -eq 'Txt' -and $_.Label -eq '@' }

            if ($verificationRecords) {
                $expectedRecord = $verificationRecords[0].Text
                Write-Log "  Expected verify record: $expectedRecord"
                $result.ExpectedVerifyRecord = $expectedRecord
            } else {
                Write-Log "  ⚠ Could not retrieve expected verification record from target" "WARN"
                $result.ExpectedVerifyRecord = "Unable to retrieve from target"
            }
        } else {
            Write-Log "  ℹ Domain NOT yet added to target tenant"
            $result.Notes = "Domain not added to target tenant yet - add domain first to get verification record"
            $result.ReadinessStatus = "Not Ready"
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            $allResults.Add($result)
            continue
        }

        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

    } catch {
        Write-Log "ERROR connecting to target tenant: $_" "ERROR"
        $result.ReadinessStatus = "Error"
        $result.Notes = "Failed to connect to target tenant: $_"
        $allResults.Add($result)
        continue
    }

    # STEP 2: Check if the EXPECTED verification record is in DNS
    Write-Log "Checking DNS for expected MS verification record..."
    $dnsCheck = Get-DNSVerificationRecord -Domain $domainToCheck

    if ($dnsCheck.Found) {
        # Check if the DNS record matches what the target expects
        if ($dnsCheck.AllRecords -match [regex]::Escape($result.ExpectedVerifyRecord)) {
            Write-Log "  ✓ Correct MS verification record found in DNS" "SUCCESS"
            $result.DNSVerifyRecord = $result.ExpectedVerifyRecord
            $result.ReadinessStatus = "Ready"
            $result.Notes = "Domain added to target and DNS verification record matches"
        } else {
            Write-Log "  ✗ MS verification record found but DOES NOT MATCH target expectation" "WARN"
            Write-Log "    Expected: $($result.ExpectedVerifyRecord)" "WARN"
            Write-Log "    Found in DNS: $($dnsCheck.AllRecords)" "WARN"
            $result.DNSVerifyRecord = $dnsCheck.Record
            $result.ReadinessStatus = "Not Ready"
            $result.Notes = "DNS has old MS record from previous migration - update with new record"
        }
    } else {
        Write-Log "  ✗ Expected MS verification record NOT found in DNS" "WARN"
        $result.DNSVerifyRecord = ""
        $result.ReadinessStatus = "Not Ready"
        $result.Notes = "Domain added to target but DNS verification record not configured"
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
            $summary += "    $($r.DomainName) - $($r.DNSVerifyRecord)`n"
        }
        $summary += "`n"
    }

    # Not Ready
    $notReadyDomains = $group.Group | Where-Object { $_.ReadinessStatus -eq "Not Ready" }
    if ($notReadyDomains) {
        $summary += "  NOT READY ($($notReadyDomains.Count)):`n"
        foreach ($r in $notReadyDomains) {
            $summary += "    $($r.DomainName) - $($r.Notes)`n"
        }
        $summary += "`n"
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

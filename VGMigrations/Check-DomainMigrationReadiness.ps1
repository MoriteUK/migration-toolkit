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
        # Query TXT records using multiple public DNS servers for reliability
        # Try Google DNS first, fallback to Cloudflare if needed
        $dnsServers = @('8.8.8.8', '1.1.1.1')
        $txtRecords = $null
        $dnsServerUsed = $null

        foreach ($dnsServer in $dnsServers) {
            try {
                Write-Log "    DNS: Querying $Domain via $dnsServer..."
                $txtRecords = Resolve-DnsName -Name $Domain -Type TXT -Server $dnsServer -ErrorAction Stop
                $dnsServerUsed = $dnsServer
                Write-Log "    DNS: Successfully queried via $dnsServer" "SUCCESS"
                break
            } catch {
                Write-Log "    DNS: Failed to query via $dnsServer - $($_.Exception.Message)" "WARN"
                continue
            }
        }

        if (-not $txtRecords) {
            Write-Log "    DNS: No TXT records found (tried all public DNS servers)" "WARN"
            return @{
                Found = $false
                Message = "No TXT records found in DNS (queried via public DNS: $($dnsServers -join ', '))"
                ActualValue = $null
            }
        }

        Write-Log "    DNS: Found $($txtRecords.Count) TXT record(s) via $dnsServerUsed"
        Write-Log "    DNS: Detailed record listing:"

        # Log all TXT records found for detailed diagnostics
        $allRecordStrings = @()
        Write-Log "    DNS: Processing $($txtRecords.Count) record objects..."
        foreach ($record in $txtRecords) {
            # Handle both single strings and arrays
            if ($record.Strings) {
                if ($record.Strings -is [array]) {
                    $recordString = $record.Strings -join ''
                    Write-Log "      [Array] Raw: $($record.Strings -join ' | ') → Joined: $recordString"
                } else {
                    $recordString = $record.Strings
                    Write-Log "      [String] Raw: $recordString"
                }
                $allRecordStrings += $recordString
            } else {
                Write-Log "      [Empty] Record has no Strings property" "WARN"
            }
        }
        Write-Log "    DNS: Extracted $($allRecordStrings.Count) record string(s) for comparison"

        # If checking for any MS= record (domain in another tenant scenario)
        if ($CheckForAnyMSRecord) {
            $msRecords = @()
            foreach ($recordString in $allRecordStrings) {
                if ($recordString -match '^MS=ms\d+') {
                    Write-Log "    DNS: MS= verification record FOUND: $recordString" "SUCCESS"
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
                Write-Log "    DNS: No MS= verification records found in the $($allRecordStrings.Count) TXT records" "WARN"
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
        $expectedTrimmed = $ExpectedValue.Trim()

        Write-Log "    DNS: Comparing against expected value: '$expectedTrimmed'"

        foreach ($recordString in $allRecordStrings) {
            $recordTrimmed = $recordString.Trim()

            if ($recordTrimmed -match '^MS=ms\d+') {
                $foundMSRecords += $recordTrimmed
                Write-Log "      → Found MS= record: '$recordTrimmed'"
            }

            # Case-insensitive comparison with trimmed strings
            if ($recordTrimmed -eq $expectedTrimmed) {
                Write-Log "    DNS: ✓✓✓ EXACT MATCH FOUND ✓✓✓: '$recordTrimmed'" "SUCCESS"
                $matchFound = $true
                break
            } else {
                # Debug: show why it didn't match
                if ($recordTrimmed -match '^MS=ms\d+') {
                    Write-Log "      ✗ No match: '$recordTrimmed' vs '$expectedTrimmed'"
                    Write-Log "        Length: $($recordTrimmed.Length) vs $($expectedTrimmed.Length)"
                    Write-Log "        Bytes: $([BitConverter]::ToString([System.Text.Encoding]::UTF8.GetBytes($recordTrimmed)))"
                    Write-Log "        Expected bytes: $([BitConverter]::ToString([System.Text.Encoding]::UTF8.GetBytes($expectedTrimmed)))"
                }
            }
        }

        if ($matchFound) {
            Write-Log "    DNS: Returning Found=TRUE" "SUCCESS"
            return @{
                Found = $true
                Message = "Correct verification record found in DNS"
                ActualValue = $ExpectedValue
            }
        } else {
            if ($foundMSRecords.Count -gt 0) {
                Write-Log "    DNS: Found MS= record(s) but none match expected value" "WARN"
                Write-Log "    DNS: Expected: '$expectedTrimmed'" "WARN"
                Write-Log "    DNS: Found: $($foundMSRecords -join ', ')" "WARN"
                Write-Log "    DNS: Returning Found=FALSE with ActualValue" "WARN"
                return @{
                    Found = $false
                    Message = "MS= record exists but does not match expected value (Found: $($foundMSRecords[0]))"
                    ActualValue = $foundMSRecords[0]
                }
            } else {
                Write-Log "    DNS: No MS= verification record found" "WARN"
                Write-Log "    DNS: Returning Found=FALSE with null" "WARN"
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

                # IMPORTANT: Check if domain was added but is unverified
                # This happens when the domain is still active in the source tenant
                Write-Log "  Domain verification status: IsVerified=$($targetDomain.IsVerified)"
                if (-not $targetDomain.IsVerified) {
                    Write-Log "  ⚠ Domain added but NOT verified (likely still active in source tenant)" "WARN"
                    Write-Log "  This is expected pre-cutover - will check DNS for MS= record"
                    Write-Log "  Setting domainInAnotherTenant=TRUE for pre-cutover validation"
                    # Treat as pre-cutover scenario even though New-MgDomain succeeded
                    $domainInAnotherTenant = $true
                } else {
                    Write-Log "  ✓ Domain is verified - normal validation will proceed"
                }
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
            } else {
                # Domain exists but is NOT verified - likely still in source tenant
                Write-Log "  ⚠ Domain exists in target but is NOT verified (likely still active in source tenant)" "WARN"
                Write-Log "  This is expected pre-cutover - will check DNS for MS= record"
                Write-Log "  Setting domainInAnotherTenant=TRUE for pre-cutover validation"
                $domainInAnotherTenant = $true
            }
        }

        # Get the verification DNS record from the target tenant (if possible)
        if ($domainInAnotherTenant -and $cannotGetVerificationRecord) {
            Write-Log "  ⚠ Cannot retrieve exact verification record - domain blocked by source tenant"
            Write-Log "  Will check DNS for any MS= verification record instead"
            $result.DNS_RecordType = "TXT"
            $result.DNS_Label = "@"
            $result.DNS_Value = "Unknown - domain in source tenant"
            # We'll check DNS for ANY MS= record below
        } elseif ($domainInAnotherTenant -and -not $cannotGetVerificationRecord) {
            Write-Log "  ℹ Domain added to target but unverified (still in source tenant)"
            Write-Log "  Retrieving verification record AND checking DNS for any MS= record"
            # Try to get the target's expected record, but also accept any MS= record
            # Fall through to normal verification record retrieval
        }

        if (-not $cannotGetVerificationRecord) {
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
    Write-Log "DEBUG: domainInAnotherTenant=$domainInAnotherTenant, cannotGetVerificationRecord=$cannotGetVerificationRecord"
    Write-Log "DEBUG: DNS_Value='$($result.DNS_Value)'"

    if ($domainInAnotherTenant) {
        # Domain is in another tenant - check for ANY MS= record first
        Write-Log "  ═══ PRE-CUTOVER SCENARIO DETECTED ═══" "INFO"
        Write-Log "  Domain still in source tenant"

        # If we have the target's expected value, try exact match first
        if ($result.DNS_Value -and $result.DNS_Value -ne "Unknown - domain in source tenant") {
            Write-Log "  Attempting exact match against target's expected value: $($result.DNS_Value)"
            $exactCheck = Check-DNSForRecord -Domain $domainToCheck -ExpectedValue $result.DNS_Value

            if ($exactCheck.Found) {
                Write-Log "  ✓ PERFECT MATCH: DNS record matches target tenant's expected value!" "SUCCESS"
                $result.ReadinessStatus = "Ready - Pre-Cutover (Exact Match)"
                $result.Notes = "DNS verification record matches target tenant exactly. Domain still in source tenant - ready for cutover."
                $result.DNSVerifyRecord = $exactCheck.ActualValue
                $allResults.Add($result)
                continue
            } else {
                Write-Log "  ℹ No exact match - checking for any MS= record instead"
            }
        }

        # Check for ANY MS= record (from source tenant)
        Write-Log "  Checking for any MS= verification record..."
        $dnsCheck = Check-DNSForRecord -Domain $domainToCheck -ExpectedValue "" -CheckForAnyMSRecord

        if ($dnsCheck.Found) {
            Write-Log "  ✓ $($dnsCheck.Message)" "SUCCESS"

            # Compare found record with target's expected value
            if ($result.DNS_Value -and $result.DNS_Value -ne "Unknown - domain in source tenant" -and $dnsCheck.ActualValue -ne $result.DNS_Value) {
                Write-Log "  ⚠ DNS has different MS= record than target expects" "WARN"
                Write-Log "    Found in DNS: $($dnsCheck.ActualValue)"
                Write-Log "    Target expects: $($result.DNS_Value)"
                $result.ReadinessStatus = "Ready - Pre-Cutover (Source Record)"
                $result.Notes = "MS= record found in DNS (from source tenant). Target expects different value: $($result.DNS_Value). Update DNS before cutover or domain verification may fail after transfer."
                $result.DNS_Value = "Target: $($result.DNS_Value) | Current: $($dnsCheck.ActualValue)"
            } else {
                $result.ReadinessStatus = "Ready - Pre-Cutover"
                $result.Notes = "MS= verification record found in DNS. Domain currently in source tenant - record will be validated on cutover."
            }
            $result.DNSVerifyRecord = $dnsCheck.ActualValue
        } else {
            Write-Log "  ✗ $($dnsCheck.Message)" "WARN"
            $result.ReadinessStatus = "Not Ready"
            $result.Notes = "No MS= verification record found in DNS. Add verification record before cutover. Target expects: $($result.DNS_Value)"
        }
    } else {
        # Normal flow - validate exact expected value
        Write-Log "  ═══ NORMAL SCENARIO (NOT PRE-CUTOVER) ═══" "INFO"
        Write-Log "  Validating exact match against target's expected value"
        Write-Log "  Expected value: '$($result.DNS_Value)'"
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

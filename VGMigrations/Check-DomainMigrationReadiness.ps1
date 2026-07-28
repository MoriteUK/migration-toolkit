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

function Get-OrCreateAppRegistration {
    param(
        [string]$TenantId,
        [string]$TenantName,
        [string]$AppName,
        [string]$LoginHint
    )

    Write-Log "Checking app registration for $TenantName..."

    try {
        # Check if we already have stored credentials
        $credFile = Join-Path $CredentialStorePath "domaincheck_$($TenantId).json"
        if (Test-Path $credFile) {
            $creds = Get-Content $credFile | ConvertFrom-Json
            Write-Log "Using stored credentials (App: $($creds.AppId))"
            return $creds
        }

        # Need to create app registration - connect interactively
        Write-Log "First time setup - please authenticate as admin"
        if ($LoginHint) {
            Write-Log "Use account: $LoginHint"
        }

        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Connect-MgGraph -TenantId $TenantId -Scopes "Application.ReadWrite.All" -NoWelcome

        # Check if app exists
        $existingApp = Get-MgApplication -Filter "displayName eq '$AppName'" -ErrorAction SilentlyContinue

        if ($existingApp) {
            Write-Log "App exists: $($existingApp.AppId)"
            $appId = $existingApp.AppId
            $objectId = $existingApp.Id
        } else {
            Write-Log "Creating app registration..."

            # Required permissions for domain checking
            $requiredResourceAccess = @(
                @{
                    ResourceAppId = "00000003-0000-0000-c000-000000000000" # MS Graph
                    ResourceAccess = @(
                        @{ Id = "dbb9058a-0e50-45d7-ae91-66909b5d4664"; Type = "Role" } # Domain.Read.All
                    )
                }
            )

            $newApp = New-MgApplication -DisplayName $AppName -SignInAudience "AzureADMyOrg" -RequiredResourceAccess $requiredResourceAccess
            $appId = $newApp.AppId
            $objectId = $newApp.Id
            Write-Log "App created: $appId"
            Start-Sleep -Seconds 3
        }

        # Create client secret
        Write-Log "Creating client secret..."
        $passwordCred = @{
            DisplayName = "Auto-$(Get-Date -Format 'yyyyMMdd')"
            EndDateTime = (Get-Date).AddYears(2)
        }
        $secret = Add-MgApplicationPassword -ApplicationId $objectId -PasswordCredential $passwordCred
        $clientSecret = $secret.SecretText

        # Ensure service principal exists
        $sp = Get-MgServicePrincipal -Filter "appId eq '$appId'" -ErrorAction SilentlyContinue
        if (-not $sp) {
            $sp = New-MgServicePrincipal -AppId $appId
            Start-Sleep -Seconds 2
        }

        # Grant admin consent
        Write-Log "Granting admin consent..."
        $graphSP = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
        try {
            New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -BodyParameter @{
                PrincipalId = $sp.Id
                ResourceId = $graphSP.Id
                AppRoleId = "dbb9058a-0e50-45d7-ae91-66909b5d4664" # Domain.Read.All
            } -ErrorAction SilentlyContinue | Out-Null
        } catch {}

        # Store credentials
        $credObject = @{
            TenantId = $TenantId
            TenantName = $TenantName
            AppId = $appId
            ClientSecret = $clientSecret
            SecretExpires = $secret.EndDateTime
            CreatedDate = Get-Date
        }
        $credObject | ConvertTo-Json | Out-File -FilePath $credFile -Encoding UTF8 -Force
        Write-Log "Credentials saved for future runs" "SUCCESS"

        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        return [PSCustomObject]$credObject

    } catch {
        Write-Log "Failed to create app registration: $_" "ERROR"
        throw
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
        $secureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential($AppId, $secureSecret)
        Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $credential -NoWelcome
    } catch {
        Write-Log "Failed to connect with app credentials: $_" "ERROR"
        throw
    }
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
        DNS_RecordType = ""
        DNS_Label = ""
        DNS_Value = ""
        DNSVerifyRecord = ""
        ReadinessStatus = ""
        Notes = ""
    }

    # STEP 1: Get or create app registration for this tenant
    Write-Log "Connecting to destination tenant: $($pair.DestinationTenant)"
    try {
        # Get app credentials (creates app on first run, reuses thereafter)
        $appCreds = Get-OrCreateAppRegistration -TenantId $pair.DestinationTenant -TenantName $pair.TenantName -AppName $AppName -LoginHint $pair.LoginCredential

        # Connect using app credentials
        Connect-WithAppCredentials -TenantId $appCreds.TenantId -AppId $appCreds.AppId -ClientSecret $appCreds.ClientSecret

        # Check if domain is already added to target
        Write-Log "Checking if domain exists in target tenant..."
        $targetDomains = Get-MgDomain
        $targetDomain = $targetDomains | Where-Object { $_.Id -eq $domainToCheck }

        if ($targetDomain) {
            Write-Log "  ℹ Domain is added to target tenant (Verified: $($targetDomain.IsVerified))"

            # Get the verification DNS record from the target tenant
            $verificationRecords = Get-MgDomainServiceConfigurationRecord -DomainId $domainToCheck

            # Find the TXT record for verification
            $txtRecord = $verificationRecords | Where-Object {
                $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.domainDnsTxtRecord'
            } | Select-Object -First 1

            if ($txtRecord) {
                $recordType = "TXT"
                $recordLabel = if ($txtRecord.Label) { $txtRecord.Label } else { "@" }
                $recordValue = $txtRecord.Text

                Write-Log "  DNS Record to configure:"
                Write-Log "    Type: $recordType"
                Write-Log "    Label: $recordLabel"
                Write-Log "    Value: $recordValue"

                $result.DNS_RecordType = $recordType
                $result.DNS_Label = $recordLabel
                $result.DNS_Value = $recordValue
                $result.DNSVerifyRecord = $recordValue
            } else {
                Write-Log "  ⚠ Could not retrieve DNS verification record from target" "WARN"
                $result.DNS_RecordType = "TXT"
                $result.DNS_Label = "Unable to retrieve"
                $result.DNS_Value = "Unable to retrieve from target"
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
        if ($dnsCheck.AllRecords -match [regex]::Escape($result.DNS_Value)) {
            Write-Log "  ✓ Correct MS verification record found in DNS" "SUCCESS"
            $result.ReadinessStatus = "Ready"
            $result.Notes = "DNS verification record matches target tenant"
        } else {
            Write-Log "  ✗ MS verification record found but DOES NOT MATCH target expectation" "WARN"
            Write-Log "    Expected: $($result.DNS_Value)" "WARN"
            Write-Log "    Found in DNS: $($dnsCheck.AllRecords)" "WARN"
            $result.ReadinessStatus = "Not Ready"
            $result.Notes = "DNS has old record - Add new TXT record: Label=$($result.DNS_Label) Value=$($result.DNS_Value)"
        }
    } else {
        Write-Log "  ✗ Expected MS verification record NOT found in DNS" "WARN"
        $result.ReadinessStatus = "Not Ready"
        $result.Notes = "Add DNS TXT record: Label=$($result.DNS_Label) Value=$($result.DNS_Value)"
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

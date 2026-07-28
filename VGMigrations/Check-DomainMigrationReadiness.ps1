#Requires -Version 7.0
<#
.SYNOPSIS
    Checks if domains are added to target tenant and if DNS verification records are ready
.DESCRIPTION
    For each domain in the source tenant:
    1. Checks if domain exists in target tenant
    2. If not, checks if MS-verify TXT record exists in public DNS
    3. Reports readiness status for each domain
.PARAMETER SourceTenantId
    Source tenant ID or domain (where domains currently exist)
.PARAMETER TargetTenantId
    Target tenant ID or domain (where domains should be migrated to)
.PARAMETER DomainList
    Optional: Specific domains to check. If omitted, checks all domains from source tenant.
.PARAMETER OutputPath
    Path where the report will be saved
.EXAMPLE
    .\Check-DomainMigrationReadiness.ps1 -SourceTenantId "contoso.onmicrosoft.com" -TargetTenantId "fabrikam.onmicrosoft.com"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$SourceTenantId,

    [Parameter(Mandatory=$true)]
    [string]$TargetTenantId,

    [string[]]$DomainList,

    [string]$OutputPath = "C:\Users\Andy White\Volaris Group\GRP Data Security (Volaris Consolidated) - M365 Migrations"
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
Write-Log "Source Tenant: $SourceTenantId"
Write-Log "Target Tenant: $TargetTenantId"

# Check if Microsoft Graph module is available
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    Write-Log "ERROR: Microsoft.Graph module not installed. Install with: Install-Module Microsoft.Graph -Scope CurrentUser" "ERROR"
    exit 1
}

# Connect to source tenant to get domains
Write-Log ""
Write-Log "=== Connecting to Source Tenant ==="
try {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Connect-MgGraph -TenantId $SourceTenantId -Scopes "Domain.Read.All" -NoWelcome
    Write-Log "Connected to source tenant"

    # Get source tenant info
    $sourceOrg = Get-MgOrganization
    $sourceTenantName = $sourceOrg.DisplayName
    Write-Log "Source tenant name: $sourceTenantName"

    # Get domains from source tenant
    if ($DomainList) {
        Write-Log "Using provided domain list: $($DomainList -join ', ')"
        $sourceDomains = $DomainList | ForEach-Object {
            [PSCustomObject]@{
                Id = $_
                IsDefault = $false
                IsVerified = $true
            }
        }
    } else {
        Write-Log "Retrieving all domains from source tenant..."
        $sourceDomains = Get-MgDomain | Where-Object {
            # Skip .onmicrosoft.com domains
            $_.Id -notmatch '\.onmicrosoft\.com$'
        }
        Write-Log "Found $($sourceDomains.Count) custom domains in source tenant"
    }

    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

} catch {
    Write-Log "ERROR connecting to source tenant: $_" "ERROR"
    exit 1
}

# Connect to target tenant
Write-Log ""
Write-Log "=== Connecting to Target Tenant ==="
try {
    Connect-MgGraph -TenantId $TargetTenantId -Scopes "Domain.Read.All" -NoWelcome
    Write-Log "Connected to target tenant"

    # Get target tenant info
    $targetOrg = Get-MgOrganization
    $targetTenantName = $targetOrg.DisplayName
    Write-Log "Target tenant name: $targetTenantName"

    # Get domains from target tenant
    Write-Log "Retrieving domains from target tenant..."
    $targetDomains = Get-MgDomain
    $targetDomainNames = $targetDomains | Select-Object -ExpandProperty Id
    Write-Log "Found $($targetDomains.Count) total domains in target tenant"

    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

} catch {
    Write-Log "ERROR connecting to target tenant: $_" "ERROR"
    exit 1
}

# Check each domain
Write-Log ""
Write-Log "=== Checking Domain Readiness ==="
$results = [System.Collections.Generic.List[object]]::new()

foreach ($domain in $sourceDomains) {
    $domainName = $domain.Id
    Write-Log ""
    Write-Log "Checking: $domainName"

    $result = [PSCustomObject]@{
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
        Write-Log "  ✓ Domain exists in target tenant" "SUCCESS"
        $result.InTargetTenant = $true
        $result.ReadinessStatus = "Already Added"
        $result.Recommendation = "Domain is already in target tenant"

        # Check verification status in target
        $targetDomain = $targetDomains | Where-Object { $_.Id -eq $domainName }
        if ($targetDomain.IsVerified) {
            $result.Notes = "Verified in target tenant"
        } else {
            $result.Notes = "Added but NOT verified in target tenant"
            $result.ReadinessStatus = "Added - Needs Verification"
            $result.Recommendation = "Complete domain verification in target tenant"
        }
    } else {
        Write-Log "  ✗ Domain NOT in target tenant" "WARN"
        $result.InTargetTenant = $false

        # Check DNS for verification record
        Write-Log "  Checking DNS for MS verification record..."
        $dnsCheck = Get-DNSVerificationRecord -Domain $domainName

        if ($dnsCheck.Found) {
            Write-Log "  ✓ MS verification record found in DNS: $($dnsCheck.Record)" "SUCCESS"
            $result.DNSVerifyRecordFound = $true
            $result.DNSVerifyRecord = $dnsCheck.Record
            $result.ReadinessStatus = "Ready to Add"
            $result.Recommendation = "Add domain to target tenant and verify"
            $result.Notes = "MS verification TXT record present in DNS"
        } else {
            Write-Log "  ✗ No MS verification record found in DNS" "WARN"
            $result.DNSVerifyRecordFound = $false
            $result.DNSVerifyRecord = $dnsCheck.Record
            $result.ReadinessStatus = "Not Ready"
            $result.Recommendation = "Add domain to target tenant to get verification record, then add TXT record to DNS"

            if ($dnsCheck.Record) {
                $result.Notes = "DNS issue: $($dnsCheck.Record)"
            } else {
                $result.Notes = "No MS verification TXT record found in public DNS"
            }
        }
    }

    $results.Add($result)
}

# Generate report
Write-Log ""
Write-Log "=== Generating Report ==="

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$reportFile = Join-Path $OutputPath "DomainMigrationReadiness_$timestamp.csv"
$summaryFile = Join-Path $OutputPath "DomainMigrationReadiness_$timestamp.txt"

# Export detailed CSV
$results | Export-Csv -Path $reportFile -NoTypeInformation -Encoding UTF8
Write-Log "Detailed report saved: $reportFile"

# Generate summary
$alreadyAdded = ($results | Where-Object { $_.ReadinessStatus -eq "Already Added" }).Count
$readyToAdd = ($results | Where-Object { $_.ReadinessStatus -eq "Ready to Add" }).Count
$notReady = ($results | Where-Object { $_.ReadinessStatus -eq "Not Ready" }).Count
$needsVerification = ($results | Where-Object { $_.ReadinessStatus -eq "Added - Needs Verification" }).Count

$summary = @"
=== DOMAIN MIGRATION READINESS SUMMARY ===
Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

SOURCE TENANT: $sourceTenantName ($SourceTenantId)
TARGET TENANT: $targetTenantName ($TargetTenantId)

TOTAL DOMAINS CHECKED: $($results.Count)

STATUS BREAKDOWN:
  ✓ Already Added to Target: $alreadyAdded
  ⚠ Added - Needs Verification: $needsVerification
  ✓ Ready to Add (DNS verified): $readyToAdd
  ✗ Not Ready (DNS not configured): $notReady

=== DETAILED STATUS ===

"@

# Already Added
if ($alreadyAdded -gt 0) {
    $summary += "`n--- ALREADY IN TARGET TENANT ($alreadyAdded) ---`n"
    foreach ($r in $results | Where-Object { $_.ReadinessStatus -eq "Already Added" }) {
        $summary += "  ✓ $($r.Domain)`n"
        $summary += "    Status: $($r.Notes)`n"
    }
}

# Needs Verification
if ($needsVerification -gt 0) {
    $summary += "`n--- ADDED BUT NEEDS VERIFICATION ($needsVerification) ---`n"
    foreach ($r in $results | Where-Object { $_.ReadinessStatus -eq "Added - Needs Verification" }) {
        $summary += "  ⚠ $($r.Domain)`n"
        $summary += "    Action: Complete domain verification in target tenant`n"
    }
}

# Ready to Add
if ($readyToAdd -gt 0) {
    $summary += "`n--- READY TO ADD TO TARGET ($readyToAdd) ---`n"
    foreach ($r in $results | Where-Object { $_.ReadinessStatus -eq "Ready to Add" }) {
        $summary += "  ✓ $($r.Domain)`n"
        $summary += "    DNS Record: $($r.DNSVerifyRecord)`n"
        $summary += "    Action: Add domain to target tenant and verify`n"
    }
}

# Not Ready
if ($notReady -gt 0) {
    $summary += "`n--- NOT READY ($notReady) ---`n"
    foreach ($r in $results | Where-Object { $_.ReadinessStatus -eq "Not Ready" }) {
        $summary += "  ✗ $($r.Domain)`n"
        $summary += "    Issue: $($r.Notes)`n"
        $summary += "    Action: $($r.Recommendation)`n"
    }
}

$summary += "`n=== NEXT STEPS ===`n"
$summary += "1. For 'Ready to Add' domains: Add to target tenant and verify`n"
$summary += "2. For 'Not Ready' domains: Add to target tenant, get TXT record, update DNS`n"
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

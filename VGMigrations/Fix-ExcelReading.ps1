# Quick script to replace COM Excel reading with ImportExcel module
$scriptPath = "C:\toolkit\VGMigrations\Check-DomainMigrationReadiness.ps1"
$content = Get-Content $scriptPath -Raw

# Find and replace the entire COM Excel section
$oldPattern = @'
# Read Excel file using COM
Write-Log "Opening Excel file..."
\$excel = \$null
\$workbook = \$null
\$migrationPairs = \[System\.Collections\.Generic\.List\[object\]\]::new\(\)

try \{
    \$excel = New-Object -ComObject Excel\.Application
'@

$newCode = @'
# Check if ImportExcel module is available
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Log "ImportExcel module not found. Installing..." "WARN"
    try {
        Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber
        Write-Log "ImportExcel module installed" "SUCCESS"
    } catch {
        Write-Log "ERROR: Failed to install ImportExcel module: $_" "ERROR"
        Write-Log "Install manually: Install-Module ImportExcel -Scope CurrentUser" "ERROR"
        exit 1
    }
}

# Read Excel file using ImportExcel (no Excel/COM needed)
Write-Log "Opening Excel file..."
$migrationPairs = [System.Collections.Generic.List[object]]::new()

try {
    Write-Log "Reading with ImportExcel module..."
    $excelData = Import-Excel -Path $TenantIDsPath -NoHeader -DataOnly

    if (-not $excelData) {
        throw "Could not read Excel file"
    }

    Write-Log "Found $($excelData.Count) rows"

    # Skip header row (row 0), process data starting from row 1
    for ($i = 1; $i -lt $excelData.Count; $i++) {
        $row = $excelData[$i]
        $rowNum = $i + 1

        # Get values by column position (P1=A, P2=B, etc.)
        $tenantName = if ($row.P1) { $row.P1.ToString().Trim() } else { "" }

        if ([string]::IsNullOrWhiteSpace($tenantName)) { continue }

        if (-not $ProcessAll) {
            $columnNValue = if ($row.P14) { $row.P14.ToString().Trim() } else { "" }
            if ($columnNValue.ToLower() -eq "yes") {
                Write-Log "Row $rowNum : [$tenantName] - Skipping (Migration Complete)"
                continue
            }

            $columnLValue = if ($row.P12) { $row.P12.ToString().Trim() } else { "" }
            if ($columnLValue.ToLower() -eq "yes") {
                Write-Log "Row $rowNum : [$tenantName] - Skipping (DNS Already Requested)"
                continue
            }
        }

        $primaryDomain = if ($row.P2) { $row.P2.ToString().Trim() } else { "" }
        $additionalDomains = if ($row.P16) { $row.P16.ToString().Trim() } else { "" }
        $destinationTenant = if ($row.P17) { $row.P17.ToString().Trim() } else { "" }
        $loginCredential = if ($row.P18) { $row.P18.ToString().Trim() } else { "" }

        $domainNames = ""
        if ($primaryDomain -match '^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$') {
            $domainNames = $additionalDomains
        } else {
            $domainNames = $primaryDomain
            if (-not [string]::IsNullOrWhiteSpace($additionalDomains)) {
                $domainNames += ",$additionalDomains"
            }
        }

        if ([string]::IsNullOrWhiteSpace($domainNames) -or [string]::IsNullOrWhiteSpace($destinationTenant)) {
            Write-Log "Row $rowNum : [$tenantName] - Skipping (No domain/destination)"
            continue
        }

        $domainList = $domainNames -split ',' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

        foreach ($domain in $domainList) {
            $migrationPairs.Add([PSCustomObject]@{
                TenantName = $tenantName
                DomainName = $domain
                DestinationTenant = $destinationTenant
                LoginCredential = $loginCredential
                Row = $rowNum
            })
            Write-Log "Row $rowNum : [$tenantName] - Domain: $domain -> Destination: $destinationTenant"
        }
    }
'@

Write-Host "This script needs manual editing. The file is too complex to safely auto-patch."
Write-Host "Instead, use the server update command provided."

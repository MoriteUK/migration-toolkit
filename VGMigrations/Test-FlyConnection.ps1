#Requires -Version 7.0
<#
.SYNOPSIS
    Tests connection to AvePoint Fly API
.DESCRIPTION
    Validates API credentials by attempting to connect to Fly
#>

try {
    # Get Fly API configuration
    $flyApiCfgPath = Join-Path $env:APPDATA "FlyMigration\config.json"
    if (-not (Test-Path $flyApiCfgPath)) {
        Write-Output "ERROR: Configuration not found"
        exit 1
    }

    $rawCfg = Get-Content $flyApiCfgPath -Raw | ConvertFrom-Json
    $apiUrl = $rawCfg.Url
    $clientId = $rawCfg.ClientId

    if (-not $apiUrl -or -not $clientId) {
        Write-Output "ERROR: API URL or Client ID not configured"
        exit 1
    }

    if ($rawCfg.EncSecret) {
        $secureSecret = $rawCfg.EncSecret | ConvertTo-SecureString
        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureSecret)
        $clientSecret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    } else {
        Write-Output "ERROR: Client secret not configured"
        exit 1
    }

    # Import Fly.Client module
    if (-not (Get-Module -Name Fly.Client -ListAvailable)) {
        Write-Output "ERROR: Fly.Client module not installed"
        exit 1
    }

    Import-Module Fly.Client -ErrorAction Stop

    # Test connection
    Connect-Fly -Url $apiUrl -ClientId $clientId -ClientSecret $clientSecret -ErrorAction Stop

    # If we got here, connection succeeded
    Write-Output "SUCCESS: Connected to Fly API"
    exit 0

} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
    exit 1
}

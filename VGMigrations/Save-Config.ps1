#Requires -Version 7.0
<#
.SYNOPSIS
    Saves configuration with encrypted client secret
.DESCRIPTION
    Saves app configuration to %APPDATA%\FlyMigration\config.json with proper encryption
.PARAMETER ConfigJson
    JSON string containing the configuration to save
#>
param(
    [Parameter(Mandatory=$true)]
    [string]$ConfigJson
)

try {
    $config = $ConfigJson | ConvertFrom-Json
    $configPath = Join-Path $env:APPDATA "FlyMigration\config.json"
    $configDir = Split-Path $configPath -Parent

    # Ensure directory exists
    if (-not (Test-Path $configDir)) {
        New-Item -ItemType Directory -Path $configDir -Force | Out-Null
    }

    # Load existing config if it exists
    $existingConfig = if (Test-Path $configPath) {
        Get-Content $configPath -Raw | ConvertFrom-Json
    } else {
        @{}
    }

    # Merge configs (new values override existing)
    foreach ($prop in $config.PSObject.Properties) {
        $existingConfig | Add-Member -MemberType NoteProperty -Name $prop.Name -Value $prop.Value -Force
    }

    # Encrypt client secret if provided as plain text
    if ($config.PSObject.Properties['ClientSecret'] -and $config.ClientSecret) {
        $encSecret = $config.ClientSecret | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString
        $existingConfig | Add-Member -MemberType NoteProperty -Name 'EncSecret' -Value $encSecret -Force
        # Remove plain text secret
        $existingConfig.PSObject.Properties.Remove('ClientSecret')
    }

    # Save to file
    $existingConfig | ConvertTo-Json -Depth 10 | Set-Content $configPath -Force

    # Return success
    Write-Output "SUCCESS"
} catch {
    Write-Error $_.Exception.Message
    exit 1
}

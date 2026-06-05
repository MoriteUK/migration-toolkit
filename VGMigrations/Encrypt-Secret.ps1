#Requires -Version 7.0
<#
.SYNOPSIS
    Encrypts a client secret using Windows DPAPI
.PARAMETER Secret
    The plain text secret to encrypt
#>
param(
    [Parameter(Mandatory=$true)]
    [string]$Secret
)

try {
    $encSecret = $Secret | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString
    Write-Output $encSecret
} catch {
    Write-Error $_.Exception.Message
    exit 1
}

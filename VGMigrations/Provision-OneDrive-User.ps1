#Requires -Version 7.0
<#
.SYNOPSIS
    Provisions OneDrive for one or more users.
    Run Connect-MicrosoftGraph.ps1 first — this script reuses the existing session.

.PARAMETER UPN
    One or more user principal names, e.g.
        .\Provision-OneDrive-User.ps1 -UPN alice@contoso.com
        .\Provision-OneDrive-User.ps1 -UPN alice@contoso.com, bob@contoso.com

.EXAMPLE
    # Connect once, then provision several users without re-authenticating:
    .\Connect-MicrosoftGraph.ps1
    .\Provision-OneDrive-User.ps1 -UPN alice@contoso.com
    .\Provision-OneDrive-User.ps1 -UPN bob@contoso.com, carol@contoso.com
#>
param(
    [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
    [string[]]$UPN
)

$ErrorActionPreference = 'Stop'

# Verify there is an active Graph connection
$ctx = Get-MgContext -ErrorAction SilentlyContinue
if (-not $ctx) {
    Write-Host "Not connected to Microsoft Graph." -ForegroundColor Red
    Write-Host "Run Connect-MicrosoftGraph.ps1 first, then re-run this script." -ForegroundColor Yellow
    exit 1
}

Write-Host "Connected as: $($ctx.Account)" -ForegroundColor Gray

$ok   = 0
$fail = 0

foreach ($user in $UPN) {
    $user = $user.Trim()
    if (-not $user) { continue }

    try {
        $drive = Get-MgUserDrive -UserId $user -ErrorAction Stop
        Write-Host "  OK   $user  ($($drive.DriveType))" -ForegroundColor Green
        $ok++
    } catch {
        $msg = $_.Exception.Message.Split([Environment]::NewLine)[0]
        Write-Host "  FAIL $user — $msg" -ForegroundColor Red
        $fail++
    }
}

Write-Host ""
Write-Host "Provisioned: $ok   Failed: $fail" -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Yellow' })

param(
    [Parameter(Mandatory = $true)]
    [string]$Domain,

    [int]$PollSeconds = 30
)

$mod = Get-Module -ListAvailable -Name ExchangeOnlineManagement -ErrorAction SilentlyContinue
if (-not $mod) {
    Write-Host "ExchangeOnlineManagement module not installed. Run: Install-Module ExchangeOnlineManagement -Scope CurrentUser" -ForegroundColor Red
    exit 1
}
Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue

$ctx = Get-ConnectionInformation -ErrorAction SilentlyContinue
if (-not $ctx) {
    Write-Host "Connecting to Exchange Online - sign in when the browser opens..." -ForegroundColor Cyan
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
}

$pass = 0
do {
    $pass++
    $accepted = Get-AcceptedDomain -Identity $Domain -ErrorAction SilentlyContinue
    $present = [bool]$accepted

    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Pass $pass - $Domain present in tenant: $present" -ForegroundColor Magenta

    if (-not $present) { break }

    Start-Sleep -Seconds $PollSeconds
} while ($true)

Write-Host "=== Done - $Domain is no longer present in the tenant ===" -ForegroundColor Magenta

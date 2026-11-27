<#
.SYNOPSIS
  Checks this machine has everything Invoke-FlyBriefs.ps1 needs, and reports what is missing.
  Read-only: it installs and changes nothing.
.EXAMPLE
  .\Test-FlyPrereqs.ps1
  .\Test-FlyPrereqs.ps1 -TryConnect      # also signs in to FLY and lists projects
#>
[CmdletBinding()]
param(
  [string]$ConfigPath = (Join-Path $PSScriptRoot 'fly-config.json'),
  [switch]$TryConnect
)
$results = [System.Collections.Generic.List[object]]::new()
function Add-Check($Item, [bool]$Ok, $Detail, $Fix) {
  $results.Add([pscustomobject]@{ Check = $Item; Result = $(if ($Ok) { 'OK' } else { 'MISSING' }); Detail = $Detail; Fix = $(if ($Ok) { '' } else { $Fix }) })
}

# Account and shell
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
Add-Check 'Windows account' $true "$(whoami)$(if ($isAdmin) { ' (elevated)' })" ''
Add-Check 'PowerShell 7+' ($PSVersionTable.PSVersion.Major -ge 7) "Version $($PSVersionTable.PSVersion)" 'winget install Microsoft.PowerShell'

# Fly.Client module
$mod = Get-Module -ListAvailable Fly.Client | Sort-Object Version -Descending | Select-Object -First 1
Add-Check 'Fly.Client module' ([bool]$mod) $(if ($mod) { "$($mod.Version) at $($mod.ModuleBase)" } else { 'not found for this account' }) 'Install-Module Fly.Client -Scope CurrentUser'

# Node.js and a Chromium browser for the PDF
$node = Get-Command node -ErrorAction SilentlyContinue
Add-Check 'Node.js' ([bool]$node) $(if ($node) { (& node --version) } else { 'not on PATH' }) 'winget install OpenJS.NodeJS.LTS  (then open a new window)'
$edge = @($env:EDGE_PATH, "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Google\Chrome\Application\chrome.exe") | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
Add-Check 'Edge/Chrome (PDF output)' ([bool]$edge) $(if ($edge) { $edge } else { 'not found' }) 'Install Microsoft Edge, or set $env:EDGE_PATH'

# Template files
foreach ($f in 'FLY_Executive_Brief.html', 'fly-brief.js', 'Invoke-FlyBriefs.ps1') {
  $p = Join-Path $PSScriptRoot $f
  Add-Check "File: $f" (Test-Path $p) $p 'Copy the whole Template folder to this machine'
}

# Config
$cfg = $null
if (Test-Path $ConfigPath) {
  try { $cfg = Get-Content $ConfigPath -Raw | ConvertFrom-Json; Add-Check 'fly-config.json' $true $ConfigPath '' }
  catch { Add-Check 'fly-config.json' $false "Invalid JSON: $($_.Exception.Message)" 'Fix the JSON syntax (commas, quotes, \\ in paths)' }
} else { Add-Check 'fly-config.json' $false "$ConfigPath not found" 'Copy fly-config.example.json to fly-config.json and fill it in' }
if ($cfg) {
  Add-Check 'Config: flyUrl' ($cfg.flyUrl -match '^https://graph-[a-z0-9-]+\.avepointonlineservices\.com/fly$') $cfg.flyUrl 'Use https://graph-<region>.avepointonlineservices.com/fly, e.g. graph-uk'
  Add-Check 'Config: clientId' ($cfg.clientId -match '^[0-9a-f-]{36}$') $cfg.clientId 'Paste the Application (client) ID from AOS > Administration > App registrations'
  $ph = @($cfg.projects | Where-Object { $_.project -match '^<|>$' -or -not $_.client -or $_.workload -notin 'OneDrive', 'SharePoint' })
  Add-Check 'Config: projects' ($cfg.projects -and -not $ph) "$(@($cfg.projects).Count) configured$(if ($ph) { ", $($ph.Count) still placeholders/invalid" })" 'Set each project to the exact FLY project name, or remove unused entries'
}

# Certificate with private key
$thumb = if ($cfg) { $cfg.certificateThumbprint } else { $null }
$certs = Get-ChildItem Cert:\CurrentUser\My, Cert:\LocalMachine\My -ErrorAction SilentlyContinue | Where-Object { ($thumb -and $_.Thumbprint -eq $thumb) -or $_.Subject -like '*FLY*' }
$match = $certs | Where-Object { $thumb -and $_.Thumbprint -eq $thumb } | Select-Object -First 1
if ($match) {
  Add-Check 'Certificate (config thumbprint)' $match.HasPrivateKey "$($match.PSParentPath -replace '.*::') expires $($match.NotAfter.ToString('d MMM yyyy'))$(if (-not $match.HasPrivateKey) { ' - NO private key' })" 'Import the .pfx (with private key), not the .cer'
  Add-Check 'Certificate not expired' ($match.NotAfter -gt (Get-Date)) $match.NotAfter.ToString('d MMM yyyy') 'Create a new certificate and upload its .cer to the app registration'
} else {
  $seen = if ($certs) { ($certs | ForEach-Object { "$($_.Subject) $($_.Thumbprint)" }) -join '; ' } else { 'no FLY certificates in CurrentUser\My or LocalMachine\My' }
  Add-Check 'Certificate (config thumbprint)' $false "Thumbprint '$thumb' not found for this account. Seen: $seen" 'Create the certificate on THIS machine/account (or import the .pfx), upload the .cer to the app registration, put the thumbprint in fly-config.json'
}

# Network reachability
foreach ($h in @('identity.avepointonlineservices.com', $(if ($cfg -and $cfg.flyUrl -match '^https://([^/]+)') { $Matches[1] }))) {
  if (-not $h) { continue }
  $ok = (Test-NetConnection $h -Port 443 -WarningAction SilentlyContinue).TcpTestSucceeded
  Add-Check "Network: ${h}:443" $ok $(if ($ok) { 'reachable' } else { 'blocked or unresolvable' }) 'Check proxy/firewall and the region in flyUrl'
}

# Optional live sign-in
if ($TryConnect -and $mod -and $match -and $cfg) {
  try {
    Import-Module Fly.Client -ErrorAction Stop
    Connect-Fly -Url $cfg.flyUrl -ClientId $cfg.clientId -Certificate $match -ErrorAction Stop
    $projects = & (Get-Module Fly.Client) { (Get-FlyProjects -Top 200 -Skip 0).data }
    Add-Check 'FLY sign-in' $true "Connected; $(@($projects).Count) project(s) visible" ''
    if ($projects) { $projects | Select-Object name, @{n = 'Type'; e = { $_.sourcePlatform } }, createTime | Sort-Object name | Format-Table -AutoSize | Out-Host }
  } catch { Add-Check 'FLY sign-in' $false $_.Exception.Message 'Check clientId, that this certificate''s .cer is uploaded to the app registration, and the flyUrl region' }
}

$results | Format-Table -AutoSize -Wrap
$missing = @($results | Where-Object Result -ne 'OK').Count
if ($missing) { Write-Host "$missing item(s) need fixing - see the Fix column." -ForegroundColor Yellow } else { Write-Host 'All checks passed.' -ForegroundColor Green }

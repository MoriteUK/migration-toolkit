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
  [string]$ConfigPath,
  [switch]$TryConnect
)
if (-not $ConfigPath) {
  # Same rule as Invoke-FlyBriefs.ps1: fly-config.json only wins when it holds real credentials
  $local = Join-Path $PSScriptRoot 'fly-config.json'
  $localOk = $false
  if (Test-Path $local) {
    try { $l = Get-Content $local -Raw | ConvertFrom-Json; $localOk = [bool]($l.EncSecret -or ($l.certificateThumbprint -and $l.certificateThumbprint -notmatch '^<')) } catch {}
  }
  $ConfigPath = if ($localOk) { $local } else { Join-Path $env:APPDATA 'FlyMigration\config.json' }
}
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

# Node.js (20+ for fs.openAsBlob) and a Chromium browser for the PDF
$node = Get-Command node -ErrorAction SilentlyContinue
$nodeVer = if ($node) { (& node --version) } else { $null }
Add-Check 'Node.js 20+' ($nodeVer -and [int]($nodeVer.TrimStart('v').Split('.')[0]) -ge 20) $(if ($nodeVer) { $nodeVer } else { 'not on PATH' }) 'winget install OpenJS.NodeJS.LTS  (then open a new window)'
$edge = @($env:EDGE_PATH, "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Google\Chrome\Application\chrome.exe") | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
Add-Check 'Edge/Chrome (PDF output)' ([bool]$edge) $(if ($edge) { $edge } else { 'not found' }) 'Install Microsoft Edge, or set $env:EDGE_PATH'

# Brief files
foreach ($f in 'FLY_Executive_Brief.html', 'fly-brief.js', 'Invoke-FlyBriefs.ps1') {
  $p = Join-Path $PSScriptRoot $f
  Add-Check "File: $f" (Test-Path $p) $p 'Copy the whole VGMigrations\FlyBrief folder to this machine'
}

# Config: fly-config.json (flyUrl/clientId/certificateThumbprint|EncSecret) or the toolkit config (Url/ClientId/EncSecret)
$cfg = $null
if (Test-Path $ConfigPath) {
  try { $cfg = Get-Content $ConfigPath -Raw | ConvertFrom-Json; Add-Check 'Config file' $true $ConfigPath '' }
  catch { Add-Check 'Config file' $false "Invalid JSON: $($_.Exception.Message)" 'Fix the JSON syntax (commas, quotes, \\ in paths)' }
} else { Add-Check 'Config file' $false "$ConfigPath not found" 'Configure the Fly API in the toolkit Settings, or copy fly-config.example.json to fly-config.json' }
$url = $clientId = $thumb = $secret = $null
if ($cfg) {
  $url      = if ($cfg.flyUrl) { $cfg.flyUrl } else { $cfg.Url }
  $clientId = if ($cfg.clientId) { $cfg.clientId } else { $cfg.ClientId }
  $thumb    = if ($cfg.certificateThumbprint -notmatch '^<') { $cfg.certificateThumbprint }
  $secret   = $cfg.EncSecret
  Add-Check 'Config: Fly URL' ($url -match '^https://graph(-[a-z0-9-]+)?\.avepointonlineservices\.com/fly/?$') $url 'Use https://graph[-<region>].avepointonlineservices.com/fly, e.g. graph-uk'
  Add-Check 'Config: client ID' ($clientId -match '^[0-9a-f-]{36}$') $clientId 'Paste the Application (client) ID from AOS > Administration > App registrations'
  if ($cfg.projects) {
    $ph = @($cfg.projects | Where-Object { $_.project -match '^<|>$' -or -not $_.client -or $_.workload -notin 'Exchange', 'OneDrive', 'SharePoint', 'Teams', 'M365Group', 'TeamChat' })
    Add-Check 'Config: projects' (-not $ph) "$(@($cfg.projects).Count) configured$(if ($ph) { ", $($ph.Count) still placeholders/invalid" })" 'Set each project to the exact FLY project name, or remove unused entries (discovery finds the rest)'
  }
}

# Credentials: DPAPI client secret, or certificate with private key
$match = $null
if ($secret) {
  $ok = $true; $detail = 'EncSecret decrypts for this account'
  try { $null = $secret | ConvertTo-SecureString } catch { $ok = $false; $detail = 'EncSecret will not decrypt for this Windows account' }
  Add-Check 'Credential (client secret)' $ok $detail 'Re-enter the client secret in toolkit Settings on THIS machine/account (DPAPI is per-user)'
} else {
  $certs = Get-ChildItem Cert:\CurrentUser\My, Cert:\LocalMachine\My -ErrorAction SilentlyContinue | Where-Object { ($thumb -and $_.Thumbprint -eq $thumb) -or $_.Subject -like '*FLY*' }
  $match = $certs | Where-Object { $thumb -and $_.Thumbprint -eq $thumb } | Select-Object -First 1
  if ($match) {
    Add-Check 'Certificate (config thumbprint)' $match.HasPrivateKey "$($match.PSParentPath -replace '.*::') expires $($match.NotAfter.ToString('d MMM yyyy'))$(if (-not $match.HasPrivateKey) { ' - NO private key' })" 'Import the .pfx (with private key), not the .cer'
    Add-Check 'Certificate not expired' ($match.NotAfter -gt (Get-Date)) $match.NotAfter.ToString('d MMM yyyy') 'Create a new certificate and upload its .cer to the app registration'
  } else {
    $seen = if ($certs) { ($certs | ForEach-Object { "$($_.Subject) $($_.Thumbprint)" }) -join '; ' } else { 'no FLY certificates in CurrentUser\My or LocalMachine\My' }
    Add-Check 'Credential' $false "No EncSecret, and thumbprint '$thumb' not found. Seen: $seen" 'Set the client secret in toolkit Settings, or install the certificate and put its thumbprint in fly-config.json'
  }
}

# Network reachability
foreach ($h in @('identity.avepointonlineservices.com', $(if ($url -match '^https://([^/]+)') { $Matches[1] }))) {
  if (-not $h) { continue }
  $ok = (Test-NetConnection $h -Port 443 -WarningAction SilentlyContinue).TcpTestSucceeded
  Add-Check "Network: ${h}:443" $ok $(if ($ok) { 'reachable' } else { 'blocked or unresolvable' }) 'Check proxy/firewall and the region in the Fly URL'
}

# Optional live sign-in
if ($TryConnect -and $mod -and $cfg -and ($secret -or $match)) {
  try {
    Import-Module Fly.Client -ErrorAction Stop
    if ($secret) {
      $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR(($secret | ConvertTo-SecureString))
      try { $plain = [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr) } finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
      Connect-Fly -Url $url -ClientId $clientId -ClientSecret $plain -ErrorAction Stop | Out-Null
    } else {
      Connect-Fly -Url $url -ClientId $clientId -Certificate $match -ErrorAction Stop | Out-Null
    }
    $projects = & (Get-Module Fly.Client) { (Get-FlyProjects -Top 200 -Skip 0).data }
    Add-Check 'FLY sign-in' $true "Connected; $(@($projects).Count) project(s) visible" ''
    if ($projects) { $projects | Select-Object name, @{n = 'Type'; e = { $_.sourcePlatform } }, createTime | Sort-Object name | Format-Table -AutoSize | Out-Host }
  } catch { Add-Check 'FLY sign-in' $false $_.Exception.Message 'Check the client ID, the credential, and the Fly URL region' }
}

$results | Format-Table -AutoSize -Wrap
$missing = @($results | Where-Object Result -ne 'OK').Count
if ($missing) { Write-Host "$missing item(s) need fixing - see the Fix column." -ForegroundColor Yellow } else { Write-Host 'All checks passed.' -ForegroundColor Green }

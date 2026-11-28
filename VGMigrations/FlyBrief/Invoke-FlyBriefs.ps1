<#
.SYNOPSIS
  Builds one-page executive briefs (HTML + PDF) from AvePoint FLY, one per migration project.
  Give a company name to get a brief for each of its workloads:
    .\Invoke-FlyBriefs.ps1 Caternet

.DESCRIPTION
  Uses AvePoint's Fly.Client PowerShell module. For every matching project it:
    1. runs the matching Export-Fly<Workload>MigrationReport (Exchange, OneDrive, SharePoint, Teams, M365 Groups, Teams Chat)
    2. expands any .zip it receives
    3. runs fly-brief.js (Node) to build <client>_<workload>_Executive_Brief.html/.pdf in that folder

  Config: uses fly-config.json next to this script when it exists (certificate or EncSecret auth,
  see fly-config.example.json). Otherwise it falls back to the Migration Toolkit's own config
  (%APPDATA%\FlyMigration\config.json: Url, ClientId, EncSecret) with auto-discovery of every project.

.PARAMETER Company
  Company name / customer prefix: the part of the FLY project name before " - ".
  "Caternet" matches "Caternet - Exchange", "Caternet - OneDrive", ...

.PARAMETER ClientName
  Name printed on the brief (e.g. "AEP Italia" for prefix "AEP"). Defaults to the config/discovered client.

.PARAMETER OutputRoot
  Overrides outputRoot from the config. Briefs go to <OutputRoot>\<client>\<project>\<yyyyMMdd-HHmm>\.

.PARAMETER InternalDomains
  Comma-separated internal email domains used to classify sharing-link recipients. Defaults to the
  config entry, then to the domains detected in the export.

.PARAMETER ResultPath
  Also write the per-project results as JSON to this file (used by the Migration Toolkit app).

.PARAMETER SkipExport
  Re-use the most recent export folder for each project and only rebuild the briefs.

.PARAMETER Project
  Only process projects whose FLY project name matches (wildcards allowed), e.g. 'AEP - OneDrive' or 'AEP*'.

.PARAMETER ListOnly
  Show which projects would be processed (after auto-discovery) and stop. Exports nothing.

.PARAMETER NoPdf
  Build the HTML brief only.

.EXAMPLE
  .\Invoke-FlyBriefs.ps1 Caternet
.EXAMPLE
  .\Invoke-FlyBriefs.ps1 AEP -ClientName "AEP Italia" -InternalDomains "aep-italia.it,easyticketing.it"
.EXAMPLE
  .\Invoke-FlyBriefs.ps1 Caternet -ListOnly
#>
[CmdletBinding()]
param(
  [Parameter(Position = 0)][string]$Company,
  [string]$ClientName,
  [string]$ConfigPath,
  [string]$OutputRoot,
  [string]$InternalDomains,
  [string]$ResultPath,
  [switch]$SkipExport,
  [string]$Project = '*',
  [switch]$ListOnly,
  [switch]$NoPdf
)
$ErrorActionPreference = 'Stop'

# ── Config: fly-config.json beside the script, else the toolkit's own config ──────────────────
if (-not $ConfigPath) {
  # fly-config.json only wins when it holds real credentials (a copied example with placeholders doesn't)
  $local = Join-Path $PSScriptRoot 'fly-config.json'
  $localOk = $false
  if (Test-Path $local) {
    try { $l = Get-Content $local -Raw | ConvertFrom-Json; $localOk = [bool]($l.EncSecret -or ($l.certificateThumbprint -and $l.certificateThumbprint -notmatch '^<')) } catch {}
  }
  $ConfigPath = if ($localOk) { $local } else { Join-Path $env:APPDATA 'FlyMigration\config.json' }
}
if (-not (Test-Path $ConfigPath)) { throw "Config not found: $ConfigPath. Configure the Fly API in the toolkit's Settings, or create fly-config.json." }
$raw = Get-Content $ConfigPath -Raw | ConvertFrom-Json
$isToolkitCfg = -not $raw.projects -and -not $raw.discover -and $raw.Url
$cfg = [pscustomobject]@{
  Url        = $(if ($raw.flyUrl) { $raw.flyUrl } else { $raw.Url })
  ClientId   = $(if ($raw.clientId) { $raw.clientId } else { $raw.ClientId })
  EncSecret  = $raw.EncSecret
  Thumbprint = $(if ($raw.certificateThumbprint -and $raw.certificateThumbprint -notmatch '^<') { $raw.certificateThumbprint })
  OutputRoot = $(if ($OutputRoot) { $OutputRoot } elseif ($raw.outputRoot) { $raw.outputRoot } else { Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'FlyBriefs' })
  Include    = $(if ($raw.include) { [string[]]$raw.include } else { @('ErrorObjects', 'WarningObjects', 'SuccessfulObjects', 'SkippedObjects', 'NotFoundObjects', 'UnsupportedObjects') })
  TzOffset   = $raw.timeZoneOffsetMinutes
  Discover   = $(if ($raw.discover) { $raw.discover } elseif ($isToolkitCfg -or -not $raw.projects) { [pscustomobject]@{ enabled = $true } })
  Projects   = @($raw.projects)
}
Write-Host "Config: $ConfigPath"
Write-Host "Output: $($cfg.OutputRoot)"

$briefJs = Join-Path $PSScriptRoot 'fly-brief.js'
if (-not (Get-Command node -ErrorAction SilentlyContinue)) { throw 'Node.js is required (https://nodejs.org).' }
$safeName = { param($s) ($s -replace '[^\w.-]+', '_').Trim('_') }

function Connect-FlyFromConfig {
  if (-not (Get-Module -ListAvailable Fly.Client)) { throw 'Fly.Client module not installed. Run: Install-Module -Name Fly.Client -Scope CurrentUser' }
  if (-not (Get-Module Fly.Client)) { Import-Module Fly.Client }
  if (-not $cfg.Url -or -not $cfg.ClientId) { throw "Config $ConfigPath has no Fly URL / ClientId." }

  if ($cfg.EncSecret) {
    # DPAPI-encrypted client secret (same as the rest of the toolkit) - only decrypts for the Windows account that encrypted it
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR(($cfg.EncSecret | ConvertTo-SecureString))
    try { $plain = [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr) } finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
    Connect-Fly -Url $cfg.Url -ClientId $cfg.ClientId -ClientSecret $plain | Out-Null
    Write-Host "Connected to FLY at $($cfg.Url) (client secret)"
  }
  elseif ($cfg.Thumbprint) {
    $cert = Get-ChildItem Cert:\CurrentUser\My, Cert:\LocalMachine\My -ErrorAction SilentlyContinue |
            Where-Object Thumbprint -eq $cfg.Thumbprint | Select-Object -First 1
    if (-not $cert) { throw "Certificate $($cfg.Thumbprint) not found in CurrentUser\My or LocalMachine\My." }
    Connect-Fly -Url $cfg.Url -ClientId $cfg.ClientId -Certificate $cert | Out-Null
    Write-Host "Connected to FLY at $($cfg.Url) (certificate)"
  }
  else {
    throw "No credentials in $ConfigPath. It needs 'EncSecret' (client secret) or 'certificateThumbprint'."
  }
}

function Invoke-WithRetry([scriptblock]$Action, [string]$What) {
  for ($i = 1; $i -le 4; $i++) {
    try { return & $Action }
    catch {
      $msg = $_.Exception.Message
      if ($msg -match '401|Unauthori') { Write-Warning "${What}: token expired, reconnecting"; Connect-FlyFromConfig }
      elseif ($msg -match '429|Too Many') { Write-Warning "${What}: throttled, waiting 60s"; Start-Sleep -Seconds 60 }
      elseif ($i -ge 2) { throw }
      else { Write-Warning "$What failed ($msg), retrying"; Start-Sleep -Seconds 15 }
    }
  }
  throw "$What failed after retries."
}

# All FLY projects via the module's internal (unexported) Get-FlyProjects, paging 200 at a time.
function Get-AllFlyProjects {
  $mod = Get-Module Fly.Client
  $all = @(); $skip = 0
  do {
    $page = & $mod { param($s) Get-FlyProjects -Top 200 -Skip $s } $skip
    $batch = @($page.data); $all += $batch; $skip += $batch.Count
  } while ($page.nextLink -and $batch.Count)
  $all
}

# Build the work list: explicit "projects" entries, plus auto-discovered ones when discovery is on.
function Get-WorkList {
  $typeMap = @{ 0 = 'Exchange'; 1 = 'Teams'; 2 = 'SharePoint'; 3 = 'OneDrive'; 4 = 'M365Group'; 5 = 'TeamChat' }   # FLY sourcePlatform codes
  $list = [ordered]@{}
  if ($cfg.Discover -and $cfg.Discover.enabled) {
    $want = if ($cfg.Discover.workloads) { @($cfg.Discover.workloads) } else { @($typeMap.Values) }
    foreach ($fp in Get-AllFlyProjects) {
      $wl = $typeMap[[int]$fp.sourcePlatform]
      if (-not $wl -or $wl -notin $want) { continue }
      if (@($cfg.Discover.exclude | Where-Object { $_ -and $fp.name -like $_ }).Count) { continue }
      $client = ($fp.name -split '\s+-\s+', 2)[0].Trim()
      $list[$fp.name] = [pscustomobject]@{ project = $fp.name; client = $client; workload = $wl; internalDomains = @(); source = 'discovered' }
    }
  }
  foreach ($p in $cfg.Projects) {
    if (-not $p.project -or $p.project -match '^<') { continue }
    $list[$p.project] = [pscustomobject]@{ project = $p.project; client = $p.client; workload = $p.workload; internalDomains = @($p.internalDomains); source = 'config' }
  }
  @($list.Values | Where-Object {
      $_.project -like $Project -and
      (-not $Company -or $_.client -eq $Company -or $_.project -like "$Company - *" -or $_.project -eq $Company)
    })
}

function Save-Results($rows) {
  if ($ResultPath) { ConvertTo-Json -InputObject @($rows) -Depth 4 | Set-Content -Path $ResultPath -Encoding UTF8 }
}

$needFly = -not $SkipExport -or ($cfg.Discover -and $cfg.Discover.enabled)
if ($needFly) { Connect-FlyFromConfig }
$work = Get-WorkList
if (-not $work) {
  Save-Results @()
  Write-Warning "No FLY projects matched company '$Company' / project '$Project'. Run with -ListOnly and no name to see every project."
  return
}

if ($ListOnly) {
  $work | Sort-Object project | Format-Table project, client, workload, source -AutoSize | Out-Host
  Write-Host "$($work.Count) project(s) would be processed. Nothing was exported." -ForegroundColor Cyan
  return
}

Write-Host "$($work.Count) project(s) to process: $(($work.project | Sort-Object) -join ', ')"
$results = @()
foreach ($p in $work) {
  $client = if ($ClientName) { $ClientName } else { $p.client }
  $root = Join-Path (Join-Path $cfg.OutputRoot (& $safeName $client)) (& $safeName $p.project)
  $out = $null
  $row = [ordered]@{ project = $p.project; client = $client; workload = $p.workload; status = ''; folder = $null; htmlPath = $null; pdfPath = $null; summaryPath = $null }
  try {
    if ($SkipExport) {
      $out = Get-ChildItem $root -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending | Select-Object -First 1 -ExpandProperty FullName
      if (-not $out) { throw "No previous export found under $root" }
    } else {
      $out = Join-Path $root (Get-Date -Format 'yyyyMMdd-HHmm')
      New-Item -ItemType Directory -Force -Path $out | Out-Null
      $cmd = "Export-Fly$($p.workload)MigrationReport"
      if (-not (Get-Command $cmd -ErrorAction SilentlyContinue)) { throw "Unsupported workload '$($p.workload)' (no $cmd in Fly.Client)" }
      $exportArgs = @{ Project = $p.project; OutFolder = $out; FileType = 'CSV'; Include = $cfg.Include }
      if ($cfg.TzOffset) { $exportArgs.TimeZoneOffset = [int]$cfg.TzOffset }
      Write-Host "[$($p.project)] exporting $($p.workload) report..."
      Invoke-WithRetry { & $cmd @exportArgs } "$($p.project) report export" | Out-Null
      Get-ChildItem $out -Filter *.zip -Recurse | ForEach-Object {
        Expand-Archive $_.FullName -DestinationPath (Join-Path $_.DirectoryName $_.BaseName) -Force
      }
    }
    $row.folder = $out
    if (-not (Get-ChildItem $out -Recurse -Filter '*Mappings.csv')) {
      throw "No *Mappings.csv in $out. Either the project has no migrated mappings yet, or FLY returned no report file (Fly.Client issue #16)."
    }
    Write-Host "[$($p.project)] building brief..."
    $nodeArgs = @($briefJs, '--in', $out, '--client', $client, '--workload', $p.workload, '--out', $out)
    $internal = if ($InternalDomains) { $InternalDomains } elseif ($p.internalDomains -and $p.internalDomains.Count) { $p.internalDomains -join ',' }
    if ($internal) { $nodeArgs += @('--internal', $internal) }
    if ($NoPdf) { $nodeArgs += '--no-pdf' }
    & node @nodeArgs
    if ($LASTEXITCODE) { throw "fly-brief.js exited with $LASTEXITCODE" }
    $row.htmlPath    = Get-ChildItem $out -Filter '*_Executive_Brief.html' | Select-Object -First 1 -ExpandProperty FullName
    $row.pdfPath     = Get-ChildItem $out -Filter '*_Executive_Brief.pdf' | Select-Object -First 1 -ExpandProperty FullName
    $row.summaryPath = Get-ChildItem $out -Filter '*_Executive_Brief.summary.json' | Select-Object -First 1 -ExpandProperty FullName
    $row.status = 'OK'
    Write-Host "[$($p.project)] done" -ForegroundColor Green
  } catch {
    Write-Warning "[$($p.project)] $($_.Exception.Message)"
    $row.status = 'Failed: ' + $_.Exception.Message
  }
  $results += [pscustomobject]$row
}

if ($needFly) { try { Disconnect-Fly | Out-Null } catch {} }
Save-Results $results
$results | Format-Table project, status, folder -AutoSize -Wrap | Out-Host

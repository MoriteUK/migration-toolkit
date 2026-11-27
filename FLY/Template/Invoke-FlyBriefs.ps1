<#
.SYNOPSIS
  Builds one-page executive briefs (HTML + PDF) from AvePoint FLY, one per migration project.
  Give a company name to get a brief for each of its workloads:
    .\Invoke-FlyBriefs.ps1 Caternet

.DESCRIPTION
  Uses AvePoint's Fly.Client PowerShell module (Install-Module Fly.Client) with an AvePoint Online
  Services app registration and certificate. For every project in the config file it:
    1. runs the matching Export-Fly<Workload>MigrationReport (Exchange, OneDrive, SharePoint, Teams, M365 Groups, Teams Chat)
    2. expands any .zip it receives
    3. runs fly-brief.js (Node) to build <client>_Executive_Brief.html/.pdf in that folder

.PARAMETER Company
  Company name (the part of the FLY project name before " - "). Processes every project for that company,
  e.g. "Caternet" matches "Caternet - Exchange", "Caternet - OneDrive", "Caternet - Sharepoint", ...

.PARAMETER ConfigPath
  JSON config (see fly-config.example.json).

.PARAMETER SkipExport
  Re-use the most recent export folder for each project and only rebuild the briefs.

.PARAMETER Project
  Only process projects whose FLY project name matches (wildcards allowed), e.g. 'AEP - OneDrive' or 'AEP*'.

.PARAMETER ListOnly
  Show which projects would be processed (after auto-discovery) and stop. Exports nothing.

.EXAMPLE
  .\Invoke-FlyBriefs.ps1 Caternet
.EXAMPLE
  .\Invoke-FlyBriefs.ps1 Caternet -ListOnly
.EXAMPLE
  .\Invoke-FlyBriefs.ps1            # every discovered project
#>
[CmdletBinding()]
param(
  [Parameter(Position = 0)][string]$Company,
  [string]$ConfigPath = (Join-Path $PSScriptRoot 'fly-config.json'),
  [switch]$SkipExport,
  [string]$Project = '*',
  [switch]$ListOnly
)
$ErrorActionPreference = 'Stop'

$cfg = Get-Content $ConfigPath -Raw | ConvertFrom-Json
$briefJs = Join-Path $PSScriptRoot 'fly-brief.js'
if (-not (Get-Command node -ErrorAction SilentlyContinue)) { throw 'Node.js is required (https://nodejs.org).' }
$safeName = { param($s) ($s -replace '[^\w.-]+', '_').Trim('_') }

function Connect-FlyFromConfig {
  if (-not (Get-Module -ListAvailable Fly.Client)) { throw 'Fly.Client module not installed. Run: Install-Module -Name Fly.Client -Scope CurrentUser' }
  if (-not (Get-Module Fly.Client)) { Import-Module Fly.Client }
  $cert = Get-ChildItem Cert:\CurrentUser\My, Cert:\LocalMachine\My -ErrorAction SilentlyContinue |
          Where-Object Thumbprint -eq $cfg.certificateThumbprint | Select-Object -First 1
  if (-not $cert) { throw "Certificate $($cfg.certificateThumbprint) not found in CurrentUser\My or LocalMachine\My." }
  Connect-Fly -Url $cfg.flyUrl -ClientId $cfg.clientId -Certificate $cert | Out-Null
  Write-Host "Connected to FLY at $($cfg.flyUrl)"
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

# Build the work list: explicit "projects" entries, plus auto-discovered ones when "discover.enabled" is true.
function Get-WorkList {
  $typeMap = @{ 0 = 'Exchange'; 1 = 'Teams'; 2 = 'SharePoint'; 3 = 'OneDrive'; 4 = 'M365Group'; 5 = 'TeamChat' }   # FLY sourcePlatform codes
  $list = [ordered]@{}
  if ($cfg.discover -and $cfg.discover.enabled) {
    $want = if ($cfg.discover.workloads) { @($cfg.discover.workloads) } else { @($typeMap.Values) }
    foreach ($fp in Get-AllFlyProjects) {
      $wl = $typeMap[[int]$fp.sourcePlatform]
      if (-not $wl -or $wl -notin $want) { continue }
      if (@($cfg.discover.exclude | Where-Object { $_ -and $fp.name -like $_ }).Count) { continue }
      $client = ($fp.name -split '\s+-\s+', 2)[0].Trim()
      $list[$fp.name] = [pscustomobject]@{ project = $fp.name; client = $client; workload = $wl; internalDomains = @(); source = 'discovered' }
    }
  }
  foreach ($p in $cfg.projects) {
    if (-not $p.project -or $p.project -match '^<') { continue }
    $list[$p.project] = [pscustomobject]@{ project = $p.project; client = $p.client; workload = $p.workload; internalDomains = @($p.internalDomains); source = 'config' }
  }
  @($list.Values | Where-Object { $_.project -like $Project -and (-not $Company -or $_.client -eq $Company -or $_.project -like "$Company*") })
}

$needFly = -not $SkipExport -or ($cfg.discover -and $cfg.discover.enabled)
if ($needFly) { Connect-FlyFromConfig }
$work = Get-WorkList
if (-not $work) { Write-Warning "No projects matched company '$Company' / project '$Project'. Run with -ListOnly and no name to see every project."; return }

if ($ListOnly) {
  $work | Sort-Object project | Format-Table project, client, workload, source -AutoSize
  Write-Host "$($work.Count) project(s) would be processed. Nothing was exported." -ForegroundColor Cyan
  return
}

$results = @()
foreach ($p in $work) {
  $root = Join-Path (Join-Path $cfg.outputRoot (& $safeName $p.client)) (& $safeName $p.project)
  $out = $null
  try {
    if ($SkipExport) {
      $out = Get-ChildItem $root -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending | Select-Object -First 1 -ExpandProperty FullName
      if (-not $out) { throw "No previous export found under $root" }
    } else {
      $out = Join-Path $root (Get-Date -Format 'yyyyMMdd-HHmm')
      New-Item -ItemType Directory -Force -Path $out | Out-Null
      $cmd = "Export-Fly$($p.workload)MigrationReport"
      if (-not (Get-Command $cmd -ErrorAction SilentlyContinue)) { throw "Unsupported workload '$($p.workload)' (no $cmd in Fly.Client)" }
      $exportArgs = @{ Project = $p.project; OutFolder = $out; FileType = 'CSV' }
      if ($cfg.include) { $exportArgs.Include = [string[]]$cfg.include }
      if ($cfg.timeZoneOffsetMinutes) { $exportArgs.TimeZoneOffset = [int]$cfg.timeZoneOffsetMinutes }
      Write-Host "[$($p.project)] exporting $($p.workload) report..."
      Invoke-WithRetry { & $cmd @exportArgs } "$($p.project) report export" | Out-Null
      Get-ChildItem $out -Filter *.zip -Recurse | ForEach-Object {
        Expand-Archive $_.FullName -DestinationPath (Join-Path $_.DirectoryName $_.BaseName) -Force
      }
    }
    if (-not (Get-ChildItem $out -Recurse -Filter '*Mappings.csv')) {
      throw "No *Mappings.csv in $out. Either the project has no migrated mappings yet, or FLY returned no report file (Fly.Client issue #16)."
    }
    $nodeArgs = @($briefJs, '--in', $out, '--client', $p.client, '--workload', $p.workload, '--out', $out)
    if ($p.internalDomains -and $p.internalDomains.Count) { $nodeArgs += @('--internal', ($p.internalDomains -join ',')) }
    & node @nodeArgs
    if ($LASTEXITCODE) { throw "fly-brief.js exited with $LASTEXITCODE" }
    $results += [pscustomobject]@{ Project = $p.project; Status = 'OK'; Folder = $out }
  } catch {
    Write-Warning "[$($p.project)] $($_.Exception.Message)"
    $results += [pscustomobject]@{ Project = $p.project; Status = 'Failed: ' + $_.Exception.Message; Folder = $out }
  }
}

if ($needFly) { try { Disconnect-Fly | Out-Null } catch {} }
$results | Format-Table -AutoSize -Wrap
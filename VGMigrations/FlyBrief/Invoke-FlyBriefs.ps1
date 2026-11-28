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

.PARAMETER ExportFolder
  Build briefs from report exports already downloaded from the FLY portal (Download Center) instead
  of exporting through the API. Any .zip files are expanded; every folder holding a *Mappings.csv
  gets its own brief, written next to the CSVs. No FLY sign-in is needed.
  Use this while the FLY API returns no report URL for export jobs (Fly.Client issue #16).

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
  [string]$ExportFolder,
  [string]$ResultPath,
  [switch]$SkipExport,
  [string]$Project = '*',
  [switch]$ListOnly,
  [switch]$NoPdf,
  [switch]$NoPortal,      # skip the FLY-portal report download; use live API data only
  [switch]$TechReport,    # build the interactive TechTeam drill-down HTML instead of the briefs
  [switch]$UseApiReport   # also try Fly.Client's report export (broken while Fly.Client issue #16 is open)
)
$ErrorActionPreference = 'Stop'
# Plain text output: the toolkit's log panel shows ANSI colour codes from Format-Table literally
if ($PSStyle) { $PSStyle.OutputRendering = 'PlainText' }

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
if (Test-Path $ConfigPath) { $raw = Get-Content $ConfigPath -Raw | ConvertFrom-Json }
elseif ($ExportFolder) { $raw = [pscustomobject]@{} }   # export-folder mode never talks to FLY
else { throw "Config not found: $ConfigPath. Configure the Fly API in the toolkit's Settings, or create fly-config.json." }
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
if (-not $ExportFolder) { Write-Host "Output: $($cfg.OutputRoot)" }

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
      $list[$fp.name] = [pscustomobject]@{ project = $fp.name; id = $fp.id; client = $client; workload = $wl; internalDomains = @(); source = 'discovered' }
    }
  }
  $ids = $null
  foreach ($p in $cfg.Projects) {
    if (-not $p.project -or $p.project -match '^<') { continue }
    $id = if ($list.Contains($p.project)) { $list[$p.project].id } else {
      if ($null -eq $ids -and (Get-Module Fly.Client)) { $ids = @{}; Get-AllFlyProjects | ForEach-Object { $ids[$_.name] = $_.id } }
      if ($ids) { $ids[$p.project] }
    }
    $list[$p.project] = [pscustomobject]@{ project = $p.project; id = $id; client = $p.client; workload = $p.workload; internalDomains = @($p.internalDomains); source = 'config' }
  }
  @($list.Values | Where-Object {
      $_.project -like $Project -and
      (-not $Company -or $_.client -eq $Company -or $_.project -like "$Company - *" -or $_.project -eq $Company)
    })
}

# Brief summary built from FLY's live project + mapping data (no report files). Matches the report's
# mapping-level figures (status counts, migrated GB, job dates); item-level, permission and sharing-link
# figures are not available from the API, so the brief renders in its "api" mode.
# Every mapping of a FLY project with its live status (mappings/summaries, paged)
function Get-FlyLiveMappings([string]$ProjectId) {
  $conf = & (Get-Module Fly.Client) { Get-FlyConfiguration }
  $base = ([string]$conf['BaseUrl']).TrimEnd('/')
  $maps = [System.Collections.Generic.List[object]]::new()
  $skip = 0
  do {
    $headers = @{ Authorization = "Bearer $((& (Get-Module Fly.Client) { Get-FlyConfiguration })['AccessToken'])"; Accept = 'application/json' }
    $r = Invoke-WithRetry { Invoke-RestMethod "$base/projects/$ProjectId/mappings/summaries?top=200&skip=$skip" -Headers $headers } "mapping summaries"
    $page = @($r.data.data)
    foreach ($x in $page) { $maps.Add($x) }
    $skip += $page.Count
  } while ($r.data.nextLink -and $page.Count)
  , $maps
}

# Downloaded FLY reports under a folder: one entry per distinct report (workload + export second),
# preferring the copy nearest the top folder. Zips are expanded first.
function Get-ExportReports([string]$Folder, [string]$ExcludeUnder) {
  Get-ChildItem $Folder -Filter *.zip -Recurse | ForEach-Object {
    $dest = Join-Path $_.DirectoryName $_.BaseName
    if (-not (Test-Path $dest)) { Write-Host "Expanding $($_.Name)"; Expand-Archive $_.FullName -DestinationPath $dest -Force }
  }
  $reports = [ordered]@{}
  Get-ChildItem $Folder -Recurse -File -Filter '*Mappings.csv' |
    Where-Object { -not $ExcludeUnder -or $_.FullName -notlike "$ExcludeUnder*" } |
    Sort-Object { ($_.FullName -split '\\').Count }, FullName | ForEach-Object {
      if ($_.Name -match '^(\w+?)(?:Summary|Migration)Report_(\d{14})') { $key = "$($Matches[1])|$($Matches[2])"; $wl = $Matches[1] } else { $key = $_.FullName; $wl = '' }
      if (-not $reports.Contains($key)) { $reports[$key] = [pscustomobject]@{ workload = $wl; path = $_.FullName } }
      else { Write-Host "Skipping duplicate copy: $($_.FullName)" }
    }
  , @($reports.Values)
}

function Get-FlyLiveSummary([string]$ProjectId, [string]$Workload) {
  $maps = Get-FlyLiveMappings $ProjectId

  # lastMigrationStatus: 0 not run, 4 finished, 5 finished with exceptions, 6 failed (others = running/stopped/queued)
  $ran = @($maps | Where-Object { [int]$_.lastMigrationStatus -ne 0 })
  $toDay = { param($t) if ([long]$t -gt 0) { ([datetime]::new([long]$t)).ToString('yyyy-MM-dd') } }
  $starts = @($ran | ForEach-Object { & $toDay $_.lastMigrationStartTime } | Where-Object { $_ } | Sort-Object)
  $failed = @($ran | Where-Object { [int]$_.lastMigrationStatus -eq 6 })
  $domains = @($maps | ForEach-Object { $_.destinationIdentity, $_.sourceIdentity } | Where-Object { $_ -match '@([^@]+)$' } |
               ForEach-Object { ($_ -split '@')[-1].ToLower() } | Where-Object { $_ -notmatch 'onmicrosoft\.com$' } | Sort-Object -Unique)
  [ordered]@{
    workload   = $Workload
    source     = 'api'
    exportDate = (Get-Date).ToString('yyyy-MM-dd')
    m          = [ordered]@{
      total          = $ran.Count
      notStarted     = $maps.Count - $ran.Count
      finished       = @($ran | Where-Object { [int]$_.lastMigrationStatus -eq 4 }).Count
      exceptions     = @($ran | Where-Object { [int]$_.lastMigrationStatus -eq 5 }).Count
      failed         = $failed.Count
      gb             = [math]::Round((($ran | Measure-Object migratedSize -Sum).Sum) / 1GB, 2)
      ok = 0; exc = 0; err = 0; skip = 0; topErr = 0
      failedWithData = @($failed | Where-Object { [long]$_.migratedSize -gt 0 }).Count
      firstStart     = $(if ($starts) { $starts[0] } else { '' })
      lastEnd        = $(if ($starts) { $starts[-1] } else { '' })
    }
    sec = $null; links = $null
    domains = $domains
    files = [ordered]@{ mappings = 'FLY live API'; security = $null; links = $null; skipped = 0 }
  }
}

# Full migration report fetched through the FLY portal with the saved AOS sign-in (fly-portal-report.js).
# Returns $true when report CSVs landed in $OutFolder; $false when the portal route is unavailable.
function Get-FlyReportViaPortal([string]$ProjectName, [string]$Workload, [string]$OutFolder) {
  $portalJs = Join-Path $PSScriptRoot 'fly-portal-report.js'
  if (-not (Test-Path $portalJs)) { return $false }
  Write-Host "[$ProjectName] requesting the full migration report through the FLY portal..."
  & node $portalJs --project $ProjectName --workload $Workload --out $OutFolder | ForEach-Object { Write-Host "  $_" }
  if ($LASTEXITCODE -eq 3) { Write-Warning 'FLY portal sign-in has expired - sign in again (Settings > AOS session) to get permission and sharing-link figures.'; return $false }
  if ($LASTEXITCODE) { throw "fly-portal-report.js exited with $LASTEXITCODE" }
  Get-ChildItem $OutFolder -Filter *.zip -Recurse | ForEach-Object { Expand-Archive $_.FullName -DestinationPath (Join-Path $_.DirectoryName $_.BaseName) -Force }
  return [bool](Get-ChildItem $OutFolder -Recurse -Filter '*Mappings.csv')
}

function Save-Results($rows) {
  if ($ResultPath) { ConvertTo-Json -InputObject @($rows) -Depth 4 | Set-Content -Path $ResultPath -Encoding UTF8 }
}

# Run fly-brief.js for one area and fill in the result row from the paths it reports.
function Invoke-AreaBrief($row, [string[]]$SourceArgs, [string]$OutDir) {
  $nodeArgs = @($briefJs) + $SourceArgs + @('--client', $row.client, '--out', $OutDir)
  if ($row.workload) { $nodeArgs += @('--workload', $row.workload) }
  if ($InternalDomains) { $nodeArgs += @('--internal', $InternalDomains) }
  if ($NoPdf) { $nodeArgs += '--no-pdf' }
  $lines = @(& node @nodeArgs 2>&1 | ForEach-Object { "$_" })
  $lines | ForEach-Object { Write-Host "  $_" }
  if ($LASTEXITCODE) { throw "fly-brief.js exited with $LASTEXITCODE" }
  $row.htmlPath = ($lines | Where-Object { $_ -match '^HTML:\s+(.+)$' } | ForEach-Object { $Matches[1].Trim() } | Select-Object -First 1)
  $row.pdfPath  = ($lines | Where-Object { $_ -match '^PDF:\s+(.+)$' }  | ForEach-Object { $Matches[1].Trim() } | Select-Object -First 1)
  if ($row.htmlPath) { $row.summaryPath = $row.htmlPath -replace '\.html$', '.summary.json' }
  $row.folder = $OutDir
  $row.status = 'OK'
}

# One document for the executive team: summary page + one page per area (static HTML + PDF).
# Adds a row with workload 'All' so the app can offer it first.
function New-FlyBriefPack($rows, [string]$Client, [string]$OutDir) {
  $sums = @($rows | Where-Object { $_.status -eq 'OK' -and $_.summaryPath -and (Test-Path $_.summaryPath) } | ForEach-Object { $_.summaryPath })
  if (-not $sums.Count) { return $null }
  Write-Host "[$Client] building the all-areas executive brief ($($sums.Count) areas)..."
  $nodeArgs = @($briefJs, '--pack', ($sums -join ';'), '--client', $Client, '--out', $OutDir)
  if ($InternalDomains) { $nodeArgs += @('--internal', $InternalDomains) }
  if ($NoPdf) { $nodeArgs += '--no-pdf' }
  $lines = @(& node @nodeArgs 2>&1 | ForEach-Object { "$_" })
  $lines | ForEach-Object { Write-Host "  $_" }
  $row = [ordered]@{ project = "$Client - all areas"; client = $Client; workload = 'All'; status = ''; source = 'combined'; folder = $OutDir; htmlPath = $null; pdfPath = $null; summaryPath = $null }
  if ($LASTEXITCODE) { $row.status = "Failed: fly-brief.js --pack exited with $LASTEXITCODE"; return [pscustomobject]$row }
  $row.htmlPath = ($lines | Where-Object { $_ -match '^PACK HTML:\s+(.+)$' } | ForEach-Object { $Matches[1].Trim() } | Select-Object -First 1)
  $row.pdfPath  = ($lines | Where-Object { $_ -match '^PACK PDF:\s+(.+)$' }  | ForEach-Object { $Matches[1].Trim() } | Select-Object -First 1)
  $row.status = 'OK'
  [pscustomobject]$row
}

$areaOrder = 'Exchange', 'OneDrive', 'SharePoint', 'Teams', 'M365Group', 'TeamChat'
$platformArea = @{ 0 = 'Exchange'; 1 = 'Teams'; 2 = 'SharePoint'; 3 = 'OneDrive'; 4 = 'M365Group'; 5 = 'TeamChat' }   # FLY sourcePlatform codes

# ── TechTeam report: interactive drill-down HTML for the tech team (HTML only, no PDF) ────────────
# Areas come from downloaded FLY reports in -ExportFolder (item-level detail); every other area of the
# customer's FLY projects (-Company) comes from live mapping status. Output: <Client>_TechTeam_Report.html
if ($TechReport) {
  if (-not $ExportFolder -and -not $Company) { throw 'The TechTeam report needs -Company (live FLY status) and/or -ExportFolder (downloaded FLY reports).' }
  $client = if ($ClientName) { $ClientName } elseif ($Company) { $Company } else { 'Migration' }
  $outDir = if ($OutputRoot) { $OutputRoot }
            elseif ($ExportFolder) { Join-Path (Resolve-Path $ExportFolder).Path 'TechTeam Report' }
            else { Join-Path (Join-Path $cfg.OutputRoot (& $safeName $client)) 'TechTeam Report' }
  New-Item -ItemType Directory -Force -Path $outDir | Out-Null
  $areas = [System.Collections.Generic.List[object]]::new()
  if ($ExportFolder) {
    if (-not (Test-Path $ExportFolder -PathType Container)) { throw "Export folder not found: $ExportFolder" }
    foreach ($r in (Get-ExportReports (Resolve-Path $ExportFolder).Path $outDir)) {
      if ($r.workload) { $areas.Add([pscustomobject]@{ workload = $r.workload; report = $r.path }); Write-Host "[$($r.workload)] downloaded report: $(Split-Path $r.path -Leaf)" }
    }
  }
  if ($Company) {
    try {
      Connect-FlyFromConfig
      $covered = @($areas | ForEach-Object workload)
      $projects = @(Get-AllFlyProjects | Where-Object {
          $area = $platformArea[[int]$_.sourcePlatform]
          $area -and $area -notin $covered -and ($_.name -like "$Company - *" -or ($_.name -split '\s+-\s+', 2)[0].Trim() -eq $Company)
        } | Sort-Object name)
      foreach ($fp in $projects) {
        $wl = $platformArea[[int]$fp.sourcePlatform]
        Write-Host "[$($fp.name)] reading live mapping status from FLY..."
        $maps = Get-FlyLiveMappings $fp.id   # assign first: piping the function directly passes the whole list as one object
        $rows = @($maps | ForEach-Object {
            [pscustomobject]@{ src = $_.sourceIdentity; dst = $_.destinationIdentity; status = [int]$_.lastMigrationStatus
              gb = [math]::Round([long]$_.migratedSize / 1GB, 2); errorItems = [int]$_.errorItemCount
              start = $(if ([long]$_.lastMigrationStartTime -gt 0) { ([datetime]::new([long]$_.lastMigrationStartTime)).ToString('yyyy-MM-dd HH:mm') } else { '' }) }
          })
        $areas.Add([pscustomobject]@{ workload = $wl; live = [pscustomobject]@{ project = $fp.name; mappings = $rows } })
      }
    } catch { Write-Warning "Could not read live FLY status: $($_.Exception.Message)" }
    finally { try { Disconnect-Fly | Out-Null } catch {} }
  }
  if (-not $areas.Count) { Save-Results @(); throw "Nothing to report: no downloaded FLY reports$(if ($ExportFolder) { " under $ExportFolder" }) and no FLY projects for '$Company'." }

  $manifest = Join-Path ([IO.Path]::GetTempPath()) ("fly-tech-" + [guid]::NewGuid().ToString('N') + '.json')
  [pscustomobject]@{ client = $client; areas = $areas } | ConvertTo-Json -Depth 8 -Compress | Set-Content $manifest -Encoding UTF8
  $row = [ordered]@{ project = "$client - TechTeam report"; client = $client; workload = 'TechTeam'; status = ''; source = "$($areas.Count) areas"; folder = $outDir; htmlPath = $null; pdfPath = $null; summaryPath = $null }
  try {
    $lines = @(& node (Join-Path $PSScriptRoot 'fly-issues.js') --manifest $manifest --out $outDir 2>&1 | ForEach-Object { "$_" })
    $lines | Where-Object { $_ -notmatch '^TECH HTML:' } | ForEach-Object { Write-Host $_ }
    if ($LASTEXITCODE) { throw "fly-issues.js exited with $LASTEXITCODE" }
    $row.htmlPath = ($lines | Where-Object { $_ -match '^TECH HTML:\s+(.+)$' } | ForEach-Object { $Matches[1].Trim() } | Select-Object -First 1)
    $row.status = 'OK'
    Write-Host "TechTeam report: $($row.htmlPath)" -ForegroundColor Green
  } catch {
    Write-Warning $_.Exception.Message
    $row.status = 'Failed: ' + $_.Exception.Message
  } finally { Remove-Item $manifest -ErrorAction SilentlyContinue }
  Save-Results @([pscustomobject]$row)
  return
}

# ── Export-folder mode: briefs from reports downloaded from the FLY portal ───────────────────────
# Every *Mappings.csv is one report; its detail files are matched by workload + export time, so a folder
# holding several areas (or copies of the same extract) never mixes their data. Duplicate copies of the
# same report are built once. Areas with no downloaded report are filled from live FLY data when -Company
# names the customer. Output: one brief per area plus the all-areas pack, in one folder.
if ($ExportFolder) {
  if (-not (Test-Path $ExportFolder -PathType Container)) { throw "Export folder not found: $ExportFolder" }
  $ExportFolder = (Resolve-Path $ExportFolder).Path
  $client = if ($ClientName) { $ClientName } elseif ($Company) { $Company } else { 'Migration' }
  $outDir = if ($OutputRoot) { $OutputRoot } else { Join-Path $ExportFolder 'Executive Brief' }
  New-Item -ItemType Directory -Force -Path $outDir | Out-Null
  $reportList = Get-ExportReports $ExportFolder $outDir
  if (-not $reportList.Count -and -not $Company) { Save-Results @(); throw "No *Mappings.csv found under $ExportFolder. Download each project's migration report (CSV) from the FLY portal's Download Center and unzip it there." }

  $results = @()
  foreach ($r in ($reportList | Sort-Object { $i = $areaOrder.IndexOf($_.workload); if ($i -lt 0) { 99 } else { $i } })) {
    $row = [ordered]@{ project = "$client - $(if ($r.workload) { $r.workload } else { 'export' })"; client = $client; workload = $r.workload; status = ''; source = 'FLY report (downloaded)'; folder = $null; htmlPath = $null; pdfPath = $null; summaryPath = $null }
    try {
      Write-Host "[$($row.project)] building brief from $(Split-Path $r.path -Leaf)"
      Invoke-AreaBrief $row @('--report', $r.path) $outDir
    } catch {
      Write-Warning "[$($row.project)] $($_.Exception.Message)"
      $row.status = 'Failed: ' + $_.Exception.Message
    }
    $results += [pscustomobject]$row
  }

  # Areas with no downloaded report: live FLY data for the customer's matching project(s)
  $covered = @($results | Where-Object status -eq 'OK' | ForEach-Object workload)
  $missing = @($areaOrder | Where-Object { $_ -notin $covered })
  if ($Company -and $missing.Count) {
    try {
      Connect-FlyFromConfig
      $projects = @(Get-AllFlyProjects | Where-Object {
          $area = $platformArea[[int]$_.sourcePlatform]
          $area -in $missing -and ($_.name -like "$Company - *" -or ($_.name -split '\s+-\s+', 2)[0].Trim() -eq $Company)
        })
      foreach ($fp in $projects) {
        $wl = $platformArea[[int]$fp.sourcePlatform]
        $row = [ordered]@{ project = $fp.name; client = $client; workload = $wl; status = ''; source = 'FLY live data'; folder = $null; htmlPath = $null; pdfPath = $null; summaryPath = $null }
        try {
          Write-Host "[$($fp.name)] no downloaded report - reading live migration status from FLY..."
          $live = Join-Path $outDir ((& $safeName "$client $wl") + '_live-summary.json')
          Get-FlyLiveSummary -ProjectId $fp.id -Workload $wl | ConvertTo-Json -Depth 6 | Set-Content $live -Encoding UTF8
          Invoke-AreaBrief $row @('--summary', $live) $outDir
        } catch {
          Write-Warning "[$($fp.name)] $($_.Exception.Message)"
          $row.status = 'Failed: ' + $_.Exception.Message
        }
        $results += [pscustomobject]$row
      }
      $stillMissing = @($missing | Where-Object { $_ -notin @($projects | ForEach-Object { $platformArea[[int]$_.sourcePlatform] }) })
      if ($stillMissing.Count) { Write-Host "No FLY project for: $($stillMissing -join ', ')" }
    } catch { Write-Warning "Could not fill missing areas from live FLY data: $($_.Exception.Message)" }
    finally { try { Disconnect-Fly | Out-Null } catch {} }
  } elseif ($missing.Count) {
    Write-Host "No downloaded report for: $($missing -join ', ') (pass -Company to fill these from live FLY data)"
  }

  $pack = New-FlyBriefPack $results $client $outDir
  if ($pack) { $results = @($pack) + $results }
  Save-Results $results
  $results | Format-Table project, status, source -AutoSize -Wrap | Out-Host
  if ($pack -and $pack.pdfPath) { Write-Host "All-areas executive brief: $($pack.pdfPath)" -ForegroundColor Green }
  return
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
  $row = [ordered]@{ project = $p.project; client = $client; workload = $p.workload; status = ''; source = $null; folder = $null; htmlPath = $null; pdfPath = $null; summaryPath = $null }
  try {
    if ($SkipExport) {
      $out = Get-ChildItem $root -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending | Select-Object -First 1 -ExpandProperty FullName
      if (-not $out) { throw "No previous export found under $root" }
    } else {
      $out = Join-Path $root (Get-Date -Format 'yyyyMMdd-HHmm')
      New-Item -ItemType Directory -Force -Path $out | Out-Null

      # 1. Full migration report through the FLY portal (saved AOS sign-in), when that session is valid
      if (-not $NoPortal) {
        try { if (Get-FlyReportViaPortal -ProjectName $p.project -Workload $p.workload -OutFolder $out) { $row.source = 'FLY report (portal)' } }
        catch { Write-Warning "[$($p.project)] portal report not available: $($_.Exception.Message)" }
      }
      # 2. Fly.Client report export - only on request: FLY currently returns reportUrl = null (Fly.Client issue #16)
      if (-not $row.source -and $UseApiReport) {
        $cmd = "Export-Fly$($p.workload)MigrationReport"
        $exportArgs = @{ Project = $p.project; OutFolder = $out; FileType = 'CSV'; Include = $cfg.Include }
        if ($cfg.TzOffset) { $exportArgs.TimeZoneOffset = [int]$cfg.TzOffset }
        Write-Host "[$($p.project)] exporting $($p.workload) report through the API..."
        Invoke-WithRetry { & $cmd @exportArgs } "$($p.project) report export" | Out-Null
        if (Get-ChildItem $out -Recurse -Filter '*Mappings.csv') { $row.source = 'FLY report (API)' }
        else { Write-Warning "[$($p.project)] FLY returned no report file (reportUrl = null, Fly.Client issue #16)" }
      }
      Get-ChildItem $out -Filter *.zip -Recurse | ForEach-Object {
        Expand-Archive $_.FullName -DestinationPath (Join-Path $_.DirectoryName $_.BaseName) -Force
      }
      # 3. Live project data straight from the FLY API - always available, mapping-level figures only
      if (-not $row.source) {
        if (-not $p.id) { throw "Project id not found for '$($p.project)'" }
        Write-Host "[$($p.project)] reading live migration status from FLY..."
        Get-FlyLiveSummary -ProjectId $p.id -Workload $p.workload | ConvertTo-Json -Depth 6 |
          Set-Content (Join-Path $out 'fly-live-summary.json') -Encoding UTF8
        $row.source = 'FLY live data'
      }
    }
    $row.folder = $out
    $hasReport = [bool](Get-ChildItem $out -Recurse -Filter '*Mappings.csv')
    $liveJson = Join-Path $out 'fly-live-summary.json'
    if (-not $hasReport -and -not (Test-Path $liveJson)) { throw "No report or live summary in $out" }
    if (-not $row.source) { $row.source = if ($hasReport) { 'FLY report' } else { 'FLY live data' } }
    Write-Host "[$($p.project)] building brief from $($row.source)..."
    $nodeArgs = @($briefJs) + $(if ($hasReport) { @('--in', $out) } else { @('--summary', $liveJson) }) + @('--client', $client, '--workload', $p.workload, '--out', $out)
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
# One document for the executive team when the run covers a single customer
$clients = @($results | Where-Object status -eq 'OK' | ForEach-Object client | Sort-Object -Unique)
if ($clients.Count -eq 1) {
  $pack = New-FlyBriefPack $results $clients[0] (Join-Path $cfg.OutputRoot (& $safeName $clients[0]))
  if ($pack) { $results = @($pack) + $results; if ($pack.pdfPath) { Write-Host "All-areas executive brief: $($pack.pdfPath)" -ForegroundColor Green } }
}
Save-Results $results
$results | Format-Table project, status, source -AutoSize -Wrap | Out-Host

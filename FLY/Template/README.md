# FLY Executive Brief

One-page CIO/CISO brief built from AvePoint FLY migration reports. There are three ways to use it. All three use the same analysis and layout code, which lives in `FLY_Executive_Brief.html`.

## 1. Manual: upload an export
1. Open `FLY_Executive_Brief.html` in Edge or Chrome.
2. Choose the export folder. The tool needs the `…Mappings.csv` file, plus `…Site security.csv` and `…SharedLink.csv` for the security section. It skips the large "Site objects" files.
3. Enter the client name and check the internal domains it detected, then click **Rebuild**.
4. Click any text on the brief to edit it.
5. Click **Save as PDF** and choose *Save as PDF* with A4 paper.

## 2. Command line: one export folder
```
node fly-brief.js --in "<export folder>" --client "AEP Italia" [--internal "aep-italia.it,easyticketing.it"] [--out "<folder>"]
```
Writes `<client>_Executive_Brief.html`, `.pdf` and `.summary.json`. The PDF is rendered with headless Edge or Chrome; set `EDGE_PATH` if neither is in the default location.

## 3. Automated: straight from AvePoint FLY
Prerequisites:
- `Install-Module Fly.Client -Scope CurrentUser` (AvePoint's PowerShell module).
- An app registration in **AvePoint Online Services → App registrations** with Fly permissions and a certificate. Install the .pfx in `CurrentUser\My` or `LocalMachine\My`.
- The Fly public API URL for your data centre (see the Fly user guide).
- Node.js.

Steps:
1. Copy `fly-config.example.json` to `fly-config.json`.
2. Fill in the URL, client ID and certificate thumbprint.
3. Choose which projects to report on:
   - **Auto-discovery** (`"discover": { "enabled": true }`): every FLY project whose type is Exchange, OneDrive, SharePoint, Teams, M365 Groups or Teams Chat is included (limit with `workloads`). The client name is the part of the project name before " - " (so "AEP - OneDrive" gives "AEP"). Use `exclude` patterns to skip projects.
   - **`projects` entries** add projects or override discovered ones, e.g. to set a nicer client name ("AEP Italia") or the internal email domains. Entries still starting with `<` are ignored.
4. Run:
```
.\Invoke-FlyBriefs.ps1 Caternet           # a brief for each of Caternet's workloads
.\Invoke-FlyBriefs.ps1 Caternet -ListOnly # show which Caternet projects would run
.\Invoke-FlyBriefs.ps1                    # export + brief for every project
.\Invoke-FlyBriefs.ps1 -Project "AEP*"    # only matching projects
.\Invoke-FlyBriefs.ps1 -ListOnly          # show which projects would run; exports nothing
.\Invoke-FlyBriefs.ps1 -SkipExport        # rebuild briefs from the latest exports
```
Output goes to `<outputRoot>\<company>\<project name>\<yyyyMMdd-HHmm>\`, e.g. `Reports\Caternet\Caternet_-_Exchange\20261127-1830\Caternet_Exchange_Executive_Brief.pdf`. Exchange, Teams, Groups and Teams Chat briefs have no sharing-link data, so their security column flags what to review separately. Run `Test-FlyPrereqs.ps1 -TryConnect` first on a new machine to check the setup. To run it on a schedule, use Windows Task Scheduler with `pwsh -File Invoke-FlyBriefs.ps1`.

Known limitation: Fly.Client issue #16 reports that the export job can finish but return no report URL. When that happens, the script reports the project as failed instead of building an empty brief.

Thresholds such as the green/amber limits and the "stalled" days are in the `CFG` block at the top of the core script in `FLY_Executive_Brief.html`.

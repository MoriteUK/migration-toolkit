# FLY Executive Brief

A one-page CIO/CISO brief built from AvePoint FLY migration reports. You can use it four ways, and all of them use the same analysis and layout code, which lives in `FLY_Executive_Brief.html`.

| File | Purpose |
|---|---|
| `FLY_Executive_Brief.html` | The brief template, plus the `<script id="core">` analysis code. It also works on its own in a browser. |
| `fly-brief.js` | Node CLI. Builds the brief from one export folder and writes `.html`, `.pdf` and `.summary.json`. |
| `Invoke-FlyBriefs.ps1` | Exports reports from FLY with Fly.Client, then runs `fly-brief.js` for each project. |
| `Test-FlyPrereqs.ps1` | Read-only check of the machine. `-TryConnect` also signs in to FLY and lists the projects. |
| `fly-config.example.json` | Optional standalone config (certificate auth, explicit projects, discovery filters). |

## 1. In the Migration Toolkit app
Go to **Reports → FLY Executive Brief**.
1. Pick the customer. The list comes from Settings → Customers, and you pick by Prefix. **Name on Brief** fills in from the customer's AccountName and **Internal Email Domains** from its Domain. You can edit both.
2. Optionally choose an output folder. The default is `Migrations\Logs\FlyBriefs\`.
3. Click **Generate Brief**. The app builds one brief for every FLY project named `{Prefix} - {Workload}`, covering Exchange, OneDrive, SharePoint, Teams, M365 Groups and Teams Chat. The log streams live, and the global **Stop** button cancels the run.

The app uses the toolkit's Fly API settings (`%APPDATA%\FlyMigration\config.json`: Url, ClientId, EncSecret), so no extra setup is needed.

## 2. Manual: open the template and choose an export
1. Open `FLY_Executive_Brief.html` in Edge or Chrome.
2. Choose the export folder. The brief needs the `…Mappings.csv` file, plus `…Site security.csv` and `…SharedLink.csv` for the security section. It skips the large "Site objects" files.
3. Enter the client name, check the internal domains, then click **Rebuild**. Click any text to edit it.
4. Click **Save as PDF** and print to PDF on A4.

## 3. Command line: one export folder
```
node fly-brief.js --in "<export folder>" --client "AEP Italia" [--workload OneDrive] [--internal "aep-italia.it,easyticketing.it"] [--out "<folder>"] [--no-pdf]
```
The PDF is rendered with headless Edge or Chrome. Set `EDGE_PATH` if neither is in its default location. Requires Node 20 or later.

## 4. Command line: straight from FLY
```
.\Invoke-FlyBriefs.ps1 Caternet                                  # a brief for each Caternet workload
.\Invoke-FlyBriefs.ps1 AEP -ClientName "AEP Italia" -InternalDomains "aep-italia.it,easyticketing.it"
.\Invoke-FlyBriefs.ps1 Caternet -ListOnly                        # show which projects would run
.\Invoke-FlyBriefs.ps1 -Project "AEP - OneDrive"                 # one project
.\Invoke-FlyBriefs.ps1 Caternet -SkipExport                      # rebuild from the latest exports
.\Invoke-FlyBriefs.ps1 Caternet -OutputRoot D:\Briefs -NoPdf -ResultPath .\results.json
```
**Config.** The script uses `fly-config.json` from this folder, but only if that file contains real credentials (`EncSecret` or a certificate thumbprint). Otherwise it falls back to the toolkit config and auto-discovers every project. To use certificate auth, explicit project entries or `discover.exclude` patterns, copy `fly-config.example.json` to `fly-config.json`. That file is git-ignored.

**Matching.** The company is the part of the project name before " - ". Projects that don't follow `X - Workload` (for example "CSI to Volaris Exchange") can only be selected with `-Project`.

**Output.** Briefs are written to `<OutputRoot>\<client>\<project>\<yyyyMMdd-HHmm>\<client>_<workload>_Executive_Brief.{html,pdf,summary.json}`. Exchange, Teams, Groups and Teams Chat briefs have no sharing-link data, so their security column flags what to review separately.

**Known limitation.** Fly.Client issue #16: an export job can finish without returning a report URL. When that happens, the project is reported as failed instead of producing an empty brief.

Thresholds such as the green/amber limits and the "stalled" days are in the `CFG` block at the top of the core script in `FLY_Executive_Brief.html`.

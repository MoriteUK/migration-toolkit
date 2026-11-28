# FLY Executive Brief

A one-page CIO/CISO brief built from AvePoint FLY migration reports. You can use it four ways, and all of them use the same analysis and layout code, which lives in `FLY_Executive_Brief.html`.

| File | Purpose |
|---|---|
| `FLY_Executive_Brief.html` | The brief template, plus the `<script id="core">` analysis code. It also works on its own in a browser. |
| `fly-brief.js` | Node CLI. Builds the brief from one export folder and writes `.html`, `.pdf` and `.summary.json`. |
| `Invoke-FlyBriefs.ps1` | Gets each project's data from FLY (full report when available, otherwise live API data), then runs `fly-brief.js`. |
| `Test-FlyPrereqs.ps1` | Read-only check of the machine. `-TryConnect` also signs in to FLY and lists the projects. |
| `fly-config.example.json` | Optional standalone config (certificate auth, explicit projects, discovery filters). |

## 1. In the Migration Toolkit app
Go to **Reports → FLY Executive Brief**.
1. Pick the customer from Settings → Customers. The brief is titled with the customer prefix.
2. Optionally choose an output folder. The default is `Migrations\Logs\FlyBriefs\`.
3. Click **Generate Brief**. The app builds one brief for every FLY project named `{Prefix} - {Workload}`, covering Exchange, OneDrive, SharePoint, Teams, M365 Groups and Teams Chat. The log streams live, and the global **Stop** button cancels the run.

The app uses the toolkit's Fly API settings (`%APPDATA%\FlyMigration\config.json`: Url, ClientId, EncSecret), so no extra setup is needed.

**Where the figures come from.** For each project, in this order:
1. **Full FLY migration report through the FLY portal**, using the saved AOS sign-in (`VGMigrations\auth\storageState.json`, the same session AOS Setup uses). This gives every figure, including permissions and sharing links. *Not built yet:* it needs `fly-portal-report.js`, which requires a current AOS sign-in to develop against. Until that exists, this step is skipped. Once it's in place, an expired sign-in is reported in the log and the run falls through to step 3.
2. **Fly.Client report export.** Only with `-UseApiReport`. FLY currently finishes these jobs but returns `reportUrl: null` (Fly.Client issue #16, open since Oct 2026), so it is off by default.
3. **Live FLY API data.** Always available, and takes about 5 seconds per project. It gives mapping-level figures that match the report: completed, with exceptions, failed, failed-with-data, GB migrated and job dates. Item counts, permissions and sharing links are not in the API, so the brief says so and flags an access review instead. The preview and run results show which source each brief used.

### The all-areas executive brief (for emailing to the executive team)
Every run that covers one customer also writes `<Customer>_Migration_Executive_Brief.pdf` and `.html`. Page 1 is a one-page summary of all six areas: status, % migrated, data, open items, permissions, anyone-edit links, a one-line highlight per area, and the decisions merged across areas. After that comes one page per migrated area. The HTML is static, with no scripts, so it opens cleanly as an email attachment. An area with no mappings in FLY is listed on the summary page but gets no page of its own.

From reports you downloaded from the FLY portal (full figures, including permissions and sharing links):
```
.\Invoke-FlyBriefs.ps1 -ExportFolder "E:\Avepoint\Reports\Extracts" -Company Amickiel
```
Zips are expanded and each `*Mappings.csv` is paired only with its own detail files (same workload, same export time), so a folder holding every area, or copies of the same extract, never mixes data between areas. Areas with no downloaded report are filled from live FLY data for that customer's project. Output goes to `<ExportFolder>\Executive Brief\` (or `-OutputRoot`).

### Fly TechTeam Report (Reports → Fly TechTeam Report)
An interactive HTML for the tech team to work through problems mid-migration, before the management PDF. HTML only, with no PDF. It opens automatically when ready.
- **Drill-down:** All areas → an area → **Issue types** (grouped by FLY error code, with item and user counts) or **Users & sites** (errors, warnings and security per mailbox, OneDrive, site, team or chat) → the individual items (file, mail item, permission, sharing link, chat member) with FLY's status, error code and message. **Mappings** lists each mapping's last-job result. Search, plus severity and kind filters. The browser's Back button works.
- **Severities:** Error = failed. Warning = migrated with problems. Security = anonymous edit links carried over. Info = not supported / informational. Successful and skipped items are not listed. Very large issue types list their first 1,500 items; the counts include everything.
- **Sources:** with a **Downloaded FLY Reports Folder**, each downloaded area is read down to item level (this reads every object file, about 35 seconds for Amickiel's ~3.5M rows). Areas without a download, or all areas if no folder is given, show each user or site's **live** migration status from FLY.
- **Fixed / Ignored:** mark an issue type, or a single item, as Fixed or Ignored, with a note.
  - **Ignored everywhere** applies an error code to every area and every customer. Use it for errors that are expected when one domain is split into many, e.g. `EX-UserNotInMapping`, where the meeting organiser is mapped to another domain.
  - Nothing is ignored until someone on the tech team chooses to. Every issue-type row and item row has an **Ignored** tick box; untick it to re-open something you think can be fixed. Unticking a single item inside an ignored type re-opens just that item.
  - Fixed / ignored rows stay visible, greyed out. Untick *Show fixed & ignored* at the top to hide them. Counts, badges and the CSV show only what is still open.
  - Marks are saved in the browser on that PC, keyed by area + error code (+ user + item), so they survive regenerating the report.
  - **Export marks / Import marks** shares them with the team. **Export open issues (CSV)** lists the open items for tickets.
- Output: `<folder>\TechTeam Report\<Customer>_TechTeam_Report.html`, or `Logs\FlyBriefs\<Customer>\TechTeam Report\` without a folder.
- Command line: `.\Invoke-FlyBriefs.ps1 -TechReport -Company Amickiel [-ExportFolder "E:\Avepoint\Reports\Extracts"]`. Built by `fly-issues.js` into the `FLY_TechTeam_Report.html` template.

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
.\Invoke-FlyBriefs.ps1 -ExportFolder "C:\Users\me\Downloads\Caternet reports" -ClientName Caternet   # from portal downloads, no API
```
**Config.** The script uses `fly-config.json` from this folder, but only if that file contains real credentials (`EncSecret` or a certificate thumbprint). Otherwise it falls back to the toolkit config and auto-discovers every project. To use certificate auth, explicit project entries or `discover.exclude` patterns, copy `fly-config.example.json` to `fly-config.json`. That file is git-ignored.

**Matching.** The company is the part of the project name before " - ". Projects that don't follow `X - Workload` (for example "CSI to Volaris Exchange") can only be selected with `-Project`.

**Output.** Briefs are written to `<OutputRoot>\<client>\<project>\<yyyyMMdd-HHmm>\<client>_<workload>_Executive_Brief.{html,pdf,summary.json}`. Exchange, Teams, Groups and Teams Chat briefs have no sharing-link data, so their security column flags what to review separately.

**Known limitation.** Fly.Client issue #16: an export job can finish without returning a report URL. FLY's own `GET /downloadcenter/reporturl` returns `{"reportUrl": null}`, so this is not a client bug. Briefs fall back to live API data. From a shell, `-ExportFolder` also builds full briefs from reports you have downloaded from the portal yourself.

Thresholds such as the green/amber limits and the "stalled" days are in the `CFG` block at the top of the core script in `FLY_Executive_Brief.html`.

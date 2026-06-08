# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repo is

This is the **Migration Toolkit distribution package** — a self-contained folder that gets copied to any Windows machine to install and run the toolkit. It contains two sub-projects that work together:

- `MigrationToolkit-Web/` — Electron desktop application (the UI shell)
- `VGMigrations/` — PowerShell automation scripts (the actual migration engine)

## Running the app

```powershell
# From MigrationToolkit-Web/
cd MigrationToolkit-Web
npm start          # production mode
npm run dev        # development mode (sets NODE_ENV=development)
```

Press `Ctrl+Shift+I` inside the running app to open DevTools.

There are no automated tests. Verification is done by running the app and exercising the UI.

## Building a distributable

```powershell
cd MigrationToolkit-Web
npm run build      # outputs NSIS installer to MigrationToolkit-Web/dist/
```

## Architecture

### Electron process model

```
main.js (main process)
  ├── Registers all IPC handlers
  ├── Spawns pwsh.exe child processes for every PS script
  ├── Writes timestamped logs to %APPDATA%\FlyMigration\Logs\
  └── preload.js (context bridge)
        └── exposes window.electronAPI to renderer
              └── renderer.js + index.html (renderer process)
                    └── Single-page app — all views are <div>s hidden/shown by switchView()
```

### Two IPC execution modes

`stream-powershell` — streams stdout/stderr in real-time via `ps-output` IPC events. Used for all long-running operations (migration stages, AOS setup, app registration). The renderer appends each chunk to a `<pre>` log panel.

`execute-powershell` — buffers output and resolves with the full result. Used for quick queries (connection test, version check, get-migration-data).

`launch-script` — opens the script in a **new visible PowerShell window** via `cmd /c start`. Used for legacy standalone scripts (discovery, runner, etc.) that have their own WinForms UI.

### PowerShell scripts

All scripts dot-source `lib.ps1` first. Key things `lib.ps1` provides:

- **`$script:FlyWorkloadDefs`** — single authoritative map of all workloads (Exchange, SharePoint, OneDrive, Teams, TeamChat, Groups) to their `Fly.Client` cmdlet names. Both the Connections view and migration stage scripts reference this.
- **`Find-FlyDestinationConnection`** — looks up the destination connection in the Fly API by customer prefix + workload type. If none exist, auto-creates them by calling `fly-connector.js` (Playwright).
- **`Find-FlyPolicy`** — looks up the migration policy for a workload; prefers the default policy.
- **`Invoke-FlyConnectorCreate`** — calls `fly-connector.js --mode=create` to drive the AOS portal via Playwright and create destination connections.
- WinForms design system constants (colours, fonts) used by scripts with their own GUI windows.

### fly-connector.js

A Node.js + Playwright script that automates the **AvePoint Online Services (AOS)** portal. Two modes:
- `--mode=login` — opens a browser for the user to sign in; saves session to `VGMigrations/auth/storageState.json`
- `--mode=create` — reads JSON tasks from stdin, creates destination connections in the Fly portal headlessly using the saved session

### Config files

| Path | Contents |
|------|----------|
| `%APPDATA%\FlyMigration\config.json` | Fly API URL, ClientId, EncSecret (DPAPI-encrypted), PortalUrl, Customers array (Prefix, AccountName, Domain, SharePointAdminUrl) |
| `%LOCALAPPDATA%\FlyMigration\shared-config.json` | AOS tenant details: TenantName, TenantSearch, AppProfileName — shared between scripts |
| `VGMigrations\version.json` | App version string |
| `VGMigrations\auth\storageState.json` | Playwright browser session for AOS (created by Login-AOS.ps1 / Aos-SignIn.ps1) |

The `ClientSecret` is encrypted via `Encrypt-Secret.ps1` (DPAPI / `ConvertFrom-SecureString`) and stored as `EncSecret`. It can only be decrypted on the same Windows user account that encrypted it.

### Navigation / view model

`renderer.js::switchView(viewName)` hides all `.view-container` divs and shows the target. The map is defined inline — add a new entry there and a matching `id` on the HTML element to add a view. View-specific data loading (dropdowns, etc.) is triggered inside `switchView` by checking `viewName`.

### Migration workflow (Connections view)

The Connections view in the UI drives a sequential workflow per customer:

1. **Create Projects** → `New-FlyProject.ps1` (once per workload)
2. **Import Mappings** → `Import-FlyMappings.ps1` (once per workload, needs a CSV mapping file)
3. **Verify / Pre-Scan / Full / Incremental** → `Start-FlyMigrationStage.ps1 -Stage <name>`

`Start-FlyMigrationWorkflow.ps1` runs steps 1–3 in a single call. All of these stream output back via `ps-output`.

### Packaging / deployment

`VGMigrations\New-Package.ps1` builds the distribution package (this folder). `VGMigrations\Deploy-ToServer.ps1` pushes it to a file server. The version is stored in `VGMigrations\version.json`; `Check-Updates.ps1` compares it against the server copy.

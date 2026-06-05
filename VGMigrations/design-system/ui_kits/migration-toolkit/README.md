# UI Kit — Migration Toolkit (desktop)

A high-fidelity HTML/React recreation of the **Fly Migration Toolkit** WinForms
app (`MoriteUK/AvepointFlyUtility`). It reproduces the look and interaction of
the real windows using clean, reusable components — cosmetic, not production
logic.

## Run it
Open `index.html`. It boots on the **Launcher** and is click-through:

- **Discovery / Misc Scripts / Domain Removal** tiles → matching sub-menus.
- **AvePoint Fly** → the numbered migration menu.
  - **1. Create App Registration** → a working **workflow form** (card + fields
    + live coloured **log console**).
  - **5. Monitor Projects** → the **Project Monitor** dark data grid (select a
    project, toggle auto-refresh, click rows).
- The **⚙ gear** in any header opens the tabbed **Settings** dialog.

## Files
| File | Contents |
|---|---|
| `index.html` | Loads React + Babel and all component scripts. |
| `chrome.jsx` | Primitives: `Window`, `Header`, `Footer`, `Tile`, `Button`, `Card`, `Field`, `Select`, `Checkbox`, `StatusDot`, plus the `FLY` token object. |
| `console.jsx` | `LogConsole` (coloured `Write-Log` output) and `MonitorGrid` (dark `DataGridView` with row tints). |
| `screens.jsx` | `Launcher`, `SubMenu`, `FlyMenu`, `AppRegForm`, `ProjectMonitor`, `SettingsDialog`. |
| `app.jsx` | Navigation shell + breadcrumb. |

## Conventions
- Each `.jsx` exports its components to `window` (Babel scopes are per-script).
- All colours/fonts come from the `FLY` object in `chrome.jsx`, which mirrors
  `colors_and_type.css` and the `$clr*`/`$Font*` values in `lib.ps1`.
- Components take plain props; state is local and illustrative only.

## Fidelity notes
- The native app is **square-cornered** WinForms. This kit applies a *restrained*
  2–4px radius and soft window shadow for on-screen polish — the brand (blue,
  Segoe UI, banner/footer frame, dark console) is unchanged. Drop to 0px radius
  and remove the shadow for a 1:1 native match.
- **Segoe UI / Consolas** are Windows system fonts; on non-Windows previews the
  stack falls back to the nearest system + monospace faces.

# UI Kit — Migration Toolkit · Fluent refresh

A modern **Microsoft-365-admin-center** restyle of the Fly Migration Toolkit,
built to answer *"can it look more professional?"* — **yes**. Same brand (AvePoint
blue, Segoe UI, the paper-plane mark), but reorganised into a contemporary
desktop-app shell.

> This is a **redesign proposal**, presented next to the faithful original in
> `../migration-toolkit/`. Compare the two side by side.

## Run it
Open `index.html`. It's a click-through app:

- **Left nav rail** replaces window-per-screen navigation — one persistent shell.
- **Home** → a migration **dashboard**: KPI stats, per-workload progress bars
  with status pills, and tool cards.
- **AvePoint Fly** → the numbered workflow as a **stepped checklist** (done /
  in-progress states). Step 1 opens a real form.
- **Project Monitor** → a **light Fluent list** with status pills and a command
  bar (Refresh, Export, auto-refresh toggle, project picker) — replacing the raw
  dark grid.
- **⚙ Settings** (top bar) → a **right slide-over panel** with pivot tabs.

## What changed (and why it reads as "professional")
| Before (native) | After (Fluent refresh) |
|---|---|
| Stacked full-width tiles | Left nav rail + content area (M365 shell) |
| New window per screen | Single persistent shell, breadcrumb context |
| Dense 13px text, ad-hoc spacing | 14px base, consistent 16/20/28 type rhythm |
| Dark data grid for status | Light list with coloured **status pills** |
| Plain menus | **Dashboard** with KPIs + progress bars |
| Text-only controls | Consistent **command bars** with icons |

## Files
| File | Contents |
|---|---|
| `index.html` | Loads React + Babel and the scripts below. |
| `fluent-shell.jsx` | Tokens (`F`), `Icon`, `Button`, `Card`, `Pill`, `Field`, `Select`, `Toggle`, `SearchBox`. |
| `fluent-chrome.jsx` | `TopBar`, `NavRail`, `PageHeader`, `Breadcrumb`, `CommandBar`, `SettingsPanel`. |
| `fluent-screens.jsx` | `Home` dashboard, `FlyMenu` stepped workflow, `Stat`, `Page`. |
| `fluent-monitor.jsx` | `Monitor` list + `AppReg` workflow form. |
| `fluent-domain.jsx` | `DomainRemoval` (guided 3-step workflow), `MessageBar`, `ConfirmRemoveDialog`. |
| `fluent-app.jsx` | Navigation shell. |

## ⚠️ WinForms feasibility notes (read before building)
You asked to keep this **buildable in WinForms**. Most of it is — the value is in
*layout and structure*, not effects. Mapping to the real app:

- **Buildable as-is:** the nav-rail + content layout (docked `Panel`s),
  command bars (`FlowLayoutPanel` of flat buttons), the light status list
  (`DataGridView` or `ListView` with owner-drawn pill cells), the dashboard
  (custom-painted panels + progress bars), Segoe UI type scale, the breadcrumb,
  the settings slide-over (a docked panel).
- **Icons:** rendered here as inline SVG in *Segoe-Fluent spirit*. In the real
  app use **Segoe Fluent Icons** (Win 11) or **Segoe MDL2 Assets** (Win 10) —
  the system icon font — by setting the glyph codepoint. **No image assets
  needed. Flagged substitution:** the SVGs here are a cross-platform stand-in so
  the preview renders on any OS.
- **Web-only polish to drop for 1:1 native:** the 4–6px **corner radius**, the
  faint card **drop-shadows**, and the slide-in **animation** are not native
  WinForms. Set `F.radius`/`F.radiusSm` to `0` and `F.shadowCard`/`shadowRaise`
  to `'none'` in `fluent-shell.jsx` for a hard-edged, fully-native look — it
  still reads as a clean, modern app.
- **Status pills** need owner-draw (rounded rect) in WinForms; a square chip is
  the zero-effort fallback.

## Domain Removal (built out)
The most complex / destructive surface is fully designed in `fluent-domain.jsx`:
- A **guided 3-step workflow** (Update on-prem UPNs → Run AD sync → Remove domain)
  as a vertical stepper. Steps unlock in order, show done/running/ready/waiting
  states, and stream into a **live activity log**.
- A Fluent **MessageBar** explaining why order matters (flips to success when done).
- A **type-to-confirm** dialog for the destructive removal — the "Remove domain"
  button stays disabled until the operator types the exact domain name.
- **Standalone tools** for running any single operation directly.

WinForms mapping: the stepper is a stack of custom panels; the MessageBar is a
tinted `Panel`; the confirm dialog is a modal `Form` with a `TextBox.TextChanged`
gate on the OK button — all straightforward.

## Honest scope
Discovery, Reports and Misc are shown as **stubs** — the shell, nav and component
patterns apply to them, but they weren't part of this sample. Ask if you'd like
any of them built out.

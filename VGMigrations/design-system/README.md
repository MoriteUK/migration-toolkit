# Fly Migration Toolkit — Design System

A design system distilled from **MoriteUK / AvepointFlyUtility** — an internal
Microsoft 365 tenant‑migration toolkit. It captures the toolkit's colours,
typography, components and voice so that new screens, dialogs, slides and docs
can be built that look and feel like part of the same product.

> **Source repository:** <https://github.com/MoriteUK/AvepointFlyUtility>
> The system was reverse‑engineered from the live PowerShell / WinForms source
> (`lib.ps1`, `menu.ps1`, `main-menu.ps1`, `monitor.ps1`). If you have access,
> explore the repo to build higher‑fidelity designs — the `*.ps1` form
> definitions are the authoritative source of every colour, font and layout
> value reproduced here.

---

## What the product is

The **Fly Migration Toolkit** ("Migration Tools — AvePoint Fly Edition") is a
Windows desktop application that IT engineers use to migrate Microsoft 365
tenants using the **AvePoint Fly** migration platform. It is **not** a web app:
it is a set of PowerShell scripts that draw a **WinForms** GUI, packaged into an
`.exe` and auto‑updated from GitHub.

It is operator software — dense, functional, run by a migration engineer on a
jump box, not a polished consumer product. The design language is therefore
**Microsoft‑Fluent‑adjacent**: a blue banner, flat buttons, white cards on a
cool‑grey canvas, a dark "console" surface for logs and live data grids.

### The surfaces (what a user actually sees)

| Surface | Role |
|---|---|
| **Launcher** (`main-menu.ps1`) | Top‑level window. Four large tiles: Discovery, AvePoint Fly, Misc Scripts, Domain Removal. An amber "update available" banner appears when a newer version is on GitHub. |
| **Sub‑menus** | Discovery / Misc / Domain Removal each open a same‑styled child window of stacked tiles with muted one‑line descriptions. |
| **AvePoint Fly menu** (`menu.ps1`) | The migration workflow: 1 Create App Registration → 2 Setup AOS Tenant → 3 Connections & Mappings → 4 Reports → 5 Monitor. |
| **Project Monitor** (`monitor.ps1`) | Live **dark data grid** of per‑workload migration status (Total / Not Started / In Progress / Complete / Failed / Warnings) with a connection status bar and auto‑refresh. |
| **Workflow forms** | App Registration, AOS Setup, Migration Runner, Reports — card panels with labelled fields, primary/secondary buttons and a coloured **log console**. |
| **Settings dialog** | Tabs: Config (Fly API URL, Client ID/Secret), Customer, Workloads, Discovery. Reached via the ⚙ gear in every header. |

### Workloads it migrates
SharePoint · Exchange · OneDrive · Teams · Teams Chat · M365 Groups
— plus tenant clean‑up: device removal, domain removal, UPN updates, GAL hiding.

---

## CONTENT FUNDAMENTALS

The toolkit's copy is **terse, imperative and engineer‑facing**. It assumes a
competent admin and never explains Microsoft 365 concepts.

- **Voice & person.** Instructions are imperative — *"Connect to the source
  tenant"*, *"Export discovery results"*. The UI rarely says "you"; it labels
  actions, not people. Confirmations address the operator directly and plainly:
  *"Trigger Azure AD Connect sync on VOL‑ane‑aad1? This will start a delta sync
  cycle."*
- **Casing.** **Title Case** for buttons, tiles and window titles
  (*"Create App Registration"*, *"Setup AOS Tenant & App"*, *"View Migration
  Reports"*). **Sentence case** for the muted subtitles beneath each tile
  (*"Register the Entra ID app and grant required API permissions"*). Log lines
  are sentence case after a bracketed level tag.
- **Numbered steps.** The core workflow is literally numbered in the labels —
  *"1. Create App Registration"*, *"2. Setup AOS Tenant & App"* — reflecting a
  strict run order. Preserve numbering when extending the flow.
- **Subtitles do the explaining.** Every tile pairs a Title‑Case action with a
  single muted sentence describing the outcome. Keep them to one line, ~8–11
  words, no terminal period in tile subtitles.
- **Status language.** Log levels are fixed tokens: `INFO`, `OK`, `WARN`,
  `ERROR`. Grid status buckets are fixed too: *Not Started, In Progress,
  Complete, Failed, Warnings*. Reuse these exact words — don't invent synonyms.
- **Domain jargon is expected.** Entra ID, AOS, UPN, GAL, delta sync, tenant,
  workload, mapping, pre‑scan, verification — used without expansion.
- **Tone.** Calm, precise, slightly cautionary around destructive actions
  (domain removal, AD sync). Destructive operations are coloured red and gated
  behind a confirm dialog. No marketing language, no exclamation, no jokes.
- **Emoji.** Essentially none in‑app. The repo README uses a single 🔵 to denote
  "AvePoint Fly platform"; product UI uses **Unicode glyphs as functional icons**
  (gear ⚙, status dot ●), never decorative emoji. Do not add emoji to UI.

---

## VISUAL FOUNDATIONS

A cool, corporate, **Windows‑desktop** aesthetic. Think Microsoft 365 admin
center rendered in WinForms: flat, rectangular, blue‑and‑grey, with a dark
operator console for anything live or log‑like.

### Colour
- **Primary** is a single AvePoint blue `#0064b4`, used for the header banner,
  all primary buttons, grid headers and the 4px accent stripe on cards. Hover
  darkens to `#004e98`. There are **no gradients** in the UI chrome (the only
  gradient is inside the app *icon* artwork).
- **Canvas** is a cool light grey `#f0f2f7`; **surfaces** are pure white.
  Text is near‑black `#1c1c20`; secondary text a blue‑grey `#646c78`; hairlines
  `#d2d7e4`.
- **Dark console** `#1a1b26` is a distinct surface used for log output and the
  Project Monitor grid — light blue‑white text on near‑black, with coloured
  level tags. Footers are an even darker navy `#141826`.
- **Status:** green `#129b3c` success · amber `#c38700` warning · red `#c31e1e`
  error/destructive. In the dark grid, failed rows get a deep maroon tint and
  warning rows a deep amber tint.
- **Imagery vibe:** there is essentially no photography. The palette reads
  **cool and corporate** — blues and greys, high contrast, no warmth, no grain.

### Type
- **Segoe UI** everywhere (regular + Semibold), the native Windows UI face.
  **Consolas** for the log console and any code/IDs. No third font.
- Scale is small and dense, matching WinForms point sizes: 9pt body, 7.5pt
  caption, 8.5pt mono, 14pt semibold for titles and big tiles. See
  `colors_and_type.css` for the px‑equivalent tokens.
- Headings use Semibold, never a heavier weight; letter‑spacing is default
  except small uppercase captions, which get slight tracking.

### Layout, spacing & shape
- **Fixed‑width dialogs.** Menus are 480px wide with content inset ~40px;
  tiles are 400×90px stacked with a 6px gap, each followed by a 26px‑tall muted
  subtitle. Larger working windows (Monitor) are resizable with docked
  header / selector / footer rows.
- **Banner + footer frame.** Almost every window is a 56px blue banner on top
  (logo + title + gear) and a 46px navy footer at the bottom (Close at right).
  Content sits between.
- **Corners are square.** WinForms `FlatStyle` buttons and `FixedSingle`
  borders mean **0px radius** throughout the native app. (When recreating on the
  web you may apply a *very* small 2–4px radius for polish, but the native
  feel is hard‑edged — keep it minimal.)
- **Spacing** is pragmatic, in whole pixels: 6 / 8 / 12 / 16 / 26 / 40px recur.
  See the Spacing cards in the Design System tab.

### Borders, shadows & elevation
- **Hairline borders, not shadows.** Cards are a white panel inside a 1px
  `#d2d7e4` border (implemented as a 1px border‑colour panel wrapping the white
  fill) with a **4px solid accent stripe down the left edge** — the signature
  card motif. WinForms has **no drop shadows**; elevation is communicated by
  borders and the blue banner, not blur.
- Text fields are white with a single `FixedSingle` 1px border. Focus is the OS
  default (no custom ring in‑app).
- No translucency or backdrop blur anywhere — it's an opaque, layered‑rectangle
  UI.

### Buttons & controls
- **Primary button:** accent‑blue fill, white text, Semibold, flat, no border,
  pointer cursor, hover → darker blue. **Secondary:** light grey `#e1e4ee`
  fill, ink text. **Destructive:** red fill (`#c31e1e` for tiles, `#c83737` for
  the footer Close).
- **Tiles** are oversized primary buttons (400×90) used for navigation.
- **Status dots** are a `●` glyph or 12px square coloured green/amber/red/grey.
- **Combo boxes / checkboxes / text fields** use native WinForms styling on a
  white field with the shared border colour.

### Motion
- **None to speak of.** WinForms here uses no animated transitions — windows
  appear, grids repopulate, logs append line‑by‑line. The only "motion" is
  hover colour changes on buttons (instant) and the log auto‑scrolling to the
  caret. When recreating on the web, keep motion minimal and functional: instant
  or ~120ms colour fades on hover, no bounces, no decorative animation. Respect
  `prefers-reduced-motion`.

### Hover / press states
- **Hover:** primary buttons darken (`#0064b4` → `#004e98`); the gear darkens
  the same way. Cursor becomes a hand on all clickable controls.
- **Press / selection:** in the data grid, selected rows use `#004e98` fill with
  white text. There is no scale/shrink press animation in the native app.

---

## ICONOGRAPHY

The toolkit is **almost icon‑free** — it leans on text labels, not pictograms.

- **App / brand mark.** A single raster icon, `FlyMigration.ico` (256px, also
  exported to `assets/fly-logo-256.png` / `-64.png`): a **white paper‑plane /
  arrow** flying right on a **blue rounded‑square** gradient — the "Fly"
  migration‑in‑motion mark. It appears at the left of every header banner
  (scaled 30–40px) and as the window/taskbar icon. This is the only piece of
  brand artwork in the product.
- **Functional glyphs, not an icon set.** The UI uses a tiny number of **Unicode
  glyphs from the Segoe UI Symbol font** as functional icons: the **gear ⚙**
  (`U+2699`) for Settings, drawn into a white circle, and the **filled dot ●**
  (`U+25CF`) for status indicators. There is **no bundled icon font, SVG sprite
  or PNG icon library** in the repo.
- **No decorative iconography.** Tiles, fields and menus carry no leading icons —
  they are text + a muted subtitle. Don't introduce a Material/Fluent icon set
  to "fill" buttons; it would be off‑brand.
- **Recreation guidance.** For web recreations that genuinely need an icon (e.g.
  a close ✕, a refresh ↻, status chevrons), use **Segoe UI Symbol / Segoe
  Fluent** Unicode glyphs first to stay native. If a richer set is unavoidable,
  the closest CDN match in spirit is **Microsoft's Fluent UI System Icons**
  (regular weight, line style) — **flag any such substitution**, as it is *not*
  present in the source product.

---

## VISUAL ASSETS

| File | What it is |
|---|---|
| `assets/FlyMigration.ico` | Original Windows app icon (256px, blue square + white plane). |
| `assets/fly-logo-256.png` | PNG export of the app icon for web/banner use. |
| `assets/fly-logo-64.png`  | Small PNG export for compact headers / favicons. |

There are no other logos, illustrations, photographs or background images in the
source product — the design is deliberately chrome‑only.

---

## INDEX — what's in this design system

| Path | Contents |
|---|---|
| `README.md` | This file — context, content & visual foundations, iconography, index. |
| `SKILL.md` | Agent‑Skill entry point (works in Claude Code too). |
| `colors_and_type.css` | All colour + type tokens as CSS variables and helper classes. |
| `assets/` | App icon (`.ico`) and PNG logo exports. |
| `preview/` | Small HTML cards rendered in the Design System tab (colours, type, spacing, components). |
| `ui_kits/migration-toolkit/` | High‑fidelity HTML/React recreation of the toolkit UI — `index.html` (interactive) plus modular JSX components. See its own `README.md`. |

### UI kits
- **`ui_kits/migration-toolkit/`** — the desktop toolkit: Launcher, AvePoint Fly
  menu, Project Monitor (dark grid) and Settings dialog, built from reusable
  `Header`, `Footer`, `Tile`, `Card`, `Button`, `Field`, `StatusDot`,
  `LogConsole` and `MonitorGrid` components.

---

## A note on "look & feel better"

This system documents the product **as it is today** (a faithful WinForms
aesthetic). The HTML/React UI kit recreates those screens pixel‑close while
giving you clean, reusable components to build from. Where the web naturally
allows it, the kit applies *restrained* polish — consistent spacing, smooth
hover fades, optional 2–4px corner softening — **without** changing the brand:
same blue, same Segoe UI, same banner/footer frame, same dark console. Treat any
larger redesign (new layout grid, added elevation, refreshed palette) as a
separate, explicit decision — see the **ask** at the end of the build.

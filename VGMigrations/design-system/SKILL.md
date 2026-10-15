---
name: fly-migration-design
description: Use this skill to generate well-branded interfaces and assets for the Fly Migration Toolkit (MoriteUK/AvepointFlyUtility) — a Microsoft 365 tenant-migration desktop tool — either for production or throwaway prototypes/mocks/etc. Contains essential design guidelines, colors, type, fonts, assets, and UI kit components for prototyping.
user-invocable: true
---

Read the `README.md` file within this skill, and explore the other available files.

If creating visual artifacts (slides, mocks, throwaway prototypes, etc), copy assets out and create static HTML files for the user to view. If working on production code, you can copy assets and read the rules here to become an expert in designing with this brand.

If the user invokes this skill without any other guidance, ask them what they want to build or design, ask some questions, and act as an expert designer who outputs HTML artifacts _or_ production code, depending on the need.

## Quick map
- `README.md` — product context, content & visual foundations, iconography, file index. **Start here.**
- `colors_and_type.css` — all colour + type tokens as CSS variables (`--fly-accent`, `--fly-bg`, etc.) and helper classes. Link or copy this into any HTML you build.
- `assets/` — the `FlyMigration` app icon (`.ico` + PNG exports). The only brand artwork.
- `preview/` — small specimen cards (colours, type, spacing, components) for reference.
- `ui_kits/migration-toolkit/` — interactive HTML/React recreation of the toolkit (Launcher, menus, Project Monitor, Settings) with reusable `chrome.jsx` / `console.jsx` / `screens.jsx` components. Reuse these as the starting point for new screens.

## Essentials to remember
- One primary blue (`#0064b4`), cool-grey canvas (`#f0f2f7`), white surfaces, dark console (`#1a1b26`). No gradients in chrome, no drop shadows natively.
- **Segoe UI** for UI, **Consolas** for logs/code. Small, dense scale.
- Signature motifs: 56px blue banner + 46px navy footer frame; oversized accent **tiles** with muted subtitles; white **cards with a 4px left accent stripe**; dark **log console** with coloured `[INFO]/[OK]/[WARN]/[ERROR]` tags; dark **monitor grid** with maroon/amber row tints.
- Voice: terse, imperative, Title-Case actions + sentence-case subtitles. No emoji, no marketing tone. Numbered workflow steps.
- Icons: essentially none — text labels, plus Unicode glyphs (⚙ gear, ● status dot). Don't add an icon set unless asked; if you must, prefer Segoe Fluent / Microsoft Fluent System Icons and flag it.

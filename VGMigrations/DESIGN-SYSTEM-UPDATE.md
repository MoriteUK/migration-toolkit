# Design System Update — Version 2.1.14

## Overview

All PowerShell WinForms UI components have been refined to match the official design system specifications documented in `design-system/`.

## What Changed

### 1. Color Standardization (lib.ps1)

**Added complete design system color palette:**
- Base colors: `$clrAccentHover`, `$clrAccentTint`
- Dark surfaces: `$clrFooter`, `$clrFooterAlt`
- Status colors: `$clrCloseRed`, `$clrBannerWarn`
- Log console colors: `$clrLogTimestamp`, `$clrLogInfo`, `$clrLogOK`, `$clrLogWarn`, `$clrLogError`, `$clrLogBody`
- Grid colors: `$clrGridText`, `$clrGridLine`
- Row tints: `$clrRowFailBg`, `$clrRowFailFg`, `$clrRowWarnBg`, `$clrRowWarnFg`

**Replaced all hardcoded FromArgb() calls with named variables:**
- `FromArgb(0, 78, 152)` → `$clrAccentHover`
- `FromArgb(20, 24, 38)` → `$clrFooter`
- `FromArgb(200, 55, 55)` → `$clrCloseRed`
- `FromArgb(220, 230, 248)` → `$clrAccentTint`
- Plus 15+ additional color mappings

### 2. Typography Improvements (lib.ps1)

**Added missing font definitions:**
- `$FontSub` - 8.5pt Segoe UI for tile subtitles (--text-sub 12px)
- `$FontTile` - 14pt Segoe UI Semibold for large navigation tiles (--text-tile 19px)

**Updated font references:**
- Replaced inline `New-Object System.Drawing.Font()` calls with shared variables
- Changed gear icon font from "Segoe UI" to "Segoe UI Symbol" for proper glyph rendering

### 3. Files Updated

**Core UI Framework:**
- `lib.ps1` - Added complete design system color and font tokens

**Main Menus:**
- `main-menu.ps1` - Updated all colors and fonts to use design system variables
- `menu.ps1` (AvePoint Fly) - Standardized colors and typography
- `discovery-menu.ps1` - Applied design system styling

**Specialized Windows:**
- `settings.ps1` - Updated header fonts and DataGridView colors
- `monitor.ps1` - Complete dark grid styling overhaul with proper status row tints

## Design System Compliance

All UI elements now precisely match the specifications in:
- `design-system/colors_and_type.css` - Complete color and typography reference
- `design-system/README.md` - Visual foundations and usage guidelines

### Color Mapping

| CSS Variable | PowerShell Variable | RGB Value |
|--------------|---------------------|-----------|
| `--fly-accent` | `$clrAccent` | #0064b4 (0, 100, 180) |
| `--fly-accent-hover` | `$clrAccentHover` | #004e98 (0, 78, 152) |
| `--fly-accent-tint` | `$clrAccentTint` | #dce6f8 (220, 230, 248) |
| `--fly-footer` | `$clrFooter` | #141826 (20, 24, 38) |
| `--fly-footer-alt` | `$clrFooterAlt` | #1c2030 (28, 32, 48) |
| `--fly-close-red` | `$clrCloseRed` | #c83737 (200, 55, 55) |
| `--fly-grid-text` | `$clrGridText` | #bed2ff (190, 210, 255) |
| `--fly-grid-line` | `$clrGridLine` | #2d374b (45, 55, 75) |
| `--fly-row-fail-bg` | `$clrRowFailBg` | #461919 (70, 25, 25) |
| `--fly-row-fail-fg` | `$clrRowFailFg` | #f07d7d (240, 125, 125) |
| `--fly-row-warn-bg` | `$clrRowWarnBg` | #41300c (65, 48, 12) |
| `--fly-row-warn-fg` | `$clrRowWarnFg` | #ebc350 (235, 195, 80) |

### Typography Mapping

| CSS Variable | PowerShell Variable | Font Specification |
|--------------|---------------------|-------------------|
| `--text-body` | `$FontBody` | Segoe UI 9pt (13px) |
| `--text-bold` | `$FontBold` | Segoe UI Semibold 9pt |
| `--text-cap` | `$FontCap` | Segoe UI Semibold 7.5pt (10px) |
| `--text-mono` | `$FontMono` | Consolas 8.5pt (12px) |
| `--text-title` | `$FontTitle` | Segoe UI Semibold 14pt (19px) |
| `--text-sub` | `$FontSub` | Segoe UI 8.5pt (12px) |
| `--text-tile` | `$FontTile` | Segoe UI Semibold 14pt (19px) |

## Benefits

1. **Consistency** - All windows now use identical colors and fonts from a single source
2. **Maintainability** - Color changes now update globally by modifying one variable
3. **Design Accuracy** - UI now precisely matches documented design system
4. **Better Dark Mode** - Proper status row tinting in Project Monitor dark grid
5. **Professional Polish** - Unified hover states, consistent spacing, proper font weights

## Backward Compatibility

✅ **Fully backward compatible** - All changes are visual refinements only. No breaking changes to functionality or data structures.

## Testing

After updating, verify:
1. Main menu launches correctly
2. All tile buttons show proper hover effects (darker blue)
3. Footer Close buttons use correct red color
4. Project Monitor dark grid has proper status row colors (maroon for failures, amber for warnings)
5. Settings dialog DataGridView headers use accent tint color

---

**Version**: 2.1.14  
**Date**: 2026-06-03  
**Design System**: design-system/colors_and_type.css

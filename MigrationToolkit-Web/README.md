# Migration Toolkit - Web Edition

Modern Electron + React UI for the Migration Toolkit, featuring the official design system with a polished, contemporary look.

## Features

- ✨ **Modern UI** - Clean, contemporary interface matching the design system
- 🎨 **Design System** - Consistent colors, typography, and spacing
- ⚡ **Fast & Responsive** - Built with Electron for native desktop performance
- 🔧 **PowerShell Integration** - Seamlessly calls existing PowerShell scripts
- 🔄 **Auto-Updates** - Built-in update checking and installation
- ⚙️ **Settings** - Easy configuration management

## Prerequisites

- **Node.js 18+** (includes npm)
- **PowerShell 7+**
- **Windows 10/11**

## Quick Start

### 1. Install Dependencies

```powershell
cd C:\Temp\Scripts\MigrationToolkit-Web
npm install
```

This installs:
- Electron (desktop app framework)
- Electron Builder (for creating .exe installer)
- Cross-env (environment variable management)

### 2. Run in Development Mode

```powershell
npm run dev
```

This launches the app with DevTools open for debugging.

### 3. Run in Production Mode

```powershell
npm start
```

This launches the app without DevTools.

### 4. Build Standalone Executable

```powershell
npm run build
```

Creates an installer in the `dist\` folder:
- `Migration Toolkit Setup.exe` - Windows installer
- Portable version also available

## Project Structure

```
MigrationToolkit-Web/
├── main.js                 # Electron main process (backend)
├── preload.js             # Secure bridge between main and renderer
├── renderer.js            # UI logic and PowerShell integration
├── index.html             # Main application window
├── package.json           # Dependencies and build configuration
├── src/
│   ├── styles/
│   │   ├── design-system.css  # Color and typography tokens
│   │   └── main.css           # Application styles
│   └── lib/               # Utility modules (future)
└── public/
    └── icon.ico           # Application icon
```

## How It Works

### Electron Architecture

```
┌─────────────────────────────────────────────┐
│  Renderer Process (UI)                      │
│  - index.html, renderer.js, CSS             │
│  - Runs in browser-like environment         │
└──────────────┬──────────────────────────────┘
               │ IPC (Inter-Process Communication)
┌──────────────▼──────────────────────────────┐
│  Main Process (Backend)                     │
│  - main.js, Node.js APIs                    │
│  - Spawns PowerShell processes              │
└──────────────┬──────────────────────────────┘
               │ child_process.spawn()
┌──────────────▼──────────────────────────────┐
│  PowerShell Scripts                         │
│  - ../VGMigrations/*.ps1                    │
│  - Existing migration logic                 │
└─────────────────────────────────────────────┘
```

### PowerShell Integration

The app calls existing PowerShell scripts via:

```javascript
// renderer.js (UI) → IPC → main.js (backend) → PowerShell
window.electronAPI.launchScript('menu.ps1')
```

**Main process spawns PowerShell:**
```javascript
spawn('pwsh.exe', [
  '-NoProfile',
  '-ExecutionPolicy', 'Bypass',
  '-File', scriptPath
], { cwd: PS_SCRIPT_PATH })
```

### Design System

All colors and fonts come from `src/styles/design-system.css`:

```css
--fly-accent: #0064b4;        /* AvePoint blue */
--fly-accent-hover: #004e98;  /* Hover state */
--fly-bg: #f0f2f7;           /* Canvas background */
--font-ui: "Segoe UI", ...;  /* Primary font */
--text-tile: 19px;           /* Large button text */
```

## Configuration

### Fly API Settings

Click the ⚙ gear icon to configure:
- **Fly API URL**: `https://graph.avepointonlineservices.com/fly`
- **Client ID**: Your Entra ID app client ID
- **Client Secret**: Your app secret

Settings are saved to:
```
%APPDATA%\FlyMigration\config.json
```

### PowerShell Scripts Location

By default, the app looks for PowerShell scripts at:
```
C:\Temp\Scripts\VGMigrations\
```

To change this, edit `main.js`:
```javascript
const PS_SCRIPT_PATH = path.join(__dirname, '..', 'VGMigrations');
```

## Development

### Enable DevTools

```powershell
npm run dev
```

Or press **Ctrl+Shift+I** in the running app.

### Hot Reload

Currently requires manual restart. To enable hot reload:

```powershell
npm install --save-dev electron-reload
```

Add to `main.js`:
```javascript
require('electron-reload')(__dirname);
```

### Modify UI

Edit these files:
- **Layout**: `index.html`
- **Styling**: `src/styles/main.css`
- **Logic**: `renderer.js`
- **Colors**: `src/styles/design-system.css`

Changes take effect on app restart.

## Deployment

### Option 1: Development Mode (for testing)

```powershell
# Share the entire folder
Copy-Item -Path "C:\Temp\Scripts\MigrationToolkit-Web" -Destination "\\server\share\" -Recurse

# User runs:
npm install
npm start
```

### Option 2: Standalone Executable

```powershell
# Build installer
npm run build

# Distribute the installer:
dist/Migration Toolkit Setup.exe
```

User double-clicks to install, no npm required.

### Option 3: Portable Build

Edit `package.json` → `build.win.target`:
```json
"target": ["nsis", "portable"]
```

Then:
```powershell
npm run build
```

Creates `dist/Migration Toolkit.exe` (portable, no installer needed).

## Troubleshooting

### "npm: command not found"

Node.js not installed or not in PATH.

**Fix:**
```powershell
winget install OpenJS.NodeJS.LTS
# Restart terminal
```

### "pwsh.exe: command not found"

PowerShell 7 not installed.

**Fix:**
```powershell
winget install Microsoft.PowerShell
```

### Scripts don't launch

Check PowerShell script path in `main.js`:
```javascript
const PS_SCRIPT_PATH = path.join(__dirname, '..', 'VGMigrations');
```

Verify scripts exist at that location.

### Window appears blank

Open DevTools (Ctrl+Shift+I) and check console for errors.

Common causes:
- CSS file path incorrect
- JavaScript error in renderer.js

## Comparison: Web vs Classic

| Feature | Classic (WinForms) | Web (Electron) |
|---------|-------------------|----------------|
| **Technology** | PowerShell + WinForms | Electron + HTML/CSS/JS |
| **Look** | Native Windows | Modern, customizable |
| **Design System** | Variables in lib.ps1 | Full CSS design system |
| **Performance** | Fast startup | Slightly slower (Chromium) |
| **Customization** | Limited by WinForms | Fully customizable |
| **File Size** | ~2 MB | ~150 MB (includes Chromium) |
| **Updates** | PowerShell auto-update | Same (calls PowerShell) |

## Roadmap

- [ ] Add Project Monitor window (dark grid)
- [ ] Implement inline PowerShell execution (not just spawn)
- [ ] Add real-time log streaming
- [ ] Build React components for complex forms
- [ ] Add TypeScript for type safety
- [ ] Implement automatic updates (via electron-updater)

## Architecture Notes

**Why Electron?**
- Modern UI with web technologies
- Cross-platform (Windows, Mac, Linux)
- Full Node.js access (can spawn PowerShell)
- Large ecosystem of UI components

**Why not pure web app?**
- Need to execute PowerShell scripts locally
- Need file system access
- Desktop app UX expected by users

**Why hybrid (Electron + PowerShell)?**
- Reuse existing PowerShell migration logic (tested, working)
- Modern UI for better UX
- Gradual migration path (can port scripts to Node.js over time)

## Version

**Current**: 2.1.14  
**Compatible with**: VGMigrations 2.1.14+

## License

Internal tool for Morite UK.

---

**Questions?** Check the main toolkit README at:
`C:\Temp\Scripts\VGMigrations\README.md`

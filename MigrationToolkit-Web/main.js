// Migration Toolkit - Electron Main Process
const { app, BrowserWindow, ipcMain, screen } = require('electron');
const path = require('path');
const { spawn } = require('child_process');

app.commandLine.appendSwitch('disable-gpu-shader-disk-cache');

// PowerShell integration — resolves correctly in both dev (source) and installed (packaged) modes
const PS_SCRIPT_PATH = app.isPackaged
  ? path.join(process.resourcesPath, 'VGMigrations')
  : path.join(__dirname, '..', 'VGMigrations');

let mainWindow;

function createWindow() {
  // Pick target display from personal config (windowMonitor is 0-based index)
  const fs = require('fs');
  const cfgPath = path.join(process.env.APPDATA, 'FlyMigration', 'config.json');
  let cfg = {};
  try { cfg = JSON.parse(fs.readFileSync(cfgPath, 'utf8')); } catch {}
  const displays = screen.getAllDisplays();
  const monIdx = (typeof cfg.windowMonitor === 'number' && cfg.windowMonitor < displays.length)
    ? cfg.windowMonitor : 0;
  const targetDisplay = displays[monIdx] || displays[0];

  mainWindow = new BrowserWindow({
    width: 1000,
    height: 720,
    minWidth: 800,
    minHeight: 600,
    x: targetDisplay.workArea.x,
    y: targetDisplay.workArea.y,
    backgroundColor: '#e8ecf3',
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js'),
      cache: false
    },
    icon: path.join(__dirname, 'public', 'icon.ico'),
    autoHideMenuBar: true,
    resizable: true,
    frame: true,
    titleBarStyle: 'default',
    show: false  // Don't show until maximized
  });

  mainWindow.loadFile('index.html');

  // Maximize onto the target display, then show
  mainWindow.maximize();
  mainWindow.show();

  // DevTools disabled for production - uncomment line below to enable for debugging
  // mainWindow.webContents.openDevTools();

  // Keyboard shortcut for DevTools (Ctrl+Shift+I)
  mainWindow.webContents.on('before-input-event', (event, input) => {
    if (input.control && input.shift && input.key.toLowerCase() === 'i') {
      mainWindow.webContents.toggleDevTools();
    }
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

// Shared: open a timestamped log file for a script run
function openScriptLog(scriptName) {
  const fs = require('fs');
  const logsDir = path.join(process.env.APPDATA, 'FlyMigration', 'Logs');
  if (!fs.existsSync(logsDir)) fs.mkdirSync(logsDir, { recursive: true });
  const ts = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
  const base = scriptName.replace(/\.ps1$/i, '');
  const logFile = fs.createWriteStream(path.join(logsDir, `${base}_${ts}.log`), { flags: 'a' });
  logFile.write(`=== ${scriptName}  ${new Date().toISOString()} ===\n`);
  return logFile;
}

// PowerShell Script Executor
function executePowerShellScript(scriptName, args = []) {
  return new Promise((resolve, reject) => {
    const scriptPath = path.join(PS_SCRIPT_PATH, scriptName);
    const psArgs = [
      '-NoProfile',
      '-ExecutionPolicy', 'Bypass',
      '-File', scriptPath,
      ...args
    ];

    const ps = spawn('pwsh.exe', psArgs, {
      cwd: PS_SCRIPT_PATH
    });

    const logFile = openScriptLog(scriptName);
    logFile.write(`Args: ${args.join(' ')}\n\n`);

    let stdout = '';
    let stderr = '';

    ps.stdout.on('data', (data) => {
      const text = data.toString();
      stdout += text;
      logFile.write(text);
    });

    ps.stderr.on('data', (data) => {
      const text = data.toString();
      stderr += text;
      logFile.write(text);
    });

    ps.on('close', (code) => {
      logFile.write(`\n=== Exit code: ${code} ===\n`);
      logFile.end();
      if (code === 0) {
        resolve({ success: true, output: stdout });
      } else {
        reject({ success: false, error: stderr || stdout, code });
      }
    });

    ps.on('error', (err) => {
      logFile.write(`\n=== Error: ${err.message} ===\n`);
      logFile.end();
      reject({ success: false, error: err.message });
    });
  });
}

// Register IPC Handlers
function registerIPCHandlers() {
  ipcMain.handle('launch-script', async (event, scriptName) => {
    try {
      // Launch PowerShell script in NEW VISIBLE window
      const scriptPath = path.join(PS_SCRIPT_PATH, scriptName);

      // Use 'start' command to open in a new window (Windows-specific)
      const child = spawn('cmd.exe', [
        '/c', 'start',
        'pwsh.exe',
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', scriptPath
      ], {
        cwd: PS_SCRIPT_PATH,
        detached: true,
        shell: true
      });

      child.unref();

      return { success: true };
    } catch (error) {
      return { success: false, error: error.message };
    }
  });

  ipcMain.handle('check-updates', async () => {
    try {
      const result = await executePowerShellScript('Check-Updates.ps1', ['-Force']);
      return result;
    } catch (error) {
      return { success: false, error: error.message };
    }
  });

  ipcMain.handle('get-version', async () => {
    try {
      const versionPath = path.join(PS_SCRIPT_PATH, 'version.json');
      const fs = require('fs');
      const versionData = JSON.parse(fs.readFileSync(versionPath, 'utf8'));
      return { success: true, version: versionData.version };
    } catch (error) {
      return { success: false, error: error.message };
    }
  });

  ipcMain.handle('get-config', async () => {
    try {
      const configPath = path.join(process.env.APPDATA, 'FlyMigration', 'config.json');
      const fs = require('fs');

      if (!fs.existsSync(configPath)) {
        return { success: true, config: null };
      }

      const configData = JSON.parse(fs.readFileSync(configPath, 'utf8').replace(/^﻿/, ''));
      return { success: true, config: configData };
    } catch (error) {
      return { success: false, error: error.message };
    }
  });

  ipcMain.handle('save-config', async (event, config) => {
    try {
      const configPath = path.join(process.env.APPDATA, 'FlyMigration', 'config.json');
      const fs = require('fs');
      const configDir = path.dirname(configPath);

      // Ensure directory exists
      if (!fs.existsSync(configDir)) {
        fs.mkdirSync(configDir, { recursive: true });
      }

      // Load existing config (strip UTF-8 BOM written by PS 5.1)
      let existingConfig = {};
      if (fs.existsSync(configPath)) {
        const rawData = fs.readFileSync(configPath, 'utf8').replace(/^﻿/, '');
        try {
          existingConfig = JSON.parse(rawData);
        } catch {
          existingConfig = {};
        }
      }

      // Merge configs
      const mergedConfig = { ...existingConfig, ...config };

      // If ClientSecret is provided (plain text), encrypt it
      if (config.ClientSecret) {
        try {
          const encryptResult = await executePowerShellScript('Encrypt-Secret.ps1', ['-Secret', config.ClientSecret]);
          if (encryptResult.success && encryptResult.output) {
            mergedConfig.EncSecret = encryptResult.output.trim();
            delete mergedConfig.ClientSecret; // Remove plain text
          }
        } catch (encError) {
          return { success: false, error: 'Failed to encrypt secret: ' + encError.message };
        }
      }

      // Save to file
      fs.writeFileSync(configPath, JSON.stringify(mergedConfig, null, 2), 'utf8');
      return { success: true };
    } catch (error) {
      return { success: false, error: error.message };
    }
  });

  ipcMain.handle('open-logs', async () => {
    try {
      const { shell } = require('electron');
      const fs = require('fs');
      const logsPath = path.join(process.env.APPDATA, 'FlyMigration', 'Logs');
      if (!fs.existsSync(logsPath)) fs.mkdirSync(logsPath, { recursive: true });
      shell.openPath(logsPath);
      return { success: true };
    } catch (error) {
      return { success: false, error: error.message };
    }
  });

  ipcMain.handle('get-migration-data', async (event, projectPrefix) => {
    try {
      const scriptPath = path.join(PS_SCRIPT_PATH, 'Get-MigrationData.ps1');
      const result = await executePowerShellScript('Get-MigrationData.ps1', ['-ProjectPrefix', projectPrefix]);

      if (result.success && result.output) {
        try {
          // Strip any non-JSON lines (Fly.Client verbose output)
          let cleanOutput = result.output.trim();
          const lines = cleanOutput.split('\n');

          // Find the first line that starts with { (JSON start)
          const jsonStartIndex = lines.findIndex(line => line.trim().startsWith('{'));
          if (jsonStartIndex >= 0) {
            cleanOutput = lines.slice(jsonStartIndex).join('\n');
          }

          const data = JSON.parse(cleanOutput);
          return { success: true, data };
        } catch (parseError) {
          return { success: false, error: 'Failed to parse migration data', output: result.output };
        }
      }

      return result;
    } catch (error) {
      return { success: false, error: error.message };
    }
  });

  ipcMain.handle('test-connection', async () => {
    try {
      const result = await executePowerShellScript('Test-FlyConnection.ps1');

      if (result.success && result.output.includes('SUCCESS')) {
        return { success: true, message: 'Connected to Fly API' };
      } else {
        // Extract error message from output
        const errorMatch = result.output.match(/ERROR: (.+)/);
        const errorMsg = errorMatch ? errorMatch[1] : (result.error || result.output || 'Unknown error');
        return { success: false, error: errorMsg };
      }
    } catch (error) {
      return { success: false, error: error.error || error.message };
    }
  });

  ipcMain.handle('execute-powershell', async (event, scriptName, args = []) => {
    try {
      const result = await executePowerShellScript(scriptName, args);
      return result;
    } catch (error) {
      return { success: false, error: error.error || error.message };
    }
  });

  // Streaming variant — sends each stdout/stderr chunk as it arrives via 'ps-output' event
  ipcMain.handle('stream-powershell', async (event, scriptName, args = []) => {
    return new Promise((resolve) => {
      const scriptPath = path.join(PS_SCRIPT_PATH, scriptName);
      const psArgs = ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', scriptPath, ...args];
      const ps = spawn('pwsh.exe', psArgs, { cwd: PS_SCRIPT_PATH });

      const logFile = openScriptLog(scriptName);
      logFile.write(`Args: ${args.join(' ')}\n\n`);

      ps.stdout.on('data', (data) => {
        const text = data.toString();
        logFile.write(text);
        if (!event.sender.isDestroyed()) event.sender.send('ps-output', text);
      });
      ps.stderr.on('data', (data) => {
        const text = data.toString();
        logFile.write(text);
        if (!event.sender.isDestroyed()) event.sender.send('ps-output', text);
      });
      ps.on('close', (code) => {
        logFile.write(`\n=== Exit code: ${code} ===\n`);
        logFile.end();
        resolve({ success: code === 0, code });
      });
      ps.on('error', (err) => {
        logFile.write(`\n=== Error: ${err.message} ===\n`);
        logFile.end();
        resolve({ success: false, error: err.message });
      });
    });
  });

  ipcMain.handle('get-shared-config', async () => {
    try {
      const fs = require('fs');
      const cfgPath = path.join(process.env.LOCALAPPDATA, 'FlyMigration', 'shared-config.json');
      if (!fs.existsSync(cfgPath)) return { success: true, config: {} };
      const cfg = JSON.parse(fs.readFileSync(cfgPath, 'utf8').replace(/^﻿/, ''));
      return { success: true, config: cfg };
    } catch (error) {
      return { success: false, error: error.message };
    }
  });

  ipcMain.handle('save-shared-config', async (event, values) => {
    try {
      const fs = require('fs');
      const cfgPath = path.join(process.env.LOCALAPPDATA, 'FlyMigration', 'shared-config.json');
      const cfgDir  = path.dirname(cfgPath);
      if (!fs.existsSync(cfgDir)) fs.mkdirSync(cfgDir, { recursive: true });
      let cfg = {};
      if (fs.existsSync(cfgPath)) {
        try { cfg = JSON.parse(fs.readFileSync(cfgPath, 'utf8').replace(/^﻿/, '')); } catch {}
      }
      Object.assign(cfg, values);
      fs.writeFileSync(cfgPath, JSON.stringify(cfg, null, 2), 'utf8');
      return { success: true };
    } catch (error) {
      return { success: false, error: error.message };
    }
  });

  ipcMain.handle('show-open-dialog', async (event, options) => {
    const { dialog } = require('electron');
    return dialog.showOpenDialog(mainWindow, options || {
      properties: ['openFile'],
      filters: [
        { name: 'CSV Files', extensions: ['csv'] },
        { name: 'All Files', extensions: ['*'] }
      ]
    });
  });

  ipcMain.handle('show-save-dialog', async (event, options) => {
    const { dialog } = require('electron');
    return dialog.showSaveDialog(mainWindow, options || {});
  });

  ipcMain.handle('open-external', async (event, url) => {
    try {
      const { shell } = require('electron');
      await shell.openExternal(url);
      return { success: true };
    } catch (error) {
      return { success: false, error: error.message };
    }
  });

  ipcMain.handle('read-vbu-csv', async (event, filePath) => {
    try {
      const fs = require('fs');
      if (!filePath || !fs.existsSync(filePath)) {
        return { success: false, error: 'File not found', rows: [] };
      }
      const content = fs.readFileSync(filePath, 'utf8');
      const lines = content.split(/\r?\n/).filter(l => l.trim());
      if (lines.length < 2) return { success: true, rows: [] };

      const header = lines[0].split(',').map(h => h.trim().toLowerCase());
      const domainIdx  = header.indexOf('email domain');
      const vbuIdIdx   = header.indexOf('vbu id');
      const vbuNameIdx = header.indexOf('vbu name');

      if (domainIdx === -1 || vbuIdIdx === -1) {
        return { success: false, error: 'Columns "Email Domain" and "VBU ID" not found in CSV', rows: [] };
      }

      const rows = lines.slice(1)
        .map(line => {
          const cols = line.split(',');
          return {
            domain:  (cols[domainIdx]  || '').trim().toLowerCase(),
            vbuId:   (cols[vbuIdIdx]   || '').trim(),
            vbuName: vbuNameIdx >= 0 ? (cols[vbuNameIdx] || '').trim() : ''
          };
        })
        .filter(r => r.domain && r.domain.includes('.'))
        .sort((a, b) => a.domain.localeCompare(b.domain));

      return { success: true, rows };
    } catch (error) {
      return { success: false, error: error.message, rows: [] };
    }
  });
}

// App lifecycle
app.whenReady().then(() => {
  registerIPCHandlers();
  createWindow();
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});

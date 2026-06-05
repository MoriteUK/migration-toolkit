// Migration Toolkit - Electron Main Process
const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const { spawn } = require('child_process');

// PowerShell integration
const PS_SCRIPT_PATH = path.join(__dirname, '..', 'VGMigrations');

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1000,
    height: 720,
    minWidth: 800,
    minHeight: 600,
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

  // Maximize window on startup, but keep resizable
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

    let stdout = '';
    let stderr = '';

    ps.stdout.on('data', (data) => {
      stdout += data.toString();
    });

    ps.stderr.on('data', (data) => {
      stderr += data.toString();
    });

    ps.on('close', (code) => {
      if (code === 0) {
        resolve({ success: true, output: stdout });
      } else {
        reject({ success: false, error: stderr || stdout, code });
      }
    });

    ps.on('error', (err) => {
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

      const configData = JSON.parse(fs.readFileSync(configPath, 'utf8'));
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

      // Load existing config
      let existingConfig = {};
      if (fs.existsSync(configPath)) {
        const rawData = fs.readFileSync(configPath, 'utf8');
        existingConfig = JSON.parse(rawData);
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
      const logsPath = path.join(PS_SCRIPT_PATH, 'logs');
      const fs = require('fs');

      // Create logs directory if it doesn't exist
      if (!fs.existsSync(logsPath)) {
        fs.mkdirSync(logsPath, { recursive: true });
      }

      // Open the logs folder in Windows Explorer
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

  ipcMain.handle('open-external', async (event, url) => {
    try {
      const { shell } = require('electron');
      await shell.openExternal(url);
      return { success: true };
    } catch (error) {
      return { success: false, error: error.message };
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

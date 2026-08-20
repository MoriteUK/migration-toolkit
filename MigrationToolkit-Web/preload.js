// Migration Toolkit - Preload Script (Bridge between Electron and Renderer)
const { contextBridge, ipcRenderer } = require('electron');

// Expose protected methods to renderer process
contextBridge.exposeInMainWorld('electronAPI', {
  // Launch PowerShell scripts
  launchScript: (scriptName) => ipcRenderer.invoke('launch-script', scriptName),

  // Check for updates
  checkUpdates: () => ipcRenderer.invoke('check-updates'),

  // Get current version
  getVersion: () => ipcRenderer.invoke('get-version'),

  // Config management
  getConfig: () => ipcRenderer.invoke('get-config'),
  saveConfig: (config) => ipcRenderer.invoke('save-config', config),

  // Open logs folder
  openLogs: () => ipcRenderer.invoke('open-logs'),

  // Get real migration data
  getMigrationData: (projectPrefix) => ipcRenderer.invoke('get-migration-data', projectPrefix),

  // Test Fly API connection
  testConnection: () => ipcRenderer.invoke('test-connection'),

  // Open external URL
  openExternal: (url) => ipcRenderer.invoke('open-external', url),

  // Execute PowerShell script with arguments
  executePowerShell: (scriptName, args) => ipcRenderer.invoke('execute-powershell', scriptName, args),

  // Stream PowerShell script output
  streamPowerShell: (scriptName, args) => ipcRenderer.invoke('stream-powershell', scriptName, args),

  // Stop the currently running streamed PowerShell script, if any
  stopPowerShell: () => ipcRenderer.invoke('stop-powershell'),

  // Listen for PowerShell output events
  onPsOutput: (callback) => ipcRenderer.on('ps-output', (event, text) => callback(text)),
  offPsOutput: () => ipcRenderer.removeAllListeners('ps-output'),

  // Open file with default application
  openFile: (filePath) => ipcRenderer.invoke('open-file', filePath),

  // Show native open dialog for file/folder selection
  showOpenDialog: (options) => ipcRenderer.invoke('show-open-dialog', options),

  // Read VBU CSV file for domain dropdown
  readVbuCsv: (filePath) => ipcRenderer.invoke('read-vbu-csv', filePath)
});

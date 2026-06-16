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

  // Native file-open dialog
  showOpenDialog: (options) => ipcRenderer.invoke('show-open-dialog', options),

  // Execute PowerShell script with arguments (buffered — returns when done)
  executePowerShell: (scriptName, args) => ipcRenderer.invoke('execute-powershell', scriptName, args),

  // Streaming variant — resolves when done; output arrives via onPsOutput callbacks
  streamPowerShell: (scriptName, args) => ipcRenderer.invoke('stream-powershell', scriptName, args),
  onPsOutput:  (cb) => ipcRenderer.on('ps-output', (_e, data) => cb(data)),
  offPsOutput: ()   => ipcRenderer.removeAllListeners('ps-output'),

  // Shared config (AOS tenant details)
  getSharedConfig: () => ipcRenderer.invoke('get-shared-config'),
  saveSharedConfig: (values) => ipcRenderer.invoke('save-shared-config', values),

  // VBU CSV reader — parses domain/VBU ID mapping from a CSV file
  readVbuCsv: (filePath) => ipcRenderer.invoke('read-vbu-csv', filePath)
});

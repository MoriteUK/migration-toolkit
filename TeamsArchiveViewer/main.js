// Teams Archive Viewer - standalone Electron shell around viewer.html.
// Deliberately minimal: one window, no menu bar, no devtools, no network access
// needed - everything the page does happens locally against files the user picks.

const { app, BrowserWindow, Menu } = require('electron');
const path = require('path');

// No File/Edit/View/Window/Help menu bar - this is a single-purpose viewer, not
// a browser, and that menu bar is one of the visual cues that reads as "generic
// Chromium app" rather than a purpose-built tool.
Menu.setApplicationMenu(null);

function createWindow() {
  const win = new BrowserWindow({
    width: 1280,
    height: 860,
    minWidth: 760,
    minHeight: 480,
    title: 'Archived Teams Chats',
    backgroundColor: '#f5f5f5',
    autoHideMenuBar: true,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true
    }
  });

  win.loadFile(path.join(__dirname, 'viewer.html'));
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});

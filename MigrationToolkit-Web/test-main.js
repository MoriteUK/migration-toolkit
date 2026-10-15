// Test Electron
const { app, BrowserWindow } = require('electron');

console.log('Electron app loaded:', typeof app);

app.whenReady().then(() => {
  console.log('App ready!');
  const win = new BrowserWindow({ width: 800, height: 600 });
  win.loadFile('index.html');
});

app.on('window-all-closed', () => {
  app.quit();
});

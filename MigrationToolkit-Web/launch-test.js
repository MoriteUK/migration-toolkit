// Throwaway diagnostic — run from MigrationToolkit-Web/ with:  npx electron launch-test.js
// It reproduces the app's console-less main-process condition and fires several
// "open a PowerShell window" strategies at once. Watch your screen: note which
// numbered windows appear (each prints "STRATEGY N — <desc>" and stays open).
// Also writes results to %TEMP%\mtk-launch-test.log. Close the windows when done.
const { app } = require('electron');
const { spawn } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

const logPath = path.join(os.tmpdir(), 'mtk-launch-test.log');
const log = (m) => { fs.appendFileSync(logPath, m + '\n'); console.log(m); };

function tryStrategy(n, desc, file, args, opts) {
  try {
    const c = spawn(file, args, opts);
    c.on('error', (e) => log(`  S${n} spawn error: ${e.message}`));
    c.on('exit', (code) => log(`  S${n} launcher exited: ${code}`));
    log(`S${n} [${desc}] spawned pid=${c.pid}  ${file} ${args.join(' ')}  opts=${JSON.stringify(opts)}`);
    if (opts.detached) c.unref();
  } catch (e) {
    log(`S${n} [${desc}] THREW: ${e.message}`);
  }
}

app.whenReady().then(() => {
  fs.writeFileSync(logPath, `mtk launch test  ${new Date().toISOString()}\n`);
  const hold = (n, d) => ['-NoProfile', '-NoExit', '-Command',
    `$Host.UI.RawUI.WindowTitle='STRATEGY ${n}'; Write-Host 'STRATEGY ${n} - ${d}' -ForegroundColor Green`];

  tryStrategy(1, 'cmd /c start, shell:true, detached',
    'cmd.exe', ['/c', 'start', '""', 'pwsh.exe', ...hold(1, 'cmd start shell:true detached')],
    { detached: true, shell: true, windowsHide: false });

  tryStrategy(2, 'cmd /c start, shell:false, detached',
    'cmd.exe', ['/c', 'start', '""', 'pwsh.exe', ...hold(2, 'cmd start shell:false detached')],
    { detached: true, shell: false, windowsHide: false });

  tryStrategy(3, 'cmd /c start, shell:true, NOT detached',
    'cmd.exe', ['/c', 'start', '""', 'pwsh.exe', ...hold(3, 'cmd start shell:true no-detach')],
    { shell: true, windowsHide: false });

  tryStrategy(4, 'pwsh direct, NOT detached, stdio ignore',
    'pwsh.exe', hold(4, 'pwsh direct no-detach stdio-ignore'),
    { stdio: 'ignore', windowsHide: false });

  tryStrategy(5, 'pwsh direct, NOT detached, default stdio',
    'pwsh.exe', hold(5, 'pwsh direct no-detach default-stdio'),
    { windowsHide: false });

  tryStrategy(6, 'conhost pwsh, NOT detached',
    'conhost.exe', ['pwsh.exe', ...hold(6, 'conhost no-detach')],
    { windowsHide: false });

  tryStrategy(7, 'pwsh direct, detached, stdio ignore',
    'pwsh.exe', hold(7, 'pwsh direct detached stdio-ignore'),
    { detached: true, stdio: 'ignore', windowsHide: false });

  log(`\nResults log: ${logPath}`);
  log('Leaving the app up 20s so windows can appear, then quitting.');
  setTimeout(() => app.quit(), 20000);
});

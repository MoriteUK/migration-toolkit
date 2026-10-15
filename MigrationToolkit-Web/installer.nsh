; Custom NSIS install hook — runs after all files have been copied
; Opens a visible PowerShell window so the user can see dependency installation progress
!macro customInstall
  DetailPrint "Installing dependencies (Node.js, Playwright, PowerShell modules)..."
  DetailPrint "A setup window will open — please wait for it to complete before using the toolkit."
  ExecWait 'cmd.exe /c start "Migration Toolkit — First-Time Setup" /wait pwsh.exe -NoProfile -ExecutionPolicy Bypass -File "$INSTDIR\resources\VGMigrations\Setup.ps1"' $0
  ${If} $0 != 0
    MessageBox MB_ICONEXCLAMATION "Setup completed with warnings (exit code $0).$\n$\nThe toolkit may still work. If you have issues, run Setup.ps1 manually from:$\n$INSTDIR\resources\VGMigrations\Setup.ps1"
  ${EndIf}
!macroend

!macro customUnInstall
  ; Nothing extra needed on uninstall
!macroend

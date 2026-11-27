# Quick verification that Reports section is in place

Write-Host "Verifying Reports Section Integration..." -ForegroundColor Cyan
Write-Host ""

$results = @()

# Check index.html has Reports section
$indexPath = Join-Path $PSScriptRoot "index.html"
$indexContent = Get-Content $indexPath -Raw

if ($indexContent -match 'data-menu="reports"') {
    $results += [PSCustomObject]@{ Check = "Sidebar section"; Status = "✅ Found"; File = "index.html" }
} else {
    $results += [PSCustomObject]@{ Check = "Sidebar section"; Status = "❌ Missing"; File = "index.html" }
}

if ($indexContent -match 'id="flyBriefView"') {
    $results += [PSCustomObject]@{ Check = "View container"; Status = "✅ Found"; File = "index.html" }
} else {
    $results += [PSCustomObject]@{ Check = "View container"; Status = "❌ Missing"; File = "index.html" }
}

# Check renderer.js has view mapping
$rendererPath = Join-Path $PSScriptRoot "renderer.js"
$rendererContent = Get-Content $rendererPath -Raw

if ($rendererContent -match "'reports-fly-brief'") {
    $results += [PSCustomObject]@{ Check = "View mapping"; Status = "✅ Found"; File = "renderer.js" }
} else {
    $results += [PSCustomObject]@{ Check = "View mapping"; Status = "❌ Missing"; File = "renderer.js" }
}

if ($rendererContent -match 'generateFlyBriefFromApi') {
    $results += [PSCustomObject]@{ Check = "Event handlers"; Status = "✅ Found"; File = "renderer.js" }
} else {
    $results += [PSCustomObject]@{ Check = "Event handlers"; Status = "❌ Missing"; File = "renderer.js" }
}

# Check main.js has backend handler
$mainPath = Join-Path $PSScriptRoot "main.js"
$mainContent = Get-Content $mainPath -Raw

if ($mainContent -match 'generate-fly-brief-from-api') {
    $results += [PSCustomObject]@{ Check = "Backend handler"; Status = "✅ Found"; File = "main.js" }
} else {
    $results += [PSCustomObject]@{ Check = "Backend handler"; Status = "❌ Missing"; File = "main.js" }
}

# Check preload.js exposes API
$preloadPath = Join-Path $PSScriptRoot "preload.js"
$preloadContent = Get-Content $preloadPath -Raw

if ($preloadContent -match 'generateFlyBriefFromApi') {
    $results += [PSCustomObject]@{ Check = "IPC bridge"; Status = "✅ Found"; File = "preload.js" }
} else {
    $results += [PSCustomObject]@{ Check = "IPC bridge"; Status = "❌ Missing"; File = "preload.js" }
}

# Check FLY template files exist
$flyTemplatePath = Join-Path $PSScriptRoot "..\FLY\Template"
$invokeFlyScript = Join-Path $flyTemplatePath "Invoke-FlyBriefs.ps1"

if (Test-Path $invokeFlyScript) {
    $results += [PSCustomObject]@{ Check = "FLY scripts"; Status = "✅ Found"; File = "FLY\Template\" }
} else {
    $results += [PSCustomObject]@{ Check = "FLY scripts"; Status = "❌ Missing"; File = "FLY\Template\" }
}

# Display results
$results | Format-Table -AutoSize

Write-Host ""
$missing = ($results | Where-Object { $_.Status -like "*Missing*" }).Count
$found = ($results | Where-Object { $_.Status -like "*Found*" }).Count

if ($missing -eq 0) {
    Write-Host "✅ All components verified! Reports section should appear after restart." -ForegroundColor Green
    Write-Host ""
    Write-Host "To restart the toolkit:" -ForegroundColor Cyan
    Write-Host "  1. Close current instance" -ForegroundColor Gray
    Write-Host "  2. Run: .\START.cmd" -ForegroundColor Gray
    Write-Host ""
} else {
    Write-Host "⚠️  $missing component(s) missing. Reports section may not work." -ForegroundColor Yellow
    Write-Host ""
}

# Check if toolkit is running
$electronProcess = Get-Process -Name "electron" -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -like "*Migration*" }
if ($electronProcess) {
    Write-Host "⚠️  Migration Toolkit is currently running (PID: $($electronProcess.Id))" -ForegroundColor Yellow
    Write-Host "    Changes require a restart to take effect!" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "    To restart now, run:" -ForegroundColor Cyan
    Write-Host "      Stop-Process -Id $($electronProcess.Id)" -ForegroundColor Gray
    Write-Host "      .\START.cmd" -ForegroundColor Gray
} else {
    Write-Host "ℹ️  Toolkit not currently running. Start with: .\START.cmd" -ForegroundColor Cyan
}

Write-Host ""

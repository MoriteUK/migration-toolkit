#Requires -Version 7.0
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Schedule-CacheRefresh.ps1 - Set up automatic Monday 3am cache refresh

.DESCRIPTION
    Creates Windows scheduled tasks to refresh all discovery caches every Monday at 3am.
    Tasks run as the current user and use their saved credentials.

.PARAMETER SharePointAdminUrl
    SPO admin URL (e.g., https://ourvolaris-admin.sharepoint.com)

.PARAMETER Uninstall
    Remove all scheduled cache refresh tasks

.EXAMPLE
    .\Schedule-CacheRefresh.ps1 -SharePointAdminUrl https://ourvolaris-admin.sharepoint.com

.EXAMPLE
    .\Schedule-CacheRefresh.ps1 -Uninstall
#>

param(
    [string]$SharePointAdminUrl = 'https://ourvolaris-admin.sharepoint.com',
    [switch]$Uninstall
)

$ErrorActionPreference = 'Stop'

# Task names with staggered start times
# Incremental updates complete in ~3.5 hours (first run: 24-31 hours)
# Incremental timings: SharePoint ~10min, Teams ~30min, DGs ~10min, Mailboxes ~2hrs
$tasks = @{
    'SharePoint Sites Cache'     = @{ Script = 'Update-SharePointSitesCache.ps1';     Time = '5:00AM'; Day = 'Saturday' }
    'Teams Channels Cache'       = @{ Script = 'Update-TeamsChannelsCache.ps1';       Time = '6:00AM'; Day = 'Saturday' }
    'Distribution Groups Cache'  = @{ Script = 'Update-DistributionGroupsCache.ps1';  Time = '11:00AM'; Day = 'Saturday' }
    'Exchange Mailboxes Cache'   = @{ Script = 'Update-MailboxesCache.ps1';           Time = '12:00PM'; Day = 'Saturday' }
}

$taskFolder = '\FlyMigration'

if ($Uninstall) {
    Write-Host "`nRemoving scheduled cache refresh tasks..." -ForegroundColor Cyan

    foreach ($taskName in $tasks.Keys) {
        $fullTaskName = "$taskFolder\$taskName"
        try {
            Unregister-ScheduledTask -TaskName $fullTaskName -Confirm:$false -ErrorAction SilentlyContinue
            Write-Host "  ✓ Removed: $taskName" -ForegroundColor Green
        }
        catch {
            Write-Host "  - Not found: $taskName" -ForegroundColor DarkGray
        }
    }

    Write-Host "`nAll cache refresh tasks removed.`n" -ForegroundColor Green
    return
}

# Validate parameters
if (-not $SharePointAdminUrl) {
    throw "SharePointAdminUrl is required. Example: .\Schedule-CacheRefresh.ps1 -SharePointAdminUrl https://yourtenantname-admin.sharepoint.com"
}

Write-Host "`n╔═══════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║  Schedule Automatic Cache Refresh - Staggered Start      ║" -ForegroundColor Cyan
Write-Host "╚═══════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""
Write-Host "SharePoint Admin URL: $SharePointAdminUrl" -ForegroundColor White
Write-Host "Schedule:             Every Saturday (incremental updates)" -ForegroundColor White
Write-Host "  • SharePoint Sites:     Saturday  5:00 AM  (est. finish ~5:10 AM)" -ForegroundColor DarkGray
Write-Host "  • Teams Channels:       Saturday  6:00 AM  (est. finish ~6:30 AM)" -ForegroundColor DarkGray
Write-Host "  • Distribution Groups:  Saturday 11:00 AM  (est. finish ~11:10 AM)" -ForegroundColor DarkGray
Write-Host "  • Exchange Mailboxes:   Saturday 12:00 PM  (est. finish ~2:00 PM)" -ForegroundColor DarkGray
Write-Host "Run as:               $env:USERDOMAIN\$env:USERNAME" -ForegroundColor White
Write-Host ""

# Create task folder if it doesn't exist
try {
    $null = Get-ScheduledTask -TaskPath $taskFolder -ErrorAction Stop
}
catch {
    Write-Host "Creating task folder: $taskFolder" -ForegroundColor DarkGray
    $scheduleService = New-Object -ComObject Schedule.Service
    $scheduleService.Connect()
    $rootFolder = $scheduleService.GetFolder("\")
    $null = $rootFolder.CreateFolder($taskFolder.TrimStart('\'))
}

# Create scheduled tasks
foreach ($taskName in $tasks.Keys) {
    $taskConfig = $tasks[$taskName]
    $scriptName = $taskConfig.Script
    $startTime  = $taskConfig.Time
    $scriptPath = Join-Path $PSScriptRoot $scriptName

    if (-not (Test-Path $scriptPath)) {
        Write-Warning "Script not found: $scriptPath - skipping"
        continue
    }

    Write-Host "Creating task: $taskName ($startTime)..." -ForegroundColor Cyan

    # Task action - run PowerShell 7 with the cache refresh script
    $action = New-ScheduledTaskAction `
        -Execute 'pwsh.exe' `
        -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -SharePointAdminUrl `"$SharePointAdminUrl`"" `
        -WorkingDirectory $PSScriptRoot

    # Trigger - Every Saturday at the specified time
    $dayOfWeek = if ($taskConfig.Day) { $taskConfig.Day } else { 'Saturday' }
    $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $dayOfWeek -At $startTime

    # Settings
    $settings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -StartWhenAvailable `
        -RunOnlyIfNetworkAvailable `
        -ExecutionTimeLimit (New-TimeSpan -Hours 4)

    # Principal - run as current user (uses their saved credentials)
    $principal = New-ScheduledTaskPrincipal `
        -UserId "$env:USERDOMAIN\$env:USERNAME" `
        -LogonType S4U `
        -RunLevel Highest

    # Register the task
    $fullTaskName = "$taskFolder\$taskName"
    try {
        $null = Register-ScheduledTask `
            -TaskName $fullTaskName `
            -Action $action `
            -Trigger $trigger `
            -Settings $settings `
            -Principal $principal `
            -Description "Automatic cache refresh for Migration Toolkit - runs every Saturday (48hr window)" `
            -Force

        Write-Host "  ✓ Created successfully" -ForegroundColor Green
    }
    catch {
        Write-Warning "  Failed to create task: $($_.Exception.Message)"
    }
}

Write-Host ""
Write-Host "╔═══════════════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║  Scheduled Tasks Created Successfully                     ║" -ForegroundColor Green
Write-Host "╚═══════════════════════════════════════════════════════════╝" -ForegroundColor Green
Write-Host ""
Write-Host "Tasks run every Saturday (staggered) as $env:USERDOMAIN\$env:USERNAME" -ForegroundColor White
Write-Host "  • SharePoint Sites:     Saturday  5:00 AM  → finish ~5:10 AM" -ForegroundColor Green
Write-Host "  • Teams Channels:       Saturday  6:00 AM  → finish ~6:30 AM" -ForegroundColor Green
Write-Host "  • Distribution Groups:  Saturday 11:00 AM  → finish ~11:10 AM" -ForegroundColor Green
Write-Host "  • Exchange Mailboxes:   Saturday 12:00 PM  → finish ~2:00 PM" -ForegroundColor Green
Write-Host ""
Write-Host "All caches refreshed by ~2:00 PM Saturday (incremental updates)!" -ForegroundColor Yellow
Write-Host ""
Write-Host "Incremental cache benefits:" -ForegroundColor Yellow
Write-Host "  - Only updates NEW/MODIFIED/DELETED items (90% faster)" -ForegroundColor DarkGray
Write-Host "  - Typical weekly refresh: ~3.5 hours (down from 24-31 hours)" -ForegroundColor DarkGray
Write-Host "  - First run still takes 24-31 hours to build initial cache" -ForegroundColor DarkGray
Write-Host "  - Fresh caches ready for Monday morning assessments" -ForegroundColor DarkGray
Write-Host "  - Staggered timing prevents overlapping tasks and API spikes" -ForegroundColor DarkGray
Write-Host ""
Write-Host "To view tasks:" -ForegroundColor Yellow
Write-Host "  1. Open Task Scheduler (taskschd.msc)" -ForegroundColor Gray
Write-Host "  2. Navigate to: Task Scheduler Library > FlyMigration" -ForegroundColor Gray
Write-Host ""
Write-Host "To test now:" -ForegroundColor Yellow
Write-Host "  Get-ScheduledTask -TaskPath '$taskFolder\*' | Start-ScheduledTask" -ForegroundColor Gray
Write-Host ""
Write-Host "To remove tasks:" -ForegroundColor Yellow
Write-Host "  .\Schedule-CacheRefresh.ps1 -Uninstall" -ForegroundColor Gray
Write-Host ""

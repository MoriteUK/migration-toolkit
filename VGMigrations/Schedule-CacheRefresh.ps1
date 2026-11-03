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
# Staggered to complete all caches by 8am Monday morning
# Based on typical durations: SharePoint ~1hr, Teams ~5hrs, DGs ~1hr
$tasks = @{
    'SharePoint Sites Cache'     = @{ Script = 'Update-SharePointSitesCache.ps1';     Time = '12:30AM' }
    'Teams Channels Cache'       = @{ Script = 'Update-TeamsChannelsCache.ps1';       Time = '1:30AM' }
    'Distribution Groups Cache'  = @{ Script = 'Update-DistributionGroupsCache.ps1';  Time = '6:30AM' }
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
Write-Host "Schedule:             Every Monday (staggered - finish by 8am)" -ForegroundColor White
Write-Host "  • SharePoint Sites:     12:30 AM  (est. finish ~1:30 AM)" -ForegroundColor DarkGray
Write-Host "  • Teams Channels:        1:30 AM  (est. finish ~6:30 AM)" -ForegroundColor DarkGray
Write-Host "  • Distribution Groups:   6:30 AM  (est. finish ~7:30 AM)" -ForegroundColor DarkGray
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

    # Trigger - Every Monday at the specified time
    $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At $startTime

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
            -Description "Automatic cache refresh for Migration Toolkit - runs every Monday at 3am" `
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
Write-Host "Tasks run every Monday (staggered) as $env:USERDOMAIN\$env:USERNAME" -ForegroundColor White
Write-Host "  • SharePoint Sites:     12:30 AM  → finish ~1:30 AM" -ForegroundColor Green
Write-Host "  • Teams Channels:        1:30 AM  → finish ~6:30 AM" -ForegroundColor Green
Write-Host "  • Distribution Groups:   6:30 AM  → finish ~7:30 AM" -ForegroundColor Green
Write-Host ""
Write-Host "All caches refreshed by 8:00 AM Monday morning!" -ForegroundColor Yellow
Write-Host ""
Write-Host "Staggered timing prevents:" -ForegroundColor Yellow
Write-Host "  - Overlapping long-running tasks" -ForegroundColor DarkGray
Write-Host "  - API throttling from parallel requests" -ForegroundColor DarkGray
Write-Host "  - High CPU/memory load spikes" -ForegroundColor DarkGray
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

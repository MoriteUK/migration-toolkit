#Requires -Version 7.0
<#
.SYNOPSIS
    Archives old log files to keep logs folder clean
.DESCRIPTION
    Moves log files older than specified days to an 'old' subfolder
.PARAMETER DaysOld
    Number of days - logs older than this will be archived (default: 7)
.PARAMETER LogPath
    Path to logs folder (default: script directory/logs)
#>
param(
    [int]$DaysOld = 7,
    [string]$LogPath = (Join-Path $PSScriptRoot "logs")
)

try {
    # Ensure logs folder exists
    if (-not (Test-Path $LogPath)) {
        Write-Warning "Logs folder not found: $LogPath"
        exit 0
    }

    # Create 'old' subfolder if it doesn't exist
    $oldFolder = Join-Path $LogPath "old"
    if (-not (Test-Path $oldFolder)) {
        New-Item -ItemType Directory -Path $oldFolder -Force | Out-Null
        Write-Host "Created archive folder: $oldFolder" -ForegroundColor Green
    }

    # Calculate cutoff date
    $cutoffDate = (Get-Date).AddDays(-$DaysOld)
    Write-Host "Archiving logs older than: $($cutoffDate.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Cyan

    # Get all log files (excluding the old folder)
    $logFiles = Get-ChildItem -Path $LogPath -File -Recurse:$false | Where-Object {
        $_.LastWriteTime -lt $cutoffDate
    }

    if ($logFiles.Count -eq 0) {
        Write-Host "No old logs to archive." -ForegroundColor Yellow
        exit 0
    }

    # Move old logs to archive folder
    $movedCount = 0
    foreach ($logFile in $logFiles) {
        try {
            $destination = Join-Path $oldFolder $logFile.Name

            # If file already exists in old folder, append timestamp to avoid conflicts
            if (Test-Path $destination) {
                $timestamp = $logFile.LastWriteTime.ToString('yyyyMMdd-HHmmss')
                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($logFile.Name)
                $extension = [System.IO.Path]::GetExtension($logFile.Name)
                $destination = Join-Path $oldFolder "$baseName-$timestamp$extension"
            }

            Move-Item -Path $logFile.FullName -Destination $destination -Force
            Write-Host "  Archived: $($logFile.Name) -> old/$([System.IO.Path]::GetFileName($destination))" -ForegroundColor Gray
            $movedCount++
        } catch {
            Write-Warning "Failed to archive $($logFile.Name): $_"
        }
    }

    Write-Host "`nArchived $movedCount log file(s)." -ForegroundColor Green

    # Optional: Clean up very old files in archive folder (older than 90 days)
    $veryOldCutoff = (Get-Date).AddDays(-90)
    $veryOldFiles = Get-ChildItem -Path $oldFolder -File | Where-Object {
        $_.LastWriteTime -lt $veryOldCutoff
    }

    if ($veryOldFiles.Count -gt 0) {
        Write-Host "`nCleaning up logs older than 90 days..." -ForegroundColor Cyan
        foreach ($file in $veryOldFiles) {
            Remove-Item $file.FullName -Force
            Write-Host "  Deleted: $($file.Name)" -ForegroundColor DarkGray
        }
        Write-Host "Deleted $($veryOldFiles.Count) very old log file(s)." -ForegroundColor Green
    }

} catch {
    Write-Error "Log archival failed: $_"
    exit 1
}

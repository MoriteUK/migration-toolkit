#Requires -Version 7.0
<#
.SYNOPSIS
    Archives log files older than 7 days to Old_Logs.zip
.DESCRIPTION
    Runs automatically on toolkit startup. Moves log files older than 7 days
    into Old_Logs.zip to keep the logs folder clean while preserving history.
.PARAMETER DaysOld
    Number of days - logs older than this will be archived (default: 7)
.PARAMETER LogPath
    Path to logs folder (default: script directory/logs)
#>
param(
    [int]$DaysOld = 7,
    [string]$LogPath = (Join-Path $PSScriptRoot "logs")
)

$ErrorActionPreference = 'SilentlyContinue'

try {
    # Ensure logs folder exists
    if (-not (Test-Path $LogPath)) {
        exit 0
    }

    # Calculate cutoff date
    $cutoffDate = (Get-Date).AddDays(-$DaysOld)
    Write-Host "Archiving logs older than: $($cutoffDate.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Cyan

    # Get cutoff date
    $cutoffDate = (Get-Date).AddDays(-$DaysOld)

    # Get log files older than specified days
    $oldLogs = Get-ChildItem -Path $LogPath -File -Filter "*.log" | Where-Object {
        $_.LastWriteTime -lt $cutoffDate
    }

    if ($oldLogs.Count -eq 0) {
        # No old logs to archive - exit silently
        exit 0
    }

    # Archive path
    $archivePath = Join-Path $LogPath 'Old_Logs.zip'

    # Load compression assembly
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    # Open or create the zip file
    if (Test-Path $archivePath) {
        $zip = [System.IO.Compression.ZipFile]::Open($archivePath, 'Update')
    } else {
        $zip = [System.IO.Compression.ZipFile]::Open($archivePath, 'Create')
    }

    $archived = 0
    foreach ($log in $oldLogs) {
        try {
            # Check if file already exists in zip
            $entryName = $log.Name
            $existingEntry = $zip.Entries | Where-Object { $_.Name -eq $entryName }

            if (-not $existingEntry) {
                # Add file to zip
                [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
                    $zip,
                    $log.FullName,
                    $entryName,
                    [System.IO.Compression.CompressionLevel]::Optimal
                ) | Out-Null

                # Delete original file after successful archive
                Remove-Item $log.FullName -Force
                $archived++
            } else {
                # File already archived, just delete it
                Remove-Item $log.FullName -Force
                $archived++
            }
        } catch {
            # Skip this file if there's an error
            continue
        }
    }

    $zip.Dispose()

    # Only show message if logs were archived (don't clutter startup)
    # Silent operation for better UX

} catch {
    # Silently fail - don't interrupt toolkit startup
    if ($zip) { $zip.Dispose() }
}

exit 0

# Log Management

## Overview

The Migration Toolkit automatically manages log files to keep the logs folder clean and organized. Old logs are automatically moved to an archive folder.

## Automatic Log Archival

### How It Works

When the toolkit starts (menu.ps1), it automatically:

1. **Identifies old logs**: Files older than 7 days in the `logs/` folder
2. **Creates archive folder**: Creates `logs/old/` if it doesn't exist
3. **Moves old logs**: Transfers old logs to the archive folder
4. **Prevents conflicts**: If a log file with the same name exists in archive, it appends a timestamp
5. **Cleans very old logs**: Automatically deletes logs older than 90 days from the archive

### Log Locations

```
VGMigrations/
├── logs/                       # Current logs (last 7 days)
│   ├── discovery-2024-06-04.log
│   ├── migration-2024-06-03.log
│   └── old/                    # Archived logs
│       ├── discovery-2024-05-28.log
│       ├── migration-2024-05-25.log
│       └── ...
```

## Manual Log Archival

### Using the Archive Script

You can manually archive logs at any time:

```powershell
# Archive logs older than 7 days (default)
.\Archive-OldLogs.ps1

# Archive logs older than 30 days
.\Archive-OldLogs.ps1 -DaysOld 30

# Specify custom log path
.\Archive-OldLogs.ps1 -LogPath "C:\CustomLogs" -DaysOld 14
```

### Using the Helper Function

From within PowerShell scripts that load lib.ps1:

```powershell
# Archive logs older than 7 days
Move-OldLogs

# Archive logs older than 14 days
Move-OldLogs -DaysOld 14
```

## Log Retention Policy

| Location | Retention Period | Action |
|----------|-----------------|--------|
| `logs/` (main folder) | 7 days | Moved to `old/` subfolder |
| `logs/old/` (archive) | 90 days | Automatically deleted |

## Configuration

To change the retention periods, edit:

1. **Automatic archival**: Edit `menu.ps1` line:
   ```powershell
   Move-OldLogs -DaysOld 7  # Change 7 to desired days
   ```

2. **Very old cleanup**: Edit `lib.ps1` in the `Move-OldLogs` function:
   ```powershell
   (Get-Date).AddDays(-90)  # Change -90 to desired days
   ```

## Viewing Logs

### Current Logs
Current logs (last 7 days) are in:
```
VGMigrations\logs\
```

### Archived Logs
Archived logs are in:
```
VGMigrations\logs\old\
```

### Opening Logs Folder

From the Migration Toolkit app:
1. Click Settings (⚙)
2. Look for "Open Logs" button
3. Or navigate manually to the logs folder

## Log File Naming

Log files typically follow these naming patterns:

- `discovery-YYYY-MM-DD.log` - Discovery script logs
- `migration-YYYY-MM-DD.log` - Migration script logs
- `runner-YYYY-MM-DD.log` - Runner script logs
- `monitor-YYYY-MM-DD.log` - Monitor script logs

When archived with timestamp conflict:
- `discovery-2024-05-28-20240528-143022.log`

## Troubleshooting

### Logs folder is full
- Run `.\Archive-OldLogs.ps1` manually
- Check if archival is running automatically (should run on menu.ps1 startup)
- Verify the `logs/old/` folder exists and is writable

### Can't find old logs
- Check `logs/old/` subfolder
- Logs older than 90 days are automatically deleted
- Check Windows Recycle Bin if recently deleted

### Archival not working automatically
- Verify `menu.ps1` contains the `Move-OldLogs` call
- Check PowerShell execution policy allows script execution
- Run manually to test: `.\Archive-OldLogs.ps1`

### Archive folder has duplicate files
- This is normal - files with same name get timestamped to prevent overwrites
- Format: `filename-YYYYMMDD-HHMMSS.log`

## Best Practices

1. **Regular Review**: Review current logs weekly
2. **Archive Important Logs**: Before archival, copy any logs you want to keep permanently
3. **Disk Space**: Monitor disk space if running heavy logging operations
4. **Manual Cleanup**: If you have many old logs, run manual archival before automatic cleanup kicks in
5. **Backup**: Consider backing up the entire `logs/old/` folder periodically if logs are critical

## Disabling Auto-Archival

If you want to disable automatic log archival:

1. Open `menu.ps1`
2. Comment out the line:
   ```powershell
   # Move-OldLogs -DaysOld 7
   ```
3. Save the file

**Note**: This is not recommended as logs can accumulate and consume disk space.

## Manual Log Cleanup

To manually clean all logs:

```powershell
# Remove all current logs (CAUTION: This deletes logs!)
Remove-Item .\logs\*.log -Force

# Remove all archived logs
Remove-Item .\logs\old\*.log -Force

# Remove everything including archive folder
Remove-Item .\logs\* -Recurse -Force
```

**Warning**: Be careful with manual cleanup commands. Always verify the path before running.

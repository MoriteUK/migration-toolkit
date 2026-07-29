# Universal Logging Module for Migration Toolkit
# Provides standardized logging to files for all scripts

# Initialize logging for a script
function Initialize-Logging {
    param(
        [string]$ScriptName,
        [string]$LogDir = (Join-Path (Split-Path $PSScriptRoot -Parent) "Logs")
    )

    # Ensure log directory exists
    if (-not (Test-Path $LogDir)) {
        New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
    }

    # Create log file with timestamp
    $timestamp = Get-Date -Format 'yyyy-MM-dd_HHmmss'
    $logFileName = "$ScriptName`_$timestamp.log"
    $script:LogFilePath = Join-Path $LogDir $logFileName

    # Write initial log entry
    $startMessage = @"
═══════════════════════════════════════════════════════════════
 $ScriptName Started
 Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
 User: $env:USERNAME
 Computer: $env:COMPUTERNAME
═══════════════════════════════════════════════════════════════
"@

    $startMessage | Out-File -FilePath $script:LogFilePath -Encoding UTF8

    return $script:LogFilePath
}

# Write a log entry
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'SUCCESS', 'WARN', 'ERROR', 'DEBUG')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
    $logEntry = "$timestamp  [$($Level.PadRight(7))]  $Message"

    # Write to console with color
    $color = switch ($Level) {
        'SUCCESS' { 'Green' }
        'WARN'    { 'Yellow' }
        'ERROR'   { 'Red' }
        'DEBUG'   { 'DarkGray' }
        default   { 'White' }
    }

    Write-Host $logEntry -ForegroundColor $color

    # Write to log file if initialized
    if ($script:LogFilePath) {
        $logEntry | Out-File -FilePath $script:LogFilePath -Append -Encoding UTF8
    }
}

# Write an exception to the log
function Write-LogException {
    param(
        [System.Management.Automation.ErrorRecord]$Exception,
        [string]$Context = ""
    )

    if ($Context) {
        Write-Log "Exception in $Context" -Level ERROR
    }

    Write-Log "Exception: $($Exception.Exception.Message)" -Level ERROR
    Write-Log "  Script: $($Exception.InvocationInfo.ScriptName):$($Exception.InvocationInfo.ScriptLineNumber)" -Level ERROR

    if ($Exception.ScriptStackTrace) {
        Write-Log "  Stack Trace:" -Level DEBUG
        $Exception.ScriptStackTrace -split "`n" | ForEach-Object {
            Write-Log "    $_" -Level DEBUG
        }
    }
}

# Finish logging
function Complete-Logging {
    param(
        [bool]$Success = $true,
        [string]$Summary = ""
    )

    $endMessage = @"

═══════════════════════════════════════════════════════════════
 Completed: $(if ($Success) { 'SUCCESS' } else { 'FAILED' })
 Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
$(if ($Summary) { " Summary: $Summary`n" } else { "" })═══════════════════════════════════════════════════════════════
"@

    if ($script:LogFilePath) {
        $endMessage | Out-File -FilePath $script:LogFilePath -Append -Encoding UTF8

        if ($Success) {
            Write-Host "`nLog saved: $script:LogFilePath" -ForegroundColor Green
        } else {
            Write-Host "`nLog saved: $script:LogFilePath" -ForegroundColor Yellow
        }
    }
}

# Export functions
Export-ModuleMember -Function Initialize-Logging, Write-Log, Write-LogException, Complete-Logging

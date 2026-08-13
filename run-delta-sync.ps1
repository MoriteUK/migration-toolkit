param(
    [int]$IntervalSeconds = 120
)

$pass = 0
while ($true) {
    $pass++
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Pass $pass - Starting Delta sync cycle" -ForegroundColor Magenta
    try {
        Start-ADSyncSyncCycle -PolicyType Delta | Out-Null
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Pass $pass - Delta sync cycle triggered" -ForegroundColor Green
    } catch {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Pass $pass - Failed: $($_.Exception.Message)" -ForegroundColor Red
    }

    Start-Sleep -Seconds $IntervalSeconds
}

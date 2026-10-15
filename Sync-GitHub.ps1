#Requires -Version 5.0
<#
.SYNOPSIS
    Stage all changes, commit, and push to GitHub in one step.

.PARAMETER Message
    Commit message. Defaults to a timestamp if not provided.

.EXAMPLE
    .\Sync-GitHub.ps1 "Fix connection dropdown bug"
    .\Sync-GitHub.ps1   # uses auto timestamp message
#>
param([string]$Message = '')

$git = 'C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer\Git\cmd\git.exe'
$root = $PSScriptRoot

if (-not (Test-Path $git)) {
    Write-Error "git not found at: $git"
    exit 1
}

Set-Location $root

$status = & $git status --short
if (-not $status) {
    Write-Host 'Nothing to commit.' -ForegroundColor Yellow
    exit 0
}

Write-Host $status

if (-not $Message) {
    $Message = "Update $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
}

& $git add .
& $git commit -m $Message
# post-commit hook pushes automatically

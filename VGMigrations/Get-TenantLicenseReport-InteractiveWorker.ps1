#Requires -Version 7.0
<#
.SYNOPSIS
    Get-TenantLicenseReport-InteractiveWorker.ps1 — isolated child-process worker for
    interactive Graph sign-in, spawned fresh per tenant by Get-TenantLicenseReport.ps1.

.DESCRIPTION
    Interactive sign-in was consistently failing with a blank "authentication failed:" error
    for tenants processed later in a run, even when the user completed sign-in correctly in
    the browser each time — while the first tenant or two in the same run succeeded fine. That
    pattern points to state (most likely a loopback redirect listener MSAL uses to catch the
    browser's callback) getting stuck and not fully releasing between tenants within one
    long-running process, so a later browser sign-in has nowhere to land even though it
    completes. Running each interactive attempt in its own fresh process guarantees no such
    state can ever leak between tenants, regardless of the exact underlying cause.

    Writes a single JSON object to -OutputJsonPath: { success, error, orgId, orgName, skus }.
    Never touches the tenants workbook or the secret columns — only takes a tenant ID.
#>
param(
    [Parameter(Mandatory)][string]$TenantId,
    [Parameter(Mandatory)][string]$OutputJsonPath
)

$ErrorActionPreference = 'Stop'

function Get-CleanErrorMessage($ErrorRecord) {
    $lines = @($ErrorRecord.Exception.Message -split "`r?`n" | Where-Object { $_.Trim() })
    if ($lines.Count -eq 0) { return $ErrorRecord.Exception.GetType().Name }
    return ($lines | Select-Object -First 3) -join ' | '
}

$result = [ordered]@{ success = $false; error = $null; orgId = $null; orgName = $null; skus = @() }

try {
    Write-Host "=== Isolated interactive sign-in for tenant $TenantId ===" -ForegroundColor Cyan
    . (Join-Path $PSScriptRoot 'Ensure-GraphModules.ps1') -GraphModules @('Microsoft.Graph.Identity.DirectoryManagement')

    # MSAL/WAM broker diagnostics get written directly to .NET's Console.Out/Error, bypassing
    # every PowerShell stream — capture them so a failure shows a real reason.
    $consoleSwallow = New-Object System.IO.StringWriter
    $origOut = [Console]::Out
    $origErr = [Console]::Error
    [Console]::SetOut($consoleSwallow)
    [Console]::SetError($consoleSwallow)
    try {
        Connect-MgGraph -TenantId $TenantId -Scopes 'Organization.Read.All' -NoWelcome -ErrorAction Stop | Out-Null
        $org = Get-MgOrganization -ErrorAction Stop | Select-Object -First 1
    } catch {
        [Console]::SetOut($origOut); [Console]::SetError($origErr)
        $swallowed = $consoleSwallow.ToString().Trim()
        $msg = Get-CleanErrorMessage $_
        if ($swallowed) { $msg = "$msg — broker output: $($swallowed -replace '[\r\n]+', ' ')" }
        throw $msg
    } finally {
        if ([Console]::Out -ne $origOut) { [Console]::SetOut($origOut) }
        if ([Console]::Error -ne $origErr) { [Console]::SetError($origErr) }
        $consoleSwallow.Dispose()
    }

    $skus = @(Get-MgSubscribedSku -All -ErrorAction Stop)

    $result.success = $true
    $result.orgId   = $org.Id
    $result.orgName = $org.DisplayName
    $result.skus    = @($skus | ForEach-Object {
        [ordered]@{
            SkuPartNumber    = $_.SkuPartNumber
            Enabled          = $_.PrepaidUnits.Enabled
            Consumed         = $_.ConsumedUnits
            Suspended        = $_.PrepaidUnits.Suspended
            Warning          = $_.PrepaidUnits.Warning
            ServicePlanNames = @($_.ServicePlans | ForEach-Object { $_.ServicePlanName })
        }
    })
    Write-Host "Signed in OK — $($skus.Count) SKU(s) found." -ForegroundColor Green
} catch {
    $result.error = "$_"
    Write-Host "FAILED: $($result.error)" -ForegroundColor Red
} finally {
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
    ($result | ConvertTo-Json -Depth 6) | Out-File -FilePath $OutputJsonPath -Encoding utf8 -Force
}

Start-Sleep -Milliseconds 800

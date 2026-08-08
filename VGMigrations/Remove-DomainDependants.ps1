param(
    [string]$OldDomain = "intranote.dk",
    [string]$NewDomain = "ourvolaris.onmicrosoft.com",
    [string]$LogPath   = "C:\m365discovery\intranote_remediation_$(Get-Date -Format yyyyMMdd_HHmmss).csv",
    [switch]$Apply      # omit this switch = dry run (WhatIf). Add -Apply to actually make changes.
)

$log = [System.Collections.Generic.List[object]]::new()
$counts = @{ Success = 0; Failed = 0; Pending = 0; ManualStep = 0 }

function Add-Log($source,$type,$identity,$oldAddr,$action,$result) {
    $log.Add([PSCustomObject]@{
        Source   = $source; Type = $type; Identity = $identity
        OldAddress = $oldAddr; Action = $action; Result = $result
    })

    $color = switch -Wildcard ($result) {
        "Success*"    { "Green";  $counts.Success++;    break }
        "FAILED*"     { "Red";    $counts.Failed++;     break }
        "Pending"     { "Yellow"; $counts.Pending++;    break }
        "ManualStep"  { "Cyan";   $counts.ManualStep++; break }
        default       { "Gray" }
    }
    Write-Host "[$source/$type] " -NoNewline -ForegroundColor DarkGray
    Write-Host "$identity " -NoNewline -ForegroundColor White
    Write-Host "-> $action " -NoNewline
    Write-Host "[$result]" -ForegroundColor $color
}

Write-Host "`n=== Mode: $(if ($Apply) { 'APPLY (live changes)' } else { 'DRY RUN (no changes made)' }) ===" -ForegroundColor Magenta
Write-Host "Old domain: $OldDomain  ->  New domain: $NewDomain`n" -ForegroundColor Magenta

# ===================== ENTRA ID USERS =====================
Write-Host "--- Scanning Entra ID objects ---" -ForegroundColor Blue
$ids = (Get-MgDomainNameReference -DomainId $OldDomain).Id
$objs = if ($ids) { Get-MgDirectoryObjectById -Ids $ids } else { @() }
$users = $objs | Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.user' }
Write-Host "Found $($users.Count) user object(s) referencing $OldDomain`n" -ForegroundColor Blue

foreach ($u in $users) {
    $userId = $u.Id
    $upn    = $u.AdditionalProperties['userPrincipalName']
    $proxies = $u.AdditionalProperties['proxyAddresses']

    # --- UPN ---
    if ($upn -like "*@$OldDomain") {
        $localPart = $upn.Split('@')[0]
        $newUpn = "$localPart@$NewDomain"
        if ($Apply) {
            try {
                Update-MgUser -UserId $userId -UserPrincipalName $newUpn -ErrorAction Stop
                Add-Log "EntraID" "User-UPN" $upn $upn "Set UPN to $newUpn" "Success"
            } catch {
                Add-Log "EntraID" "User-UPN" $upn $upn "Set UPN to $newUpn" "FAILED (likely conflict) - MANUAL REVIEW: $($_.Exception.Message)"
            }
        } else {
            Add-Log "EntraID" "User-UPN" $upn $upn "[DryRun] Would set UPN to $newUpn" "Pending"
        }
    }

    # --- Proxy addresses (aliases) ---
    foreach ($p in $proxies) {
        if ($p -match "@$OldDomain$") {
            $prefix    = ($p -split ':')[0]
            $localPart = ($p -split ':')[1].Split('@')[0]
            $newAddr   = "$prefix`:$localPart@$NewDomain"
            Add-Log "EntraID" "User-ProxyAddress" $upn $p "Requires full-array PATCH - flagged for manual step" "ManualStep"
        }
    }
}

# ===================== EXCHANGE ONLINE OBJECTS =====================
Write-Host "`n--- Scanning Exchange Online objects ---" -ForegroundColor Blue
$exoTypes = @(
    @{ Cmdlet = 'Get-Mailbox';          SetCmdlet = 'Set-Mailbox';          Label = 'Mailbox' }
    @{ Cmdlet = 'Get-MailContact';      SetCmdlet = 'Set-MailContact';      Label = 'MailContact' }
    @{ Cmdlet = 'Get-DistributionGroup';SetCmdlet = 'Set-DistributionGroup';Label = 'DistributionGroup' }
    @{ Cmdlet = 'Get-UnifiedGroup';     SetCmdlet = 'Set-UnifiedGroup';     Label = 'M365Group' }
    @{ Cmdlet = 'Get-MailPublicFolder'; SetCmdlet = 'Set-MailPublicFolder'; Label = 'MailPublicFolder' }
)

foreach ($t in $exoTypes) {
    Write-Host "Checking $($t.Label)..." -ForegroundColor DarkCyan
    $items = & $t.Cmdlet -ResultSize Unlimited | Where-Object { $_.EmailAddresses -match $OldDomain }
    Write-Host "  Found $($items.Count) $($t.Label) object(s)" -ForegroundColor DarkCyan

    foreach ($item in $items) {
        $matching = $item.EmailAddresses | Where-Object { $_ -match "@$OldDomain$" }
        foreach ($addr in $matching) {
            $prefix    = ($addr -split ':')[0]
            $localPart = ($addr -split ':')[1].Split('@')[0]
            $newAddr   = "$prefix`:$localPart@$NewDomain"
            $identity  = $item.Identity

            if ($Apply) {
                try {
                    & $t.SetCmdlet -Identity $identity -EmailAddresses @{Add=$newAddr; Remove=$addr} -ErrorAction Stop
                    Add-Log "Exchange" $t.Label $identity $addr "Added $newAddr, removed old" "Success"
                } catch {
                    try {
                        & $t.SetCmdlet -Identity $identity -EmailAddresses @{Remove=$addr} -ErrorAction Stop
                        Add-Log "Exchange" $t.Label $identity $addr "Address in use elsewhere - removed old only" "Success (no replacement added)"
                    } catch {
                        Add-Log "Exchange" $t.Label $identity $addr "Remove old address" "FAILED: $($_.Exception.Message)"
                    }
                }
            } else {
                Add-Log "Exchange" $t.Label $identity $addr "[DryRun] Would add $newAddr, remove old (or remove-only if taken)" "Pending"
            }
        }
    }
}

$log | Export-Csv -Path $LogPath -NoTypeInformation

Write-Host "`n=== Summary ===" -ForegroundColor Magenta
Write-Host "Success:     $($counts.Success)" -ForegroundColor Green
Write-Host "Failed:      $($counts.Failed)" -ForegroundColor Red
Write-Host "Pending (dry run): $($counts.Pending)" -ForegroundColor Yellow
Write-Host "Manual step needed: $($counts.ManualStep)" -ForegroundColor Cyan
Write-Host "`nLog written to $LogPath"

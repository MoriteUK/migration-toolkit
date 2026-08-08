param(
    [string]$CsvPath   = "C:\m365discovery\intranotedk.onmicrosoft.com_AD_usage.csv",
    [string]$OldDomain = "intranotedk.onmicrosoft.com",
    [string]$NewDomain = "ourvolaris.onmicrosoft.com",
    [string]$LogPath   = "C:\m365discovery\AD_remediation_$(Get-Date -Format yyyyMMdd_HHmmss).csv",
    [switch]$Apply      # omit = dry run. Add -Apply to make real changes.
)

Import-Module ActiveDirectory

$rows = Import-Csv -Path $CsvPath
$log = [System.Collections.Generic.List[object]]::new()
$counts = @{ Success = 0; Failed = 0; Pending = 0 }

function Add-Log($dn,$attr,$oldVal,$newVal,$result) {
    $log.Add([PSCustomObject]@{ DN=$dn; Attribute=$attr; OldValue=$oldVal; NewValue=$newVal; Result=$result })
    $color = switch -Wildcard ($result) {
        "Success*" { "Green";  $counts.Success++; break }
        "FAILED*"  { "Red";    $counts.Failed++;  break }
        "Pending"  { "Yellow"; $counts.Pending++; break }
        default    { "Gray" }
    }
    Write-Host "[$attr] $dn : $oldVal -> $newVal " -NoNewline
    Write-Host "[$result]" -ForegroundColor $color
}

Write-Host "`n=== Mode: $(if ($Apply) { 'APPLY (live changes)' } else { 'DRY RUN (no changes made)' }) ===" -ForegroundColor Magenta
Write-Host "$OldDomain -> $NewDomain`n" -ForegroundColor Magenta
Write-Host "Processing $($rows.Count) object(s) from $CsvPath`n" -ForegroundColor Blue

foreach ($row in $rows) {
    $dn = $row.DN
    try {
        $obj = Get-ADObject -Identity $dn -Properties mail, userPrincipalName, proxyAddresses -ErrorAction Stop
    } catch {
        Add-Log $dn "Lookup" "" "" "FAILED: object not found ($($_.Exception.Message))"
        continue
    }

    # --- scalar attributes (mail, UPN) ---
    $scalarReplace = @{}
    if ($obj.mail -match [regex]::Escape($OldDomain)) {
        $newMail = $obj.mail -replace [regex]::Escape($OldDomain), $NewDomain
        $scalarReplace['mail'] = $newMail
    }
    if ($obj.userPrincipalName -and $obj.userPrincipalName -match [regex]::Escape($OldDomain)) {
        $newUpn = $obj.userPrincipalName -replace [regex]::Escape($OldDomain), $NewDomain
        $scalarReplace['userPrincipalName'] = $newUpn
    }

    if ($scalarReplace.Count -gt 0) {
        $preview = ($scalarReplace.Keys | ForEach-Object { "$_`: $($scalarReplace[$_])" }) -join ' | '
        if ($Apply) {
            try {
                Set-ADObject -Identity $dn -Replace $scalarReplace -ErrorAction Stop
                Add-Log $dn "mail/UPN" $preview "" "Success"
            } catch {
                Add-Log $dn "mail/UPN" $preview "" "FAILED: $($_.Exception.Message) - MANUAL REVIEW"
            }
        } else {
            Add-Log $dn "mail/UPN" $preview "" "Pending"
        }
    }

    # --- proxyAddresses (array, must be its own call) ---
    if ($obj.proxyAddresses) {
        $newProxies = @($obj.proxyAddresses | ForEach-Object {
            if ($_ -match [regex]::Escape($OldDomain)) { $_ -replace [regex]::Escape($OldDomain), $NewDomain } else { $_ }
        })
        if (($newProxies -join ',') -ne (@($obj.proxyAddresses) -join ',')) {
            $changedCount = (@($obj.proxyAddresses) | Where-Object { $_ -match [regex]::Escape($OldDomain) }).Count
            $preview = "proxyAddresses ($changedCount entries updated)"
            if ($Apply) {
                try {
                    Set-ADObject -Identity $dn -Replace @{ proxyAddresses = $newProxies } -ErrorAction Stop
                    Add-Log $dn "proxyAddresses" $preview "" "Success"
                } catch {
                    Add-Log $dn "proxyAddresses" $preview "" "FAILED: $($_.Exception.Message) - MANUAL REVIEW"
                }
            } else {
                Add-Log $dn "proxyAddresses" $preview "" "Pending"
            }
        }
    }

    if ($scalarReplace.Count -eq 0 -and -not $obj.proxyAddresses) {
        Add-Log $dn "None" "" "" "Pending"
    }
}

$log | Export-Csv -Path $LogPath -NoTypeInformation

Write-Host "`n=== Summary ===" -ForegroundColor Magenta
Write-Host "Success: $($counts.Success)" -ForegroundColor Green
Write-Host "Failed:  $($counts.Failed)" -ForegroundColor Red
Write-Host "Pending (dry run): $($counts.Pending)" -ForegroundColor Yellow
Write-Host "`nLog written to $LogPath"

param(
    [string]$CsvPath   = "C:\temp\bravurasecurity_com_AD_usage.csv",
    [string]$OldDomain = "bravurasecurity.com",
    [string]$NewDomain = "ourvolaris.onmicrosoft.com",
    [string]$LogPath   = "C:\temp\bravurasecurity_com_remediation_$(Get-Date -Format yyyyMMdd_HHmmss).csv",
    [switch]$Apply      # omit = dry run. Add -Apply to make real changes.
)

Import-Module ActiveDirectory

$rows = Import-Csv -Path $CsvPath
$log = [System.Collections.Generic.List[object]]::new()
$counts = @{ Success = 0; Failed = 0; Pending = 0; RemovedOnly = 0 }

function Add-Log($dn,$attr,$detail,$result) {
    $log.Add([PSCustomObject]@{ DN=$dn; Attribute=$attr; Detail=$detail; Result=$result })
    $color = switch -Wildcard ($result) {
        "Success*"     { "Green";  $counts.Success++;     break }
        "RemovedOnly*" { "DarkYellow"; $counts.RemovedOnly++; break }
        "FAILED*"      { "Red";    $counts.Failed++;      break }
        "Pending"      { "Yellow"; $counts.Pending++;     break }
        default        { "Gray" }
    }
    Write-Host "[$attr] $dn : $detail " -NoNewline
    Write-Host "[$result]" -ForegroundColor $color
}

function Test-ValueInUseElsewhere($ldapFilter, $selfDN) {
    $hit = Get-ADObject -LDAPFilter $ldapFilter -ErrorAction SilentlyContinue
    return ($hit | Where-Object { $_.DistinguishedName -ne $selfDN }) -ne $null
}

Write-Host "`n=== Mode: $(if ($Apply) { 'APPLY (live changes)' } else { 'DRY RUN (no changes made)' }) ===" -ForegroundColor Magenta
Write-Host "$OldDomain -> $NewDomain`n" -ForegroundColor Magenta

$rows = $rows | Sort-Object DN -Unique
Write-Host "Processing $($rows.Count) object(s) from $CsvPath`n" -ForegroundColor Blue

foreach ($row in $rows) {
    $dn = $row.DN
    try {
        $obj = Get-ADObject -Identity $dn -Properties mail, userPrincipalName, proxyAddresses -ErrorAction Stop
    } catch {
        Add-Log $dn "Lookup" "" "FAILED: object not found ($($_.Exception.Message))"
        continue
    }

    # --- mail (single-valued: swap, or clear if new value already used elsewhere) ---
    if ($obj.mail -match [regex]::Escape($OldDomain)) {
        $newMail = $obj.mail -replace [regex]::Escape($OldDomain), $NewDomain
        $conflict = Test-ValueInUseElsewhere "(mail=$newMail)" $dn

        if ($Apply) {
            try {
                if ($conflict) {
                    Set-ADObject -Identity $dn -Clear mail -ErrorAction Stop
                    Add-Log $dn "mail" "$($obj.mail) already used elsewhere - cleared" "RemovedOnly"
                } else {
                    Set-ADObject -Identity $dn -Replace @{ mail = $newMail } -ErrorAction Stop
                    Add-Log $dn "mail" "$($obj.mail) -> $newMail" "Success"
                }
            } catch {
                Add-Log $dn "mail" "$($obj.mail) -> $newMail" "FAILED: $($_.Exception.Message) - MANUAL REVIEW"
            }
        } else {
            $action = if ($conflict) { "[DryRun] Would CLEAR (target already in use)" } else { "[DryRun] Would swap to $newMail" }
            Add-Log $dn "mail" "$($obj.mail) $action" "Pending"
        }
    }

    # --- userPrincipalName (single-valued, MUST remain set: swap only, conflict = manual review) ---
    if ($obj.userPrincipalName -and $obj.userPrincipalName -match [regex]::Escape($OldDomain)) {
        $newUpn = $obj.userPrincipalName -replace [regex]::Escape($OldDomain), $NewDomain
        $conflict = Test-ValueInUseElsewhere "(userPrincipalName=$newUpn)" $dn

        if ($conflict) {
            Add-Log $dn "UPN" "$($obj.userPrincipalName) -> $newUpn already in use - NOT cleared (UPN required)" "FAILED: conflict - MANUAL REVIEW"
        } elseif ($Apply) {
            try {
                Set-ADObject -Identity $dn -Replace @{ userPrincipalName = $newUpn } -ErrorAction Stop
                Add-Log $dn "UPN" "$($obj.userPrincipalName) -> $newUpn" "Success"
            } catch {
                Add-Log $dn "UPN" "$($obj.userPrincipalName) -> $newUpn" "FAILED: $($_.Exception.Message) - MANUAL REVIEW"
            }
        } else {
            Add-Log $dn "UPN" "$($obj.userPrincipalName) -> $newUpn" "Pending"
        }
    }

    # --- proxyAddresses (multi-valued: swap each, or remove-only if target already used elsewhere) ---
    $oldEntries = @($obj.proxyAddresses) | Where-Object { $_ -match [regex]::Escape($OldDomain) } | ForEach-Object { [string]$_ }
    if ($oldEntries.Count -gt 0) {
        $toRemove = @()
        $toAdd = @()
        $detailParts = @()

        foreach ($old in $oldEntries) {
            $prefix    = ($old -split ':')[0]
            $localPart = ($old -split ':',2)[1].Split('@')[0]
            $newAddr   = "$prefix`:$localPart@$NewDomain"
            $conflict  = Test-ValueInUseElsewhere "(proxyAddresses=$newAddr)" $dn

            $toRemove += $old
            if ($conflict) {
                $detailParts += "$old -> REMOVED ONLY (target in use)"
            } else {
                $toAdd += $newAddr
                $detailParts += "$old -> $newAddr"
            }
        }

        if ($Apply) {
            try {
                if ($toAdd.Count -gt 0) {
                    Set-ADObject -Identity $dn -Remove @{ proxyAddresses = $toRemove } -Add @{ proxyAddresses = $toAdd } -ErrorAction Stop
                } else {
                    Set-ADObject -Identity $dn -Remove @{ proxyAddresses = $toRemove } -ErrorAction Stop
                }
                $result = if ($toAdd.Count -lt $toRemove.Count) { "RemovedOnly (partial)" } else { "Success" }
                Add-Log $dn "proxyAddresses" ($detailParts -join ' | ') $result
            } catch {
                Add-Log $dn "proxyAddresses" ($detailParts -join ' | ') "FAILED: $($_.Exception.Message) - MANUAL REVIEW"
            }
        } else {
            Add-Log $dn "proxyAddresses" ($detailParts -join ' | ') "Pending"
        }
    }
}

$log | Export-Csv -Path $LogPath -NoTypeInformation

Write-Host "`n=== Summary ===" -ForegroundColor Magenta
Write-Host "Success:        $($counts.Success)" -ForegroundColor Green
Write-Host "Removed only:   $($counts.RemovedOnly)" -ForegroundColor DarkYellow
Write-Host "Failed:         $($counts.Failed)" -ForegroundColor Red
Write-Host "Pending (dry run): $($counts.Pending)" -ForegroundColor Yellow
Write-Host "`nLog written to $LogPath"
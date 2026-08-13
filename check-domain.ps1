$domain = "bravurasecurity.com"
$outDir = "C:\temp"
$pollSeconds = 300
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

Import-Module ActiveDirectory

function Test-Match($value) {
    if (-not $value) { return $false }
    return ($value -match [regex]::Escape($domain))
}

function Get-DomainMatches {
    $results = [System.Collections.Generic.List[object]]::new()

    $users = Get-ADUser -Filter * -Properties UserPrincipalName, mail, proxyAddresses, mailNickname, DisplayName
    foreach ($u in $users) {
        $hits = @()
        if (Test-Match $u.UserPrincipalName) { $hits += "UPN:$($u.UserPrincipalName)" }
        if (Test-Match $u.mail)              { $hits += "mail:$($u.mail)" }
        foreach ($p in $u.proxyAddresses) { if (Test-Match $p) { $hits += "proxy:$p" } }
        if ($hits.Count -gt 0) {
            $results.Add([PSCustomObject]@{
                Source = "AD-User"; SamAccountName = $u.SamAccountName; DisplayName = $u.DisplayName
                DN = $u.DistinguishedName; Matches = ($hits -join '; ')
            })
        }
    }

    $groups = Get-ADGroup -Filter * -Properties mail, proxyAddresses, DisplayName
    foreach ($g in $groups) {
        $hits = @()
        if (Test-Match $g.mail) { $hits += "mail:$($g.mail)" }
        foreach ($p in $g.proxyAddresses) { if (Test-Match $p) { $hits += "proxy:$p" } }
        if ($hits.Count -gt 0) {
            $results.Add([PSCustomObject]@{
                Source = "AD-Group"; SamAccountName = $g.SamAccountName; DisplayName = $g.DisplayName
                DN = $g.DistinguishedName; Matches = ($hits -join '; ')
            })
        }
    }

    $contacts = Get-ADObject -Filter { objectClass -eq "contact" } -Properties mail, proxyAddresses, DisplayName
    foreach ($c in $contacts) {
        $hits = @()
        if (Test-Match $c.mail) { $hits += "mail:$($c.mail)" }
        foreach ($p in $c.proxyAddresses) { if (Test-Match $p) { $hits += "proxy:$p" } }
        if ($hits.Count -gt 0) {
            $results.Add([PSCustomObject]@{
                Source = "AD-Contact"; SamAccountName = ""; DisplayName = $c.DisplayName
                DN = $c.DistinguishedName; Matches = ($hits -join '; ')
            })
        }
    }

    return [PSCustomObject]@{
        Results      = $results
        UserCount    = $users.Count
        GroupCount   = $groups.Count
        ContactCount = $contacts.Count
    }
}

$csvPath = Join-Path $outDir "${domain}_AD_usage.csv"
$pass = 0

do {
    $pass++
    $scan = Get-DomainMatches
    $results = $scan.Results

    $userHits    = ($results | Where-Object Source -eq "AD-User").Count
    $groupHits   = ($results | Where-Object Source -eq "AD-Group").Count
    $contactHits = ($results | Where-Object Source -eq "AD-Contact").Count

    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Pass $pass - Users: $($scan.UserCount) scanned / $userHits matches | Groups: $($scan.GroupCount) scanned / $groupHits matches | Contacts: $($scan.ContactCount) scanned / $contactHits matches | Total remaining: $($results.Count)" -ForegroundColor Magenta

    $results | Export-Csv -Path $csvPath -NoTypeInformation

    if ($results.Count -eq 0) { break }

    Start-Sleep -Seconds $pollSeconds
} while ($true)

Write-Host "=== Done - 0 objects referencing @$domain remain ===" -ForegroundColor Magenta
Write-Host "Last export: $csvPath" -ForegroundColor Magenta

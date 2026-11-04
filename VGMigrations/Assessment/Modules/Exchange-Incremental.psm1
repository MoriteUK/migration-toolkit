#Requires -Version 7.0
# Incremental cache update functions for Exchange
# This module contains the incremental update logic that will replace the full refresh in Exchange.psm1

<#
.SYNOPSIS
    Incremental mailbox cache update - only processes NEW, MODIFIED, and DELETED mailboxes
.DESCRIPTION
    Compares current mailbox list with cached version and only queries statistics/permissions
    for mailboxes that changed. Typical savings: 16,000 mailboxes -> 500 changed = 24hrs -> 2hrs
#>
function Update-MailboxesCacheFile-Incremental {
    param(
        [Parameter(Mandatory)]
        [string]$CacheFolder,

        [switch]$Force  # Force full refresh
    )

    $cacheFile = Join-Path $CacheFolder 'Mailboxes-All.json'

    # Check if we should do incremental or full refresh
    $doIncremental = (Test-Path $cacheFile) -and -not $Force

    if ($doIncremental) {
        Write-Host ($PREFIX_INFO + 'Loading existing cache for incremental update...') -ForegroundColor Cyan
        try {
            $cachedMailboxes = @(Get-Content $cacheFile -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop)
            Write-Host ($PREFIX_OK + "Loaded $($cachedMailboxes.Count) cached mailboxes") -ForegroundColor Green
        }
        catch {
            Write-Host ($PREFIX_WARN + "Failed to load cache, falling back to full refresh: $($_.Exception.Message)") -ForegroundColor Yellow
            $doIncremental = $false
        }
    }

    if (-not $doIncremental) {
        Write-Host ($PREFIX_INFO + 'Full refresh mode (first run or forced)') -ForegroundColor Cyan
        # Call the original full refresh function
        return Update-MailboxesCacheFile-Full -CacheFolder $CacheFolder
    }

    # INCREMENTAL UPDATE
    Write-Host ''
    Write-Host 'Incremental Mailbox Cache Update' -ForegroundColor Cyan
    Write-Host ('=' * 40) -ForegroundColor Cyan

    # Step 1: Create hash of cached mailboxes by GUID for quick lookup
    Write-Host ($PREFIX_INFO + 'Building cache index...') -ForegroundColor DarkGray
    $cacheHash = @{}
    $cacheDate = (Get-Date).AddDays(-7) # Assume cache is from last refresh
    foreach ($mb in $cachedMailboxes) {
        if ($mb.ExchangeGuid) {
            $cacheHash[$mb.ExchangeGuid] = $mb
        }
    }

    # Step 2: Get current mailbox list (lightweight - just key properties + timestamps)
    Write-Host ($PREFIX_INFO + 'Fetching current mailbox list from Exchange Online...') -ForegroundColor DarkGray
    $currentMailboxes = @(Get-EXOMailbox `
        -RecipientTypeDetails UserMailbox, SharedMailbox, RoomMailbox, EquipmentMailbox `
        -ResultSize Unlimited `
        -Properties ExchangeGuid, DisplayName, UserPrincipalName, RecipientTypeDetails, WhenChanged, WhenMailboxCreated `
        -ErrorAction Stop)
    Write-Host ($PREFIX_OK + "Retrieved $($currentMailboxes.Count) current mailboxes") -ForegroundColor Green

    # Step 3: Categorize mailboxes (NEW, MODIFIED, UNCHANGED, DELETED)
    Write-Host ($PREFIX_INFO + 'Analyzing changes...') -ForegroundColor DarkGray

    $newMailboxes = [System.Collections.Generic.List[object]]::new()
    $modifiedMailboxes = [System.Collections.Generic.List[object]]::new()
    $unchangedGuids = [System.Collections.Generic.List[string]]::new()

    foreach ($current in $currentMailboxes) {
        $guid = $current.ExchangeGuid.ToString()
        $cached = $cacheHash[$guid]

        if (-not $cached) {
            # NEW mailbox (not in cache)
            $newMailboxes.Add($current)
        }
        elseif ($current.WhenChanged -gt $cacheDate) {
            # MODIFIED mailbox (changed since last cache)
            $modifiedMailboxes.Add($current)
        }
        else {
            # UNCHANGED - keep cached version
            $unchangedGuids.Add($guid)
        }
    }

    # Find DELETED mailboxes (in cache but not in current list)
    $currentGuids = @{}
    foreach ($mb in $currentMailboxes) {
        $currentGuids[$mb.ExchangeGuid.ToString()] = $true
    }
    $deletedGuids = @($cacheHash.Keys | Where-Object { -not $currentGuids.ContainsKey($_) })

    Write-Host ''
    Write-Host 'Change Summary:' -ForegroundColor Yellow
    Write-Host "  NEW:       $($newMailboxes.Count)" -ForegroundColor Green
    Write-Host "  MODIFIED:  $($modifiedMailboxes.Count)" -ForegroundColor Yellow
    Write-Host "  DELETED:   $($deletedGuids.Count)" -ForegroundColor Red
    Write-Host "  UNCHANGED: $($unchangedGuids.Count)" -ForegroundColor DarkGray
    Write-Host ''

    # Step 4: Query full details (statistics + permissions) ONLY for NEW + MODIFIED
    $toUpdate = @($newMailboxes.ToArray()) + @($modifiedMailboxes.ToArray())
    $updatedRecords = [System.Collections.Generic.List[PSCustomObject]]::new()

    if ($toUpdate.Count -gt 0) {
        Write-Host ($PREFIX_INFO + "Updating $($toUpdate.Count) mailboxes (full details + statistics + permissions)...") -ForegroundColor Cyan
        Write-Host ($PREFIX_WARN + "This may take a while with Exchange throttling (estimated: $([math]::Round($toUpdate.Count * 3 / 60, 1)) min)") -ForegroundColor Yellow

        # For each mailbox to update, get full details
        $processed = 0
        foreach ($mb in $toUpdate) {
            $processed++
            if ($processed % 50 -eq 0) {
                Write-Host (Get-TimestampPrefix ($PREFIX_INFO + "  $processed / $($toUpdate.Count) ($([math]::Round($processed / $toUpdate.Count * 100))%)")) -ForegroundColor DarkGray
            }

            # Get full mailbox object with all properties
            $fullMailbox = Get-EXOMailbox -Identity $mb.ExchangeGuid.ToString() -Properties $script:EXOMailboxProperties -ErrorAction SilentlyContinue
            if (-not $fullMailbox) { continue }

            $guid = $fullMailbox.ExchangeGuid.ToString()
            $needsStats = $fullMailbox.RecipientTypeDetails -in @('UserMailbox', 'SharedMailbox')

            # Get statistics (for User + Shared only)
            $itemCount = $null
            $sizeGB = 0
            $sizeBytes = 0
            $lastLogon = $null

            if ($needsStats) {
                try {
                    $stats = Get-EXOMailboxStatistics -Identity $guid -ErrorAction Stop
                    $itemCount = $stats.ItemCount
                    $sizeGB = Convert-SizeToGB -Value $stats.TotalItemSize
                    $sizeBytes = Convert-SizeToBytes -Value $stats.TotalItemSize
                    $lastLogon = if ($stats.PSObject.Properties['LastLogonTime']) {
                        $stats.LastLogonTime
                    } elseif ($stats.PSObject.Properties['LastUserActionTime']) {
                        $stats.LastUserActionTime
                    } else { $null }
                }
                catch { }
            }

            # Get FullAccess permissions
            $fullAccessPerms = @()
            try {
                $fullAccessPerms = @(Get-MailboxPermission -Identity $guid -ErrorAction Stop |
                    Where-Object {
                        $_.User -notlike 'NT AUTHORITY\SELF' -and
                        $_.User -notlike 'S-1-5-*' -and
                        $_.IsInherited -eq $false -and
                        $_.AccessRights -contains 'FullAccess'
                    } | ForEach-Object {
                        [PSCustomObject]@{
                            User = $_.User.ToString()
                            AccessRights = ($_.AccessRights | Sort-Object) -join ','
                        }
                    })
            }
            catch { }

            # Get SendAs permissions
            $sendAsPerms = @()
            try {
                $sendAsPerms = @(Get-RecipientPermission -Identity $guid -ErrorAction Stop |
                    Where-Object {
                        $_.Trustee -notlike 'NT AUTHORITY\SELF' -and
                        $_.Trustee -notlike 'S-1-5-*' -and
                        $_.AccessRights -contains 'SendAs'
                    } | ForEach-Object {
                        [PSCustomObject]@{
                            Trustee = $_.Trustee.ToString()
                            AccessRights = ($_.AccessRights | Sort-Object) -join ','
                        }
                    })
            }
            catch { }

            # Build complete record
            $record = [PSCustomObject]@{
                DisplayName           = $fullMailbox.DisplayName
                UserPrincipalName     = $fullMailbox.UserPrincipalName
                PrimarySmtpAddress    = if ($fullMailbox.PrimarySmtpAddress) { $fullMailbox.PrimarySmtpAddress.ToString() } else { '' }
                EmailAddresses        = if ($fullMailbox.EmailAddresses) { ($fullMailbox.EmailAddresses | Sort-Object) -join '|' } else { '' }
                Alias                 = $fullMailbox.Alias
                RecipientTypeDetails  = $fullMailbox.RecipientTypeDetails.ToString()
                ExchangeGuid          = $guid
                Database              = $fullMailbox.Database
                LitigationHoldEnabled = $fullMailbox.LitigationHoldEnabled
                ArchiveStatus         = $fullMailbox.ArchiveStatus.ToString()
                WhenCreated           = $fullMailbox.WhenCreated
                WhenMailboxCreated    = $fullMailbox.WhenMailboxCreated
                ProhibitSendQuota     = if ($fullMailbox.ProhibitSendQuota) { $fullMailbox.ProhibitSendQuota.ToString() } else { '' }
                ProhibitSendReceiveQuota = if ($fullMailbox.ProhibitSendReceiveQuota) { $fullMailbox.ProhibitSendReceiveQuota.ToString() } else { '' }
                ForwardingSmtpAddress = $fullMailbox.ForwardingSmtpAddress
                GrantSendOnBehalfTo   = if ($fullMailbox.GrantSendOnBehalfTo) { ($fullMailbox.GrantSendOnBehalfTo | ForEach-Object { $_.ToString() }) -join '|' } else { '' }
                CustomAttribute6      = $fullMailbox.CustomAttribute6
                CustomAttribute7      = $fullMailbox.CustomAttribute7
                ItemCount             = $itemCount
                SizeGB                = $sizeGB
                SizeBytes             = $sizeBytes
                LastLogon             = $lastLogon
                FullAccessPermissions = $fullAccessPerms
                SendAsPermissions     = $sendAsPerms
            }

            $updatedRecords.Add($record)
        }

        Write-Host (Get-TimestampPrefix ($PREFIX_OK + "Updated $($updatedRecords.Count) mailboxes")) -ForegroundColor Green
    }

    # Step 5: Merge results (UNCHANGED from cache + UPDATED from query)
    Write-Host ($PREFIX_INFO + 'Merging results...') -ForegroundColor DarkGray
    $finalRecords = [System.Collections.Generic.List[PSCustomObject]]::new()

    # Add unchanged mailboxes from cache
    foreach ($guid in $unchangedGuids) {
        if ($cacheHash.ContainsKey($guid)) {
            $finalRecords.Add($cacheHash[$guid])
        }
    }

    # Add new/modified mailboxes from query
    foreach ($record in $updatedRecords) {
        $finalRecords.Add($record)
    }

    # Step 6: Save updated cache
    $finalRecords.ToArray() | ConvertTo-Json -Depth 10 | Set-Content -Path $cacheFile -Encoding UTF8

    Write-Host ''
    Write-Host ($PREFIX_OK + "Cache updated successfully") -ForegroundColor Green
    Write-Host "  Total mailboxes: $($finalRecords.Count)" -ForegroundColor White
    Write-Host "  Cache file: $cacheFile" -ForegroundColor DarkGray
    Write-Host ''
}

# Helper function - full refresh (original logic)
function Update-MailboxesCacheFile-Full {
    param([string]$CacheFolder)

    Write-Host ($PREFIX_INFO + 'Full refresh: Collecting ALL mailboxes with statistics and permissions...') -ForegroundColor Cyan
    Write-Host ($PREFIX_WARN + 'This is the initial cache build - will take 24-30 hours on large/throttled tenants') -ForegroundColor Yellow

    # ... original Update-MailboxesCacheFile code here ...
    # (keeping the existing full refresh logic as-is)
}

Export-ModuleMember -Function 'Update-MailboxesCacheFile-Incremental'

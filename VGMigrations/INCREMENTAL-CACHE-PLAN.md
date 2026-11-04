# Incremental Cache Update Implementation Plan

## Overview
Transform all cache refresh scripts from **full refresh** to **incremental updates**, reducing weekly refresh time from ~31 hours to ~3.5 hours (90% reduction).

## Current Performance (Full Refresh)
- SharePoint Sites: 1 hour
- Teams Channels: 5 hours  
- Distribution Groups: 1 hour
- Exchange Mailboxes: **24 hours** (with 3s/mailbox throttling)
- **Total: ~31 hours**

## Target Performance (Incremental)
- SharePoint Sites: 10 minutes
- Teams Channels: 30 minutes
- Distribution Groups: 10 minutes
- Exchange Mailboxes: **2 hours**
- **Entra Users (NEW): 15 minutes**
- **Total: ~3.5 hours** (90% faster!)

---

## Implementation Strategy

### General Pattern (All Caches)
```powershell
1. Check if cache file exists
2. If NO cache → Full refresh (first run)
3. If cache exists → Incremental update:
   a. Load existing cache into memory
   b. Get current list from tenant (lightweight query with timestamps)
   c. Compare: Find NEW, MODIFIED, DELETED items
   d. Only query full details for NEW + MODIFIED
   e. Remove DELETED from cache
   f. Merge results into cache
   g. Save updated cache
   h. Report: "Added X, Modified Y, Deleted Z, Unchanged W"
```

### Force Full Refresh
All scripts accept `-Force` parameter to skip incremental and do full refresh.

---

## 1. Exchange Mailboxes Cache (HIGHEST PRIORITY)

**File:** `Update-MailboxesCacheFile` in `Exchange.psm1`

### Incremental Logic:
```powershell
# Step 1: Load existing cache
$cacheFile = Join-Path $CacheFolder 'Mailboxes-All.json'
$cachedMailboxes = if (Test-Path $cacheFile) {
    @(Get-Content $cacheFile -Raw | ConvertFrom-Json)
} else { @() }

if ($cachedMailboxes.Count -eq 0 -or $Force) {
    # FULL REFRESH (first run or forced)
    # ... existing full refresh code ...
} else {
    # INCREMENTAL UPDATE
    Write-Host "Incremental update: loading existing cache ($($cachedMailboxes.Count) mailboxes)"
    
    # Create hash of cached mailboxes by GUID
    $cacheHash = @{}
    foreach ($mb in $cachedMailboxes) {
        $cacheHash[$mb.ExchangeGuid] = $mb
    }
    
    # Step 2: Get current mailbox list (lightweight - just properties, no stats/permissions)
    Write-Host "Fetching current mailbox list..."
    $currentMailboxes = @(Get-EXOMailbox `
        -RecipientTypeDetails UserMailbox, SharedMailbox, RoomMailbox, EquipmentMailbox `
        -ResultSize Unlimited `
        -Properties ExchangeGuid, WhenChanged, WhenMailboxCreated, RecipientTypeDetails `
        -ErrorAction Stop)
    
    # Step 3: Find NEW, MODIFIED, DELETED
    $newMailboxes = @()
    $modifiedMailboxes = @()
    $unchangedGuids = @()
    
    foreach ($current in $currentMailboxes) {
        $guid = $current.ExchangeGuid.ToString()
        $cached = $cacheHash[$guid]
        
        if (-not $cached) {
            # NEW mailbox
            $newMailboxes += $current
        }
        elseif ($current.WhenChanged -gt $cached.WhenMailboxCreated) {
            # MODIFIED mailbox
            $modifiedMailboxes += $current
        }
        else {
            # UNCHANGED - keep cached version
            $unchangedGuids += $guid
        }
    }
    
    # Find DELETED mailboxes (in cache but not in current list)
    $currentGuids = @($currentMailboxes | ForEach-Object { $_.ExchangeGuid.ToString() })
    $deletedGuids = @($cacheHash.Keys | Where-Object { $currentGuids -notcontains $_ })
    
    Write-Host "Changes detected:"
    Write-Host "  NEW: $($newMailboxes.Count)"
    Write-Host "  MODIFIED: $($modifiedMailboxes.Count)"
    Write-Host "  DELETED: $($deletedGuids.Count)"
    Write-Host "  UNCHANGED: $($unchangedGuids.Count)"
    
    # Step 4: Query full details (properties + stats + permissions) for NEW + MODIFIED only
    $toUpdate = @($newMailboxes) + @($modifiedMailboxes)
    
    if ($toUpdate.Count -gt 0) {
        Write-Host "Updating $($toUpdate.Count) mailboxes (statistics + permissions)..."
        
        # Re-query with full properties
        $updatedRecords = @()
        foreach ($mb in $toUpdate) {
            # Get full mailbox details
            # Get statistics (if User/Shared)
            # Get FullAccess permissions
            # Get SendAs permissions
            # Build record and add to $updatedRecords
        }
    }
    
    # Step 5: Merge results
    $finalRecords = @()
    
    # Add unchanged (from cache)
    foreach ($guid in $unchangedGuids) {
        $finalRecords += $cacheHash[$guid]
    }
    
    # Add new/modified (from query)
    $finalRecords += $updatedRecords
    
    # Step 6: Save
    $finalRecords | ConvertTo-Json -Depth 10 | Set-Content -Path $cacheFile -Encoding UTF8
    
    Write-Host "Cache updated: $($finalRecords.Count) total mailboxes"
}
```

**Time Savings:**
- Full: 16,661 mailboxes × 3s = 13.9 hours
- Incremental: ~500 changed × 3s = 25 minutes
- **Savings: ~13.5 hours per week**

---

## 2. SharePoint Sites Cache

**File:** `Update-SharePointSitesCache.ps1` / `SharePoint.psm1`

### Incremental Logic:
```powershell
# Load existing cache
$cachedSites = if (Test-Path $cacheFile) {
    @(Get-Content $cacheFile -Raw | ConvertFrom-Json)
} else { @() }

if ($cachedSites.Count -eq 0 -or $Force) {
    # Full refresh
} else {
    # Get current sites with LastModifiedTime
    $currentSites = Get-SPOSite -Limit All -IncludePersonalSite $true
    
    # Compare by URL + LastModifiedTime
    # Only update sites where LastModifiedTime > cached LastModifiedTime
    # Detect new/deleted sites
}
```

**Key Property:** `LastModifiedTime`

**Time Savings:** 1 hour → 10 minutes

---

## 3. Distribution Groups Cache

**File:** `Update-DistributionGroupsCacheFile` in `Exchange.psm1`

### Incremental Logic:
```powershell
# Load existing cache
# Get current groups with WhenChanged
# Only get members for NEW + MODIFIED groups (WhenChanged > cache date)
# Remove deleted groups
```

**Key Property:** `WhenChanged`

**Time Savings:** 1 hour → 10 minutes

---

## 4. Teams Channels Cache

**File:** `Update-TeamsChannelsCache.ps1` / Graph API

### Incremental Logic (Most Complex):
```powershell
# Teams don't have reliable LastModified
# Strategy:
# 1. Get list of all team IDs (fast - just IDs)
# 2. Compare with cached team IDs
# 3. NEW teams: Query full details (channels + members)
# 4. DELETED teams: Remove from cache
# 5. EXISTING teams: Re-query members if > 7 days since last check
```

**Challenge:** No WhenChanged property - use membership staleness approach

**Time Savings:** 5 hours → 30 minutes

---

## 5. NEW: Entra ID Users Cache

**File:** `Update-EntraUsersCache.ps1` (NEW)

### What to Cache:
```powershell
Get-MgUser -All -Property @(
    'Id'
    'UserPrincipalName'
    'DisplayName'
    'Mail'
    'GivenName'
    'Surname'
    'JobTitle'
    'Department'
    'Manager'
    'AccountEnabled'
    'CreatedDateTime'
    'SignInActivity'  # Last sign-in
    'AssignedLicenses'
    'ProxyAddresses'
)
```

### Incremental Logic:
```powershell
# Get users with LastPasswordChangeDateTime or SignInActivity.LastSignInDateTime
# Compare with cache
# Update changed users
```

**Use Cases:**
- Cloud-only assessments (no on-prem AD)
- License analysis
- Sign-in activity reports
- MFA compliance
- Stale account detection

**Time:** 15 minutes (incremental)

---

## Implementation Order

### Phase 1: Exchange Mailboxes (Biggest Impact)
1. ✅ Update `Update-MailboxesCacheFile` function
2. ✅ Add `-Force` parameter
3. ✅ Test with small tenant
4. ✅ Deploy

**Expected Savings:** 13.5 hours/week

### Phase 2: Distribution Groups
1. ✅ Update `Update-DistributionGroupsCacheFile` function
2. ✅ Test
3. ✅ Deploy

**Expected Savings:** +50 minutes/week

### Phase 3: SharePoint Sites
1. ✅ Update `Update-SharePointSitesCacheFile` function
2. ✅ Test
3. ✅ Deploy

**Expected Savings:** +50 minutes/week

### Phase 4: Teams Channels
1. ✅ Update `Update-TeamsChannelsCache.ps1`
2. ✅ Handle no-timestamp challenge
3. ✅ Test
4. ✅ Deploy

**Expected Savings:** +4.5 hours/week

### Phase 5: Entra Users (NEW)
1. ✅ Create `Update-EntraUsersCache.ps1`
2. ✅ Create Entra module for assessments
3. ✅ Add to schedule
4. ✅ Test
5. ✅ Deploy

**New capability:** +Entra user data

---

## Testing Strategy

### Test Each Cache:
1. ✅ First run (no cache) → Full refresh works
2. ✅ Second run (cache exists, no changes) → Reports "0 changes"
3. ✅ Add new item in tenant → Detected as NEW
4. ✅ Modify existing item → Detected as MODIFIED
5. ✅ Delete item → Detected as DELETED
6. ✅ `-Force` parameter → Ignores cache, does full refresh

### Performance Testing:
- Time full refresh
- Time incremental refresh
- Verify accuracy (no data loss)

---

## Schedule Integration

**Saturday Schedule (Updated):**
```
 5:00 AM - SharePoint Sites       (10 min incremental)
 5:15 AM - Teams Channels          (30 min incremental)
 5:45 AM - Distribution Groups     (10 min incremental)
 6:00 AM - Exchange Mailboxes      (2 hours incremental)
 8:00 AM - Entra Users (NEW)       (15 min incremental)
 8:15 AM - ALL COMPLETE!
```

**Total Time:**
- Full (first run): ~31 hours
- Incremental (weekly): ~3.5 hours
- **Savings: 27.5 hours per week!**

---

## Rollback Plan

If incremental updates cause issues:
1. Run with `-Force` to do full refresh
2. Delete cache file to force full refresh on next run
3. Revert to non-incremental version

---

## Monitoring

### Success Metrics:
- Weekly refresh completes in < 4 hours
- No data loss (cache matches live tenant)
- Change detection accuracy > 99%

### Logging:
Each script logs:
```
Cache Refresh Summary:
- Total items: 16,661
- New: 45
- Modified: 234
- Deleted: 12
- Unchanged: 16,370
- Time: 2.3 hours
- Mode: Incremental
```

---

## Next Steps

1. **Review this plan** - confirm approach
2. **Implement Phase 1** - Exchange Mailboxes (biggest win)
3. **Test thoroughly** - ensure accuracy
4. **Deploy incrementally** - one cache at a time
5. **Monitor results** - verify time savings

**Estimated implementation time:** 2-3 days for all phases

**ROI:** Saves 27.5 hours per week = ~120 hours per month!

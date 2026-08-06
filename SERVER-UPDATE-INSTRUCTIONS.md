# Server Update to Version 2.9.40

Since git is not installed on the server, follow these manual steps:

## Step 1: Copy Files to Server

Copy these 3 files from your local machine to the server's VGMigrations folder:

From: `C:\toolkit\VGMigrations\`

Files to copy:
1. `Check-DomainMigrationReadiness.ps1`
2. `version.json`
3. `Update-ToV2.9.40.ps1` (verification script)

To: Server location (e.g., `C:\Toolkit\VGMigrations\` or wherever the toolkit is installed)

## Step 2: Verify on Server

On the server, run PowerShell 7 and execute:

```powershell
cd C:\Toolkit\VGMigrations  # Adjust path as needed
Get-Content version.json | Select-String '"version"'
```

Should show: `"version": "2.9.40"`

## Step 3: Test

Run the domain readiness check and verify you see the new debug output:
- `DNS: Querying intranote.dk via 8.8.8.8...`
- `Domain verification status: IsVerified=...`
- `PRE-CUTOVER SCENARIO DETECTED` or `NORMAL SCENARIO`

## What's New in 2.9.40

- DNS queries now use public DNS servers (Google 8.8.8.8, Cloudflare 1.1.1.1) to avoid cached results
- Enhanced detection of unverified domains (added to target but still in source tenant)
- Comprehensive DNS record logging with byte-level comparison
- Improved pre-cutover validation (exact match first, then any MS= record)
- Better status reporting: "Ready - Pre-Cutover (Exact Match)"
- Warns when source/target MS= records differ

## Alternative: Use Remote PowerShell from Your Machine

If you have admin access to the server, you can copy files remotely:

```powershell
# From your local machine
$serverPath = "\\ServerName\C$\Toolkit\VGMigrations"
Copy-Item "C:\toolkit\VGMigrations\Check-DomainMigrationReadiness.ps1" $serverPath -Force
Copy-Item "C:\toolkit\VGMigrations\version.json" $serverPath -Force
```

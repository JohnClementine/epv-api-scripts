# Step 2: Sync safes from SOURCE to DESTINATION (new) environment
# Run this from the "Migration via REST" directory
# Requires: ExportOfAccounts.csv from Step 1

# Force .NET to use WinHttpHandler (trusts Windows cert store like your browser)
$env:DOTNET_SYSTEM_NET_HTTP_USESOCKETSHTTPHANDLER = '0'

# --- Config ---
$SourcePVWAURL  = "https://prdcapw1-esaw2a.uprising.t-mobile.com/PasswordVault"
$SourceAuthType = "ldap"

$DestPVWAURL    = "https://prdcarkpvwa1-esaw2a.uprising.t-mobile.com/passwordvault"
$DestAuthType   = "cyberark"

$CPMOverride    = "PasswordManager"

# --- Go ---
Import-Module '.\Migrate.psm1' -Force

Write-Host "Loading accounts from CSV..." -ForegroundColor Cyan
Import-Accounts -importCSV ".\ExportOfAccounts.csv"

Write-Host "Connecting to SOURCE..." -ForegroundColor Cyan
New-SourceSession -srcPVWAURL $SourcePVWAURL -srcAuthType $SourceAuthType

Write-Host "Connecting to DESTINATION..." -ForegroundColor Cyan
New-DestinationSession -dstPVWAURL $DestPVWAURL -dstAuthType $DestAuthType

Write-Host "Syncing safes..." -ForegroundColor Cyan
Sync-Safes -CreateSafes -UpdateSafeMembers -CPMOverride $CPMOverride -OwnersToExclude @("Safe_Admin") -maxJobCount 10

Write-Host "Done! Verify safes in the destination PVWA." -ForegroundColor Green
Write-Host "Then run Step3-SyncAccounts.ps1" -ForegroundColor Yellow

Close-SourceSession
Close-DestinationSession

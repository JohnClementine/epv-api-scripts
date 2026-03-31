# Step 1: Export accounts (and safe info) from the SOURCE (old) environment
# Run this from the "Migration via REST" directory

# --- Config ---
$SourcePVWAURL  = "https://prdcapw1-esaw2a.uprising.t-mobile.com/PasswordVault"
$SourceAuthType = "ldap"

# --- Go ---
Import-Module '.\Migrate.psm1' -Force

Write-Host "Connecting to SOURCE..." -ForegroundColor Cyan
New-SourceSession -srcPVWAURL $SourcePVWAURL -srcAuthType $SourceAuthType

Write-Host "Exporting accounts to ExportOfAccounts.csv..." -ForegroundColor Cyan
Export-Accounts

Write-Host "Done! Review ExportOfAccounts.csv, remove rows you don't want to migrate." -ForegroundColor Green
Write-Host "Then run Step2-SyncSafes.ps1" -ForegroundColor Yellow

Close-SourceSession

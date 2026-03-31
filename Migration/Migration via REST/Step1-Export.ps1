# Step 1: Export accounts (and safe info) from the SOURCE (old) environment
# Run this from the "Migration via REST" directory

# Force .NET to use WinHttpHandler (trusts Windows cert store like your browser)
$env:DOTNET_SYSTEM_NET_HTTP_USESOCKETSHTTPHANDLER = '0'

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

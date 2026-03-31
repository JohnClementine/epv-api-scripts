# Step 3: Sync accounts from SOURCE to DESTINATION (new) environment
# Run this from the "Migration via REST" directory
# Requires: ExportOfAccounts.csv from Step 1

# --- SSL Bypass (PowerShell 7) ---
# Global scope so it reaches inside module function calls
$global:PSDefaultParameterValues['Invoke-WebRequest:SkipCertificateCheck']  = $true
$global:PSDefaultParameterValues['Invoke-RestMethod:SkipCertificateCheck']  = $true
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

# --- Config ---
$SourcePVWAURL  = "https://prdcapw1-esaw2a.uprising.t-mobile.com/PasswordVault"
$SourceAuthType = "ldap"

$DestPVWAURL    = "https://prdcarkpvwa1-esaw2a.uprising.t-mobile.com/passwordvault"
$DestAuthType   = "cyberark"

# --- Go ---
Import-Module '.\Migrate.psm1' -Force

Write-Host "Loading accounts from CSV..." -ForegroundColor Cyan
Import-Accounts -importCSV ".\ExportOfAccounts.csv"

Write-Host "Connecting to SOURCE..." -ForegroundColor Cyan
New-SourceSession -srcPVWAURL $SourcePVWAURL -srcAuthType $SourceAuthType

Write-Host "Connecting to DESTINATION..." -ForegroundColor Cyan
New-DestinationSession -dstPVWAURL $DestPVWAURL -dstAuthType $DestAuthType

Write-Host "Syncing accounts..." -ForegroundColor Cyan
Sync-Accounts -VerifyPlatform -maxJobCount 10

Write-Host "Done! Verify accounts in the destination PVWA." -ForegroundColor Green
Write-Host "Recommended: delete the import user and ensure only one env has active CPMs." -ForegroundColor Yellow

Close-SourceSession
Close-DestinationSession

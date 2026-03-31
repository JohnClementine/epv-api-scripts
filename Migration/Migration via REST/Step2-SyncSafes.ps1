# Step 2: Sync safes from SOURCE to DESTINATION (new) environment
# Run this from the "Migration via REST" directory
# Requires: ExportOfAccounts.csv from Step 1

# --- SSL Bypass (PowerShell 7 - GLOBAL scope so it reaches inside modules) ---
$global:PSDefaultParameterValues['Invoke-WebRequest:SkipCertificateCheck']  = $true
$global:PSDefaultParameterValues['Invoke-RestMethod:SkipCertificateCheck']  = $true
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

# Suppress the "It is not Recommended to disable SSL" Inquire prompt inside the module
$global:WarningPreference = 'SilentlyContinue'

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
New-SourceSession -srcPVWAURL $SourcePVWAURL -srcAuthType $SourceAuthType -DisableSSLVerify

Write-Host "Connecting to DESTINATION..." -ForegroundColor Cyan
New-DestinationSession -dstPVWAURL $DestPVWAURL -dstAuthType $DestAuthType -DisableSSLVerify

Write-Host "Syncing safes..." -ForegroundColor Cyan
Sync-Safes -CreateSafes -UpdateSafeMembers -CPMOverride $CPMOverride -maxJobCount 10

Write-Host "Done! Verify safes in the destination PVWA." -ForegroundColor Green
Write-Host "Then run Step3-SyncAccounts.ps1" -ForegroundColor Yellow

Close-SourceSession
Close-DestinationSession

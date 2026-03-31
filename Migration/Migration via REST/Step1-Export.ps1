# Step 1: Export accounts (and safe info) from the SOURCE (old) environment
# Run this from the "Migration via REST" directory

# --- SSL Bypass (PowerShell 7) ---
# Global scope so it reaches inside module function calls
$global:PSDefaultParameterValues['Invoke-WebRequest:SkipCertificateCheck']  = $true
$global:PSDefaultParameterValues['Invoke-RestMethod:SkipCertificateCheck']  = $true
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

# --- Config ---
$SourcePVWAURL  = "https://prdcapw1-esaw2a.uprising.t-mobile.com/PasswordVault"
$SourceAuthType = "ldap"

# --- Go ---
Import-Module '.\Migrate.psm1' -Force

Write-Host "Connecting to SOURCE..." -ForegroundColor Cyan
# NOTE: Do NOT pass -DisableSSLVerify here. The module's SSL bypass has a
# -WarningAction Inquire prompt that crashes portable pwsh. We already
# disabled SSL globally above, so the module's Invoke-WebRequest and
# Invoke-RestMethod calls will pick up -SkipCertificateCheck automatically.
New-SourceSession -srcPVWAURL $SourcePVWAURL -srcAuthType $SourceAuthType

Write-Host "Exporting accounts to ExportOfAccounts.csv..." -ForegroundColor Cyan
Export-Accounts

Write-Host "Done! Review ExportOfAccounts.csv, remove rows you don't want to migrate." -ForegroundColor Green
Write-Host "Then run Step2-SyncSafes.ps1" -ForegroundColor Yellow

Close-SourceSession

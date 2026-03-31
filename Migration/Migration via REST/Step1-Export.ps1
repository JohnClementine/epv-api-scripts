# Step 1: Export accounts (and safe info) from the SOURCE (old) environment
# Run this from the "Migration via REST" directory

# --- SSL Bypass (PowerShell 7) ---
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

# --- Config ---
$SourcePVWAURL  = "https://prdcapw1-esaw2a.uprising.t-mobile.com/PasswordVault"
$SourceAuthType = "ldap"

# --- Go ---
Import-Module '.\Migrate.psm1' -Force

# Inject -SkipCertificateCheck into the MODULE's session state.
# $global:PSDefaultParameterValues does NOT reach inside modules - they have
# their own isolated scope. This reaches into each module and sets it there.
& (Get-Module 'Migrate') {
    $PSDefaultParameterValues['Invoke-WebRequest:SkipCertificateCheck']  = $true
    $PSDefaultParameterValues['Invoke-RestMethod:SkipCertificateCheck']  = $true
}
& (Get-Module 'CyberArk-Migration') {
    $PSDefaultParameterValues['Invoke-WebRequest:SkipCertificateCheck']  = $true
    $PSDefaultParameterValues['Invoke-RestMethod:SkipCertificateCheck']  = $true
}

Write-Host "Connecting to SOURCE..." -ForegroundColor Cyan
New-SourceSession -srcPVWAURL $SourcePVWAURL -srcAuthType $SourceAuthType

Write-Host "Exporting accounts to ExportOfAccounts.csv..." -ForegroundColor Cyan
Export-Accounts

Write-Host "Done! Review ExportOfAccounts.csv, remove rows you don't want to migrate." -ForegroundColor Green
Write-Host "Then run Step2-SyncSafes.ps1" -ForegroundColor Yellow

Close-SourceSession

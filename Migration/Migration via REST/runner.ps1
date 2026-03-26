<#
.SYNOPSIS
    Runner script to sync safes (and optionally accounts) from an older on-prem
    CyberArk environment to a newer on-prem environment using the Migration REST module.

.DESCRIPTION
    This script automates the workflow described in the Migration via REST README:
      1. Import the migration module
      2. Connect to the source environment
      3. Export accounts (to discover safes)
      4. Optionally review/import a filtered account list
      5. Connect to the destination environment
      6. Sync safes (create + set memberships)
      7. Optionally sync accounts

    Edit the CONFIGURATION section below before running.

.NOTES
    - Requires PowerShell 6.0 or greater
    - All three module files (Migrate.psm1, CyberArk-Migration.psm1, Invoke-Process.ps1)
      must be in the same directory as this script.
    - Run from the directory containing these files, e.g.:
        cd "Migration via REST"
        .\runner.ps1
#>

#Requires -Version 6.0

# ============================================================================
# CONFIGURATION - Edit these values before running
# ============================================================================

# --- Source (old) environment ---
$SourcePVWAURL   = "https://old-pvwa.yourcompany.com/PasswordVault"
$SourceAuthType  = "cyberark"   # cyberark, ldap, or radius

# --- Destination (new) environment ---
$DestPVWAURL     = "https://new-pvwa.yourcompany.com/PasswordVault"
$DestAuthType    = "cyberark"   # cyberark, ldap, or radius

# --- Safe sync options ---
$CreateSafes       = $true   # Create safes in destination if they don't exist
$UpdateSafeMembers = $true   # Update safe memberships to match source

# --- CPM mapping (optional) ---
# Set these if the CPM name differs between environments.
# Leave empty to keep the same CPM name.
$CPMOld       = ""            # e.g. "PasswordManager"
$CPMNew       = ""            # e.g. "PasswordManager_New"
$CPMOverride  = ""            # Set to force a single CPM for all safes (overrides Old/New)

# --- Additional owners to exclude from migration (optional) ---
# Built-in system owners are already excluded by the module.
$OwnersToExclude = @()        # e.g. @("SvcAccount1", "LegacyAdmin")

# --- Account list CSV (optional) ---
# After the initial export you can review/edit ExportOfAccounts.csv and
# re-import a filtered version. Set to $null to skip re-import and use the
# full export.
$FilteredCSV = $null          # e.g. ".\FilteredAccounts.csv"

# --- Account sync (optional) ---
$SyncAccountsToo   = $false   # Set to $true to also sync accounts after safes
$SkipCheckSecret   = $false   # Skip secret comparison on existing accounts
$GetRemoteMachines = $false   # Sync remote machine access lists
$VerifyPlatform    = $true    # Verify platforms exist in destination before creating accounts
$NoCreate          = $false   # Set to $true to only update existing accounts, not create new ones

# --- Parallelism ---
$MaxJobCount = 10             # Number of concurrent jobs for safe/account processing

# --- SSL ---
$DisableSSL = $false          # Set to $true only if using self-signed certs (NOT RECOMMENDED)

# ============================================================================
# EXECUTION - No changes needed below this line
# ============================================================================

$ErrorActionPreference = "Stop"

# Ensure we're running from the correct directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Push-Location $scriptDir

try {
    # ------------------------------------------------------------------
    # Step 1 - Load the migration module
    # ------------------------------------------------------------------
    Write-Host "`n=============================" -ForegroundColor Cyan
    Write-Host " Step 1: Loading module" -ForegroundColor Cyan
    Write-Host "=============================`n" -ForegroundColor Cyan

    Import-Module '.\Migrate.psm1' -Force

    # ------------------------------------------------------------------
    # Step 2 - Connect to the SOURCE environment
    # ------------------------------------------------------------------
    Write-Host "`n=============================" -ForegroundColor Cyan
    Write-Host " Step 2: Source session" -ForegroundColor Cyan
    Write-Host "=============================`n" -ForegroundColor Cyan

    $srcParams = @{
        srcPVWAURL  = $SourcePVWAURL
        srcAuthType = $SourceAuthType
    }
    if ($DisableSSL) { $srcParams["DisableSSLVerify"] = $true }

    New-SourceSession @srcParams

    # ------------------------------------------------------------------
    # Step 3 - Export accounts from source (discovers safes)
    # ------------------------------------------------------------------
    Write-Host "`n=============================" -ForegroundColor Cyan
    Write-Host " Step 3: Exporting accounts" -ForegroundColor Cyan
    Write-Host "=============================`n" -ForegroundColor Cyan

    Export-Accounts

    Write-Host "`nExported to ExportOfAccounts.csv - review the file if needed." -ForegroundColor Yellow

    # ------------------------------------------------------------------
    # Step 4 - (Optional) Import a filtered account list
    # ------------------------------------------------------------------
    if ($FilteredCSV) {
        Write-Host "`n=============================" -ForegroundColor Cyan
        Write-Host " Step 4: Importing filtered account list" -ForegroundColor Cyan
        Write-Host "=============================`n" -ForegroundColor Cyan

        Import-Accounts -importCSV $FilteredCSV
    }
    else {
        Write-Host "`nStep 4: Skipped (using full export)`n" -ForegroundColor DarkGray
    }

    # ------------------------------------------------------------------
    # Step 5 - Connect to the DESTINATION environment
    # ------------------------------------------------------------------
    Write-Host "`n=============================" -ForegroundColor Cyan
    Write-Host " Step 5: Destination session" -ForegroundColor Cyan
    Write-Host "=============================`n" -ForegroundColor Cyan

    $dstParams = @{
        dstPVWAURL  = $DestPVWAURL
        dstAuthType = $DestAuthType
    }
    if ($DisableSSL) { $dstParams["DisableSSLVerify"] = $true }

    New-DestinationSession @dstParams

    # ------------------------------------------------------------------
    # Step 6 - Sync safes
    # ------------------------------------------------------------------
    Write-Host "`n=============================" -ForegroundColor Cyan
    Write-Host " Step 6: Syncing safes" -ForegroundColor Cyan
    Write-Host "=============================`n" -ForegroundColor Cyan

    $safeParams = @{
        maxJobCount = $MaxJobCount
    }
    if ($CreateSafes)       { $safeParams["CreateSafes"]       = $true }
    if ($UpdateSafeMembers) { $safeParams["UpdateSafeMembers"] = $true }

    if ($CPMOverride) {
        $safeParams["CPMOverride"] = $CPMOverride
    }
    elseif ($CPMOld -and $CPMNew) {
        $safeParams["CPMOld"] = $CPMOld
        $safeParams["CPMNew"] = $CPMNew
    }

    if ($OwnersToExclude.Count -gt 0) {
        $safeParams["OwnersToExclude"] = $OwnersToExclude
    }

    Sync-Safes @safeParams

    Write-Host "`nSafe sync complete." -ForegroundColor Green

    # ------------------------------------------------------------------
    # Step 7 - (Optional) Sync accounts
    # ------------------------------------------------------------------
    if ($SyncAccountsToo) {
        Write-Host "`n=============================" -ForegroundColor Cyan
        Write-Host " Step 7: Syncing accounts" -ForegroundColor Cyan
        Write-Host "=============================`n" -ForegroundColor Cyan

        $acctParams = @{
            maxJobCount = $MaxJobCount
        }
        if ($SkipCheckSecret)   { $acctParams["SkipCheckSecret"]   = $true }
        if ($GetRemoteMachines) { $acctParams["getRemoteMachines"] = $true }
        if ($VerifyPlatform)    { $acctParams["VerifyPlatform"]    = $true }
        if ($NoCreate)          { $acctParams["noCreate"]          = $true }

        Sync-Accounts @acctParams

        Write-Host "`nAccount sync complete." -ForegroundColor Green
    }
    else {
        Write-Host "`nStep 7: Skipped (SyncAccountsToo = `$false)`n" -ForegroundColor DarkGray
    }

    # ------------------------------------------------------------------
    # Done
    # ------------------------------------------------------------------
    Write-Host "`n=============================" -ForegroundColor Green
    Write-Host " Migration complete!" -ForegroundColor Green
    Write-Host "=============================`n" -ForegroundColor Green
    Write-Host "Recommended next steps:" -ForegroundColor Yellow
    Write-Host "  1. Verify safes and memberships in the destination PVWA" -ForegroundColor Yellow
    Write-Host "  2. Delete the import user after verification to remove residual full-access" -ForegroundColor Yellow
    Write-Host "  3. Ensure only one environment has active CPMs per safe" -ForegroundColor Yellow
}
catch {
    Write-Host "`nERROR: $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
}
finally {
    # Clean up sessions
    try { Close-SourceSession } catch {}
    try { Close-DestinationSession } catch {}
    Pop-Location
}

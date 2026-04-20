# Resync Guide - Update Changed Credentials from Old to New Environment

## Context

- Safes already exist in the new environment (done in Step 2 previously)
- Accounts already exist in the new environment (done in Step 3 previously)
- CPM was turned on in the OLD environment, so some passwords may have been rotated
- Goal: pull the (possibly new) passwords from OLD and push them to NEW

## Key facts to know before starting

1. **Sync-Accounts will NOT duplicate accounts.** It matches by safe+name+address+username. If a match exists on the destination, it UPDATES (including syncing the password if different). If no match, it creates.

2. **The existing CSV is still valid.** The CSV only contains metadata (id, name, address, safe, platform, username). Passwords are always fetched live at sync time. Account IDs on the source don't change when CPM rotates passwords.

3. **Running Sync-Accounts triggers checkouts on the source.** Every `Get-Secret` call counts as a checkout. With CPM now ON, the old environment might auto-release them after a timeout (and rotate again). Or they may lock again. Either way, ReleaseAll.ps1 is the safety net.

4. **You cannot avoid a password change cycle if CPM is aggressive.** If CPM rotates on every checkout, every sync triggers another rotation. The new environment will get whatever was the password at the moment of retrieval.

## Tonight's runbook

### Step 0: Backup what you have

Before anything, back up the existing CSV in case you want to compare before/after.

```powershell
cd "C:\Cyberark-Software\Scripting\Migration"
Copy-Item .\ExportOfAccounts.csv .\ExportOfAccounts.backup-$(Get-Date -Format 'yyyyMMdd-HHmm').csv
```

### Step 1: Re-export from SOURCE (optional but recommended)

This refreshes the CSV in case any accounts were added/changed structurally on the source. It does NOT retrieve passwords, only metadata — so it will NOT trigger checkouts.

```powershell
.\step1.ps1
```

(Or whatever you renamed `Step1-Export.ps1` to.)

Expected output: "Exporting accounts to ExportOfAccounts.csv... Done!"

If the account list looks right and you don't want to re-export, you can skip this step.

### Step 2: SKIP Sync-Safes

Do NOT run `Step2-SyncSafes.ps1`. Safes already exist. Running it again is harmless but wastes time.

### Step 3: Re-run Sync-Accounts

This is the main event. It will:
- Connect to both environments
- For each account: retrieve the secret from SOURCE, compare to DESTINATION
- If they differ (which they will for any CPM-rotated accounts), UPDATE the destination
- Accounts where secrets already match are skipped (fast)

```powershell
.\step3.ps1
```

Watch the output. Per-account logs go to `.\LogFiles-Accounts\`. Expect it to take roughly as long as the original run.

### Step 4: Release locked accounts on SOURCE

Same as last time. Many accounts will be checked out by your admin user from the Get-Secret calls.

```powershell
.\ReleaseAll.ps1
```

Enter your old environment `administrator` credentials (or edit the `$AuthType` at the top of the script to `"ldap"` if you prefer to use your LDAP account).

## What to expect

- **Password updates in NEW env:** Any account whose password changed in OLD will get pushed to NEW.
- **Same 10-13 key-based account errors as last time:** SSH key accounts still can't sync via the password API. Ignore them or handle manually.
- **"Unable to locate destination account" messages:** These are usually false alarms — check if the account actually exists in the destination. If not, it's a real creation failure (usually a platform mismatch).

## If something goes wrong

- **All logs are in `.\LogFiles-Accounts\`** — one log file per account, named by safe+account+id.
- **The master log is `.\Migrate.log`** in the Migration directory.
- **Rerunning Step 3 is always safe** — it will re-check everything, skip matches, re-attempt failures.

## After you're done

1. Spot-check a handful of accounts in the NEW PVWA. Pick ones you know had activity — try to retrieve the password and confirm it matches what's currently in the OLD environment.
2. Run `.\ReleaseAll.ps1` one more time if you see lots of checked-out accounts in the OLD environment.
3. If your org plans to keep the OLD environment around for a bit: keep CPM on ONLY the NEW environment going forward to avoid dual-rotation confusion.

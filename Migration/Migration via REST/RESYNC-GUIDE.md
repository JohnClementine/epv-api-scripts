# Resync Guide - Update Changed Credentials from Old to New Environment

## Context

- Safes already exist in the new environment (done in Step 2 previously)
- Accounts already exist in the new environment (done in Step 3 previously)
- CPM was turned on in the OLD environment, so some passwords may have been rotated
- Goal: pull the (possibly new) passwords from OLD and push them to NEW
- **Plan: turn OFF the OLD CPM before the sync.** It's fine if accounts stay locked in the old environment — that environment is on its way out anyway. (ReleaseAll.ps1 never worked reliably, don't bother with it.)

## Key facts to know before starting

1. **Sync-Accounts will NOT duplicate accounts.** It matches by safe+name+address+username. If a match exists on the destination, it UPDATES (including syncing the password if different). If no match, it creates.

2. **The existing CSV is still valid.** The CSV only contains metadata (id, name, address, safe, platform, username). Passwords are always fetched live at sync time. Account IDs on the source don't change when CPM rotates passwords.

3. **Running Sync-Accounts will lock accounts on the source.** Every `Get-Secret` call counts as a checkout. With CPM turned OFF on the old side, locked accounts will just stay locked — which is fine since we're done with that environment.

4. **Turn OFF CPM on the OLD environment before syncing.** Otherwise CPM could rotate passwords mid-sync and you'd get stale values pushed to the new environment.

## Tonight's runbook

### Step 0a: Turn OFF the CPM on the OLD environment

Do this FIRST. This prevents CPM from rotating passwords while you're syncing, which would give the new environment stale values.

### Step 0b: Backup what you have

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

### Step 4: Don't bother releasing accounts on the OLD environment

With CPM off on the OLD side, locked accounts will just sit there. That's fine — we're done with that environment. Skip ReleaseAll.ps1.

## What to expect

- **Password updates in NEW env:** Any account whose password changed in OLD will get pushed to NEW.
- **Same 10-13 key-based account errors as last time:** SSH key accounts still can't sync via the password API. Ignore them or handle manually.
- **"Unable to locate destination account" messages:** These are usually false alarms — check if the account actually exists in the destination. If not, it's a real creation failure (usually a platform mismatch).

## If something goes wrong

- **All logs are in `.\LogFiles-Accounts\`** — one log file per account, named by safe+account+id.
- **The master log is `.\Migrate.log`** in the Migration directory.
- **Rerunning Step 3 is always safe** — it will re-check everything, skip matches, re-attempt failures.

## After you're done

1. Spot-check a handful of accounts in the NEW PVWA. Pick ones you know had activity — try to retrieve the password and confirm it works.
2. CPM stays OFF on the OLD environment and ON only on the NEW environment going forward. No dual-rotation confusion.

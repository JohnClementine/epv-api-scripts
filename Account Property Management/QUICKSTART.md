# Quick Start — Bulk-update CyberArk account fields from a spreadsheet

This tool lets you **export** your CyberArk accounts to a spreadsheet (CSV),
**edit** the values you care about (like **Address** and **Notes**) in Excel,
and **import** them back to update the matching accounts — all without touching
passwords. You do **not** need to know PowerShell. Just copy, paste, and run.

> 🔒 It never reads, changes, or rotates passwords. It only updates the fields you list.

---

## Before you start (one time)

1. Keep these files together in the same folder (you already have them):
   `Export-CyberArkAccountsToCsv.ps1`, `Import-CyberArkAccountPropertyUpdates.ps1`,
   `CyberArkAccountProperties.psm1`.
2. Open **PowerShell in that folder**: in File Explorer, go into the
   `Account Property Management` folder, then **Shift + right-click** an empty
   spot → **“Open PowerShell window here.”**
3. In the commands below, change only these **two** things to your own:
   - the URL `https://brambles.privilegecloud.cyberark.cloud/passwordvault`
   - the username `johnc@cyberark.cloud.38599`

When a command asks you to log in, type your password. If you’re shown a numbered
list of login methods, pick the number next to **“Enter Password.”**

---

## Step 1 — Export your accounts to a spreadsheet

Copy–paste this (one line), then press Enter:

```powershell
.\Export-CyberArkAccountsToCsv.ps1 -PVWAUrl 'https://brambles.privilegecloud.cyberark.cloud/passwordvault' -IdentityUserName 'johnc@cyberark.cloud.38599' -OutputCsv .\accounts.csv
```

This creates **`accounts.csv`** in the same folder.

## Step 2 — Edit the spreadsheet

Open **`accounts.csv`** in Excel and edit only the columns you want to change
(for example **`Address`** and **`Notes`**). Then **Save** (keep the `.csv`
format if Excel asks).

| Column | What to do |
|---|---|
| **`AccountID`** | **Leave it alone.** This is how each row is matched — don’t edit or delete it. |
| **`Address`**, **`Notes`** (and other named columns) | Edit these freely. Blank = leave that account’s value unchanged. |
| Columns starting with **`_`** (e.g. `_CreatedTime`) | Ignore them — they’re just info and are never changed. |

> Tip: You can delete columns you don’t need. Just **keep `AccountID`** and the
> columns you intend to update.

## Step 3 — Preview, then apply

**Always preview first** with `-WhatIf`. This shows exactly what *would* change
and writes **nothing**:

```powershell
.\Import-CyberArkAccountPropertyUpdates.ps1 -PVWAUrl 'https://brambles.privilegecloud.cyberark.cloud/passwordvault' -IdentityUserName 'johnc@cyberark.cloud.38599' -InputCsv .\accounts.csv -PropertiesToUpdate Address,Notes -WhatIf
```

You’ll see lines like:
`Row 1 [242_259] … WOULD update -> Address: '1.1.1.1' -> '9.9.9.9' | Notes: …`

When that looks right, run it **for real** by removing `-WhatIf`:

```powershell
.\Import-CyberArkAccountPropertyUpdates.ps1 -PVWAUrl 'https://brambles.privilegecloud.cyberark.cloud/passwordvault' -IdentityUserName 'johnc@cyberark.cloud.38599' -InputCsv .\accounts.csv -PropertiesToUpdate Address,Notes
```

At the end you’ll see a summary, e.g.
`Summary: Updated=10 NoChange=1 NotFound=1 Error=0`. A detailed
`…results.csv` file (one line per account) is saved in the folder.

> Want to update different fields? Change the names after `-PropertiesToUpdate`,
> separated by commas (e.g. `-PropertiesToUpdate Notes` or `-PropertiesToUpdate Address,Notes,Environment`).

---

## Golden rules

- ✅ **Preview with `-WhatIf` first**, every time.
- ✅ **Keep the `AccountID` column** — it’s the key that matches rows to accounts.
- ✅ A **blank cell means “don’t change”** — it won’t erase anything.
- ✅ One bad row won’t stop the rest; check the summary and `results.csv`.
- ❌ Passwords are never touched. `Address`, `Notes`, etc. are just labels/metadata.

## If something goes wrong

| Message | Fix |
|---|---|
| “running scripts is disabled on this system” | Run this once in the window, then retry: `Set-ExecutionPolicy -Scope Process Bypass -Force` |
| “The variable '$matches' …” at login | Check the URL is spelled correctly: `privilegecloud` (not `privilegeclou`). |
| Login shows Push/SMS but no password option, or it hangs | Pick the **“Enter Password”** option if shown. If only Push/SMS appear, ask your CyberArk admin for help or about an OAuth login. |
| `Status = NotFound` for a row | That `AccountID` no longer exists in CyberArk — safe to ignore or remove the row. |
| `Status = Error` for a row | Open the `…results.csv`; the **Error** column explains why (e.g. no permission on that safe). Other rows still processed. |

---

*Need more detail (other login methods, options, etc.)? See `README.md` in this folder.*

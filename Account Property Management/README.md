# Account Property Management (Export / Import custom File Categories)

A two-step PowerShell workflow for bulk-editing **custom File Categories /
account properties** (such as `Notes`) on CyberArk **Privilege Cloud (ISPSS)**
or self-hosted **PVWA** accounts:

1. **Export** every account (with pagination) to a CSV, including all custom
   File Categories returned by CyberArk.
2. **Edit** the CSV by hand to add/update property values.
3. **Import** the edited CSV to update the matching existing accounts, keyed by
   the CyberArk **account ID**.

> `Notes` is **not** a special password field. It is treated as an ordinary
> custom File Category / account property (`platformAccountProperties`), exactly
> like any other custom property returned by CyberArk.

These scripts **only read account metadata**. They never retrieve passwords,
never change passwords, and never trigger password-management actions. Updates
are restricted to `/platformAccountProperties/*` (File Categories).

> 👉 **New to PowerShell?** Start with [`QUICKSTART.md`](QUICKSTART.md) — a
> copy‑paste, one‑page guide. This README is the fuller reference.

---

## Files

| File | Purpose |
|------|---------|
| `CyberArk Account Field Updater - User Guide.docx` | Printable Word user guide for non-technical users. |
| `QUICKSTART.md` | One-page, copy-paste guide for non-technical users. |
| `Export-CyberArkAccountsToCsv.ps1` | Export all accounts + custom properties to CSV. |
| `Import-CyberArkAccountPropertyUpdates.ps1` | Update selected custom properties on existing accounts from CSV. |
| `CyberArkAccountProperties.psm1` | Shared helper module (auth, pagination, REST, flattening, JSON-Patch). |
| `sample-accounts.csv` | Small fictional example of the export/import CSV shape. |

The scripts auto-import the helper module from the same folder, so keep the
three `.ps1`/`.psm1` files together.

---

## Requirements

- Windows PowerShell **5.1+** or PowerShell **7+**.
- Network access to your CyberArk tenant (and to CyberArk Identity).
- A user/identity with permission to **list accounts**, **view account details**,
  and **update account properties** in the relevant safes.
- For ISPSS auth, the repo's
  [`Identity Authentication/IdentityAuth.psm1`](../Identity%20Authentication)
  module. The scripts load it automatically from that sibling folder; override
  with `-IdentityAuthModulePath`, or use `-DownloadIdentityAuth` to fetch it.

---

## Authentication

ISPSS authentication is delegated to the repository's **`IdentityAuth.psm1`**
(`Get-IdentityHeader`), so MFA / push / SAML+PIN are handled by that module.
Both scripts share the same options via `-PVWAUrl`, `-AuthType`, `-Credential`,
`-IdentityUserName` and `-LogonToken`.

### 1. Identity — interactive (default, recommended for a person at a keyboard)

Pass `-IdentityUserName`; the module prompts for the password and any MFA.

```powershell
.\Export-CyberArkAccountsToCsv.ps1 `
    -PVWAUrl 'https://mytenant.privilegecloud.cyberark.cloud/PasswordVault' `
    -IdentityUserName 'me@corp.com' `
    -OutputCsv .\accounts.csv
```

### 2. Identity — username + password (PSCredential)

```powershell
$cred = Get-Credential   # Identity username + password (MFA still prompts if configured)
.\Export-CyberArkAccountsToCsv.ps1 `
    -PVWAUrl 'https://mytenant.privilegecloud.cyberark.cloud/PasswordVault' `
    -Credential $cred `
    -OutputCsv .\accounts.csv
```

### 3. OAuth client credentials (headless automation)

`-Credential` username = OAuth **Client ID**, password = OAuth **Client Secret**.

```powershell
$oauth = Get-Credential   # Username = OAuth client ID, Password = client secret
.\Export-CyberArkAccountsToCsv.ps1 `
    -PVWAUrl 'https://mytenant.privilegecloud.cyberark.cloud/PasswordVault' `
    -AuthType OAuth -Credential $oauth `
    -OutputCsv .\accounts.csv
```

### 4. Pre-obtained token / Identity header

Authenticate yourself with `Get-IdentityHeader` (handy when you want full
control of the Identity flow) and pass the header via `-LogonToken`. The session
is **not** logged off when you pass a token.

```powershell
Import-Module '..\Identity Authentication\IdentityAuth.psm1'
$hdr = Get-IdentityHeader -IdentityUserName 'me@corp.com' `
    -PCloudURL 'https://mytenant.privilegecloud.cyberark.cloud/PasswordVault'

.\Export-CyberArkAccountsToCsv.ps1 `
    -PVWAUrl 'https://mytenant.privilegecloud.cyberark.cloud/PasswordVault' `
    -LogonToken $hdr -OutputCsv .\accounts.csv
```

### 5. Classic PVWA logon (self-hosted)

```powershell
$cred = Get-Credential   # vault username + password
.\Export-CyberArkAccountsToCsv.ps1 `
    -PVWAUrl 'https://pvwa.corp.local/PasswordVault' `
    -AuthType CyberArk -Credential $cred `
    -OutputCsv .\accounts.csv
# Add -SkipCertificateValidation for self-signed lab certificates.
```

---

## Step 1 — Export

```powershell
.\Export-CyberArkAccountsToCsv.ps1 `
    -PVWAUrl 'https://mytenant.privilegecloud.cyberark.cloud/PasswordVault' `
    -IdentityUserName 'me@corp.com' `
    -OutputCsv .\accounts.csv
```

Useful options:

- `-SafeName 'WindowsServers'` — export a single safe only.
- `-Search 'admin'` — free-text filter.
- `-FetchAccountDetails Always|Auto|Never` — `Auto` (default) fetches full
  per-account details only when the list payload does not already include the
  custom properties; `Always` forces a detail call per account; `Never` is
  fastest but may omit some custom properties.
- `-PageSize 1000` — accounts per API page.

### CSV columns

| Column group | Examples | Editable? |
|--------------|----------|-----------|
| **Editable top-level fields** | `Name`, `Username`, `Address` | **Yes** — list them in `-PropertiesToUpdate` to change them. |
| **Keys / identity** | `AccountID`, `SafeName`, `PlatformID`, `SecretType` | Identify the account; **never changed** by the import. |
| **Custom File Categories** | `Notes`, `Environment`, `Location`, `OwnerName`, `Port`, ... (one column per property) | **Yes** — edit these. |
| **Metadata (underscore-prefixed)** | `_AutomaticManagementEnabled`, `_ManualManagementReason`, `_CreatedTime`, `_CategoryModificationTime` | Read-only / informational; ignored by the import. |

- The CSV contains the **union** of every custom property across all accounts;
  an account that does not have a given property shows a blank cell for it.
- A custom property whose name collides with a top-level field (rare) is
  prefixed with `FC_` (e.g. `FC_Name`).

---

## Step 2 — Edit the CSV

Open the CSV in Excel / a text editor and edit only the columns you intend to
change — for example `Address` and/or `Notes`.

- **Keep the `AccountID` column** — it is the authoritative key.
- Leave a cell blank to make **no change** to that property (default behaviour).
- You do not need to keep columns you are not updating, but it does no harm to.

---

## Step 3 — Import (dry-run first!)

Always preview with `-WhatIf` (or `-DryRun`) before a real run.

### Dry-run

```powershell
# Update both the Address field and the Notes File Category in one run
.\Import-CyberArkAccountPropertyUpdates.ps1 `
    -PVWAUrl 'https://mytenant.privilegecloud.cyberark.cloud/PasswordVault' `
    -IdentityUserName 'me@corp.com' `
    -InputCsv .\accounts.csv `
    -PropertiesToUpdate Address,Notes `
    -WhatIf
```

### Real run

```powershell
$oauth = Get-Credential   # OAuth client ID + secret (or use -IdentityUserName)
.\Import-CyberArkAccountPropertyUpdates.ps1 `
    -PVWAUrl 'https://mytenant.privilegecloud.cyberark.cloud/PasswordVault' `
    -AuthType OAuth -Credential $oauth `
    -InputCsv .\accounts.csv `
    -PropertiesToUpdate Address,Notes `
    -LogPath .\import.log
```

Behaviour:

- **`-PropertiesToUpdate`** lists exactly which columns/properties may change.
  Everything else in the CSV is ignored. You can mix editable top-level fields
  (`Address`, `Username`, `Name`) and custom File Categories (e.g. `Notes`) in
  the same run.
- **Top-level fields** (`Address`/`Username`/`Name`) are replaced when the value
  differs; a blank cell is ignored (they are never cleared).
- **Custom File Categories**: a **new** value is added, a value that **differs**
  is replaced, an **unchanged** value is skipped. Blank cells are left untouched
  by default; use **`-RemoveEmptyValues`** to remove a File Category when blank.
- `AccountID`, `SafeName`, `PlatformID` and `SecretType` are keys/identity and
  are never modified (listing one just logs a warning).
- Each account is updated with a **single PATCH** containing all its changes.
- Rows are independent: if one row fails, the script logs it and continues.
- Use **`-ThrottleMs 200`** to pace requests if you hit rate limits (HTTP 429).
  The helper also retries 429/5xx/network errors with exponential back-off.

---

## Logging & results

- **`-LogPath`** — timestamped text log (mirrors console output). Used by both
  scripts; defaults to `.\CyberArk-AccountProperties-<timestamp>.log`.
- The import additionally writes a **per-row results CSV** next to the log
  (`<logname>.results.csv`) and returns the same result objects to the pipeline.

Per-row result columns: `Row`, `AccountID`, `Name`, `SafeName`,
`ChangedFields`, `Status`, `Error`.

Possible `Status` values:

| Status | Meaning |
|--------|---------|
| `Updated` | Account was patched. |
| `WouldUpdate` | Dry-run/`-WhatIf`: change detected but not applied. |
| `NoChange` | Desired values already match the account. |
| `NotFound` | No account exists for that AccountID (HTTP 404). |
| `Skipped` | Missing AccountID (or skipped at a confirmation prompt). |
| `Error` | Update failed; see `Error` column / log. |

---

## Parameter quick reference

Common to both scripts:

| Parameter | Description |
|-----------|-------------|
| `-PVWAUrl` | Privilege Cloud / PVWA base URL. |
| `-AuthType` | `Identity` (default), `OAuth`, `CyberArk`, `LDAP`, `RADIUS`. |
| `-Credential` | Auth credential (Identity user/pass, OAuth client ID/secret, or vault user/pass). |
| `-IdentityUserName` | Identity username for interactive (MFA) login via IdentityAuth.psm1. |
| `-LogonToken` | Pre-obtained header hashtable or token string. |
| `-IdentityTenantURL` | Identity tenant URL passed to IdentityAuth.psm1 (else auto-discovered). |
| `-IdentityAuthModulePath` | Explicit path to IdentityAuth.psm1. |
| `-DownloadIdentityAuth` | Fetch IdentityAuth.psm1 from the repo if not found locally. |
| `-LogPath` | Text log file path. |
| `-SkipCertificateValidation` | Ignore TLS cert errors (self-hosted labs). |

Export-only: `-OutputCsv`, `-SafeName`, `-Search`, `-PageSize`,
`-FetchAccountDetails`, `-Delimiter`, `-ConcurrentSession`.

Import-only: `-InputCsv`, `-PropertiesToUpdate`, `-AccountIDColumn`,
`-RemoveEmptyValues`, `-WhatIf` / `-DryRun`, `-ThrottleMs`, `-Delimiter`,
`-ConcurrentSession`.

---

## How it maps to the CyberArk REST API

| Step | API call |
|------|----------|
| Authenticate (Identity / OAuth) | `IdentityAuth.psm1` `Get-IdentityHeader` (CyberArk Identity → Bearer token) |
| Authenticate (classic) | `POST {PVWAUrl}/api/auth/{AuthType}/Logon` |
| List accounts | `GET {PVWAUrl}/api/Accounts?limit=&offset=` (paginated) |
| Account details | `GET {PVWAUrl}/api/Accounts/{id}` |
| Update top-level field | `PATCH {PVWAUrl}/api/Accounts/{id}` — e.g. `[{ "op":"replace", "path":"/address", "value":"..." }]` |
| Update File Category | `PATCH {PVWAUrl}/api/Accounts/{id}` — e.g. `[{ "op":"replace", "path":"/platformAccountProperties/Notes", "value":"..." }]` |

See the [CyberArk Privilege Cloud ISPSS REST API Cookbook](../.REST%20API%20Cookbooks/CyberArk%20Privilege%20Cloud%20ISPSS%20REST%20API%20Cookbook)
in this repository for more API examples.

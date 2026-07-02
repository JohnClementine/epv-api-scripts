<#
.SYNOPSIS
    Updates custom File Categories / account properties of existing CyberArk
    accounts from a (hand-edited) CSV, keyed by the CyberArk account ID.

.DESCRIPTION
    Reads a CSV produced by Export-CyberArkAccountsToCsv.ps1 (or any CSV that
    contains an AccountID column plus the property columns you want to update),
    then for each row:

      1. Looks up the existing account by its CyberArk account ID
         (the authoritative key - the account ID identifies the row, not Name/Safe).
      2. Compares the desired value of each property named in
         -PropertiesToUpdate against the current account value.
      3. Issues a single PATCH containing only the changed fields.

    Updatable properties are:
      * Editable top-level fields: Address (/address), Username (/userName),
        Name (/name), PlatformID (/platformId) - replaced when the value differs.
      * Custom File Categories (platformAccountProperties), e.g. Notes - added
        when new, replaced when changed, and removed when -RemoveEmptyValues is
        set and the cell is blank.
    Both kinds can be updated in the same run (e.g. -PropertiesToUpdate Address,Notes).

    NOTE on PlatformID: changing an account's platform is consequential. It only
    succeeds for compatible platforms, may require the new platform's mandatory
    properties to be set in the same run, and does NOT move the account to a
    different safe. Always preview with -WhatIf and pilot on a few accounts first.

    AccountID, SafeName and SecretType are treated as keys/identity and are never
    modified. The script never changes the secret, never reads passwords, and
    never triggers password management. Notes is handled as an ordinary custom
    File Category.

    Supports -WhatIf / -DryRun, logs one result line per row, and continues
    processing when an individual row fails.

.PARAMETER PVWAUrl
    Privilege Cloud / PVWA base URL (see Export script for examples).

.PARAMETER InputCsv
    Path to the CSV to import. Must contain an account-ID column (see
    -AccountIDColumn) and one column per property in -PropertiesToUpdate.

.PARAMETER PropertiesToUpdate
    Names of the properties to update, e.g. -PropertiesToUpdate Address,Notes.
    Each must match a CSV column header. Address/Username/Name update the
    matching top-level account field; any other name is applied as a custom
    File Category (platformAccountProperties).

.PARAMETER Credential
    Authentication credential (see Export script).

.PARAMETER LogonToken
    Pre-obtained authorization token / Identity header (see Export script).

.PARAMETER AuthType
    Identity (default), OAuth, CyberArk, LDAP or RADIUS. Identity and OAuth both
    authenticate through the repo's IdentityAuth.psm1. Ignored with -LogonToken.

.PARAMETER IdentityUserName
    Identity username for an interactive (username + MFA) login via
    IdentityAuth.psm1. Used with -AuthType Identity when no -Credential is given.

.PARAMETER IdentityTenantURL
    Optional CyberArk Identity tenant URL passed through to IdentityAuth.psm1
    (auto-discovered by the module if omitted).

.PARAMETER IdentityAuthModulePath
    Optional explicit path to IdentityAuth.psm1 (defaults to the sibling
    "Identity Authentication\IdentityAuth.psm1" in the repo).

.PARAMETER DownloadIdentityAuth
    Download IdentityAuth.psm1 from the epv-api-scripts repo if not found locally.

.PARAMETER LogPath
    Text log file path. A per-row results CSV is written alongside it
    (<logname>.results.csv). Defaults are timestamped in the current directory.

.PARAMETER AccountIDColumn
    CSV column holding the CyberArk account ID. Defaults to AccountID (falls
    back to 'id' if AccountID is not present).

.PARAMETER RemoveEmptyValues
    When set, a blank cell for a managed custom File Category removes that
    property from the account. By default, blank cells are left untouched (no
    accidental clearing). Top-level fields (Address/Username/Name) are never
    cleared by a blank cell.

.PARAMETER DryRun
    Preview changes without writing them. Equivalent to -WhatIf.

.PARAMETER ConcurrentSession
    Classic logon only: allow a concurrent vault session.

.PARAMETER SkipCertificateValidation
    Ignore TLS certificate errors (self-hosted labs with self-signed certs).

.PARAMETER Delimiter
    CSV delimiter. Defaults to comma.

.PARAMETER ThrottleMs
    Optional millisecond pause between rows (helps avoid rate limiting).

.EXAMPLE
    # Dry run - update both Address and Notes in one run (interactive Identity login)
    .\Import-CyberArkAccountPropertyUpdates.ps1 `
        -PVWAUrl 'https://mytenant.privilegecloud.cyberark.cloud/PasswordVault' `
        -IdentityUserName 'me@corp.com' `
        -InputCsv .\accounts-edited.csv `
        -PropertiesToUpdate Address,Notes `
        -WhatIf

.EXAMPLE
    # Real run - OAuth client credentials (headless), via IdentityAuth.psm1
    $oauth = Get-Credential   # Username = OAuth client ID, Password = client secret
    .\Import-CyberArkAccountPropertyUpdates.ps1 `
        -PVWAUrl 'https://mytenant.privilegecloud.cyberark.cloud/PasswordVault' `
        -AuthType OAuth -Credential $oauth `
        -InputCsv .\accounts-edited.csv `
        -PropertiesToUpdate Address,Notes `
        -LogPath .\import.log
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory)]
    [Alias('url', 'PCloudURL')]
    [string]$PVWAUrl,

    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [Alias('csv', 'InputFile')]
    [string]$InputCsv,

    [Parameter(Mandatory)]
    [string[]]$PropertiesToUpdate,

    [Parameter()]
    [System.Management.Automation.PSCredential]$Credential,

    [Parameter()]
    [AllowNull()]
    $LogonToken,

    [Parameter()]
    [ValidateSet('Identity', 'OAuth', 'CyberArk', 'LDAP', 'RADIUS')]
    [string]$AuthType = 'Identity',

    [Parameter()]
    [string]$IdentityUserName,

    [Parameter()]
    [string]$IdentityTenantURL,

    [Parameter()]
    [string]$IdentityAuthModulePath,

    [Parameter()]
    [switch]$DownloadIdentityAuth,

    [Parameter()]
    [string]$LogPath,

    [Parameter()]
    [string]$AccountIDColumn = 'AccountID',

    [Parameter()]
    [switch]$RemoveEmptyValues,

    [Parameter()]
    [switch]$DryRun,

    [Parameter()]
    [switch]$ConcurrentSession,

    [Parameter()]
    [switch]$SkipCertificateValidation,

    [Parameter()]
    [string]$Delimiter,

    [Parameter()]
    [int]$ThrottleMs = 0
)

$ErrorActionPreference = 'Stop'

# -DryRun is a friendly alias for -WhatIf
if ($DryRun) { $WhatIfPreference = $true }

# Import the shared helper module (sits next to this script)
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Import-Module (Join-Path $scriptRoot 'CyberArkAccountProperties.psm1') -Force

$resolvedLog = Initialize-CALog -LogPath $LogPath
$resultsCsv  = [System.IO.Path]::ChangeExtension($resolvedLog, '.results.csv')

Write-CALog -Type Header -Message "=== Import-CyberArkAccountPropertyUpdates started ==="
Write-CALog -Type Info   -Message "PVWA URL            : $PVWAUrl"
Write-CALog -Type Info   -Message "Input CSV           : $InputCsv"
Write-CALog -Type Info   -Message "Properties to update: $($PropertiesToUpdate -join ', ')"
Write-CALog -Type Info   -Message "Remove empty values : $([bool]$RemoveEmptyValues)"
Write-CALog -Type Info   -Message "Log file            : $resolvedLog"
Write-CALog -Type Info   -Message "Results CSV         : $resultsCsv"
if ($WhatIfPreference) {
    Write-CALog -Type Warning -Message "DRY-RUN / -WhatIf mode: no changes will be written to CyberArk."
}

$session = $null
try {
    # ----- Read and validate the CSV --------------------------------------
    $importParams = @{ LiteralPath = $InputCsv }
    if (-not [string]::IsNullOrEmpty($Delimiter)) { $importParams.Delimiter = $Delimiter }
    $rows = @(Import-Csv @importParams)
    Write-CALog -Type Info -Message "Loaded $($rows.Count) row(s) from the CSV."

    if ($rows.Count -eq 0) {
        Write-CALog -Type Warning -Message "The CSV contains no data rows - nothing to do."
        return
    }

    $headers = $rows[0].PSObject.Properties.Name

    if ($headers -notcontains $AccountIDColumn) {
        if ($headers -contains 'id') {
            Write-CALog -Type Warning -Message "Column '$AccountIDColumn' not found; using 'id' as the account ID column."
            $AccountIDColumn = 'id'
        } else {
            throw "Input CSV is missing the account ID column '$AccountIDColumn' (and no 'id' column was found)."
        }
    }

    # Reject identity/key columns and read-only metadata. Updatable entries are
    # the editable top-level fields (Address/Username/Name) plus any custom File
    # Category (platformAccountProperties); the resolution happens per-property
    # in Get-CAPropertyUpdateOperation.
    $protectedColumns = Get-CAProtectedColumns
    $protectedProps   = @($PropertiesToUpdate | Where-Object { ($protectedColumns -contains $_) -or ($_ -like '_*') -or ($_ -ieq $AccountIDColumn) })
    if ($protectedProps.Count -gt 0) {
        Write-CALog -Type Warning -Message "These -PropertiesToUpdate entries are identity/key/metadata columns and will NOT be updated: $($protectedProps -join ', ')"
    }

    $candidateProps = @($PropertiesToUpdate | Where-Object { -not (($protectedColumns -contains $_) -or ($_ -like '_*') -or ($_ -ieq $AccountIDColumn)) })
    $missingProps   = @($candidateProps | Where-Object { $headers -notcontains $_ })
    $effectiveProps = @($candidateProps | Where-Object { $headers -contains $_ })
    if ($missingProps.Count -gt 0) {
        Write-CALog -Type Warning -Message "These -PropertiesToUpdate columns are not in the CSV and will be skipped: $($missingProps -join ', ')"
    }
    if ($effectiveProps.Count -eq 0) {
        throw "None of the -PropertiesToUpdate columns are updatable properties present in the CSV. Nothing to update."
    }
    Write-CALog -Type Info -Message "Effective properties: $($effectiveProps -join ', ')"

    # ----- Authenticate ---------------------------------------------------
    if ($null -eq $LogonToken) {
        if ($AuthType -eq 'Identity' -and -not $Credential -and [string]::IsNullOrEmpty($IdentityUserName)) {
            $IdentityUserName = Read-Host -Prompt "Enter your CyberArk Identity username (e.g. you@corp.com)"
        } elseif ($AuthType -ne 'Identity' -and -not $Credential) {
            $Credential = Get-Credential -Message "Enter CyberArk credentials (OAuth client ID/secret, or vault username/password)"
        }
    }

    $connect = @{ PVWAUrl = $PVWAUrl; AuthType = $AuthType }
    if ($Credential)                                          { $connect.Credential = $Credential }
    if ($null -ne $LogonToken)                                { $connect.LogonToken = $LogonToken }
    if (-not [string]::IsNullOrEmpty($IdentityUserName))      { $connect.IdentityUserName = $IdentityUserName }
    if (-not [string]::IsNullOrEmpty($IdentityTenantURL))     { $connect.IdentityTenantURL = $IdentityTenantURL }
    if (-not [string]::IsNullOrEmpty($IdentityAuthModulePath)) { $connect.IdentityAuthModulePath = $IdentityAuthModulePath }
    if ($DownloadIdentityAuth)                                { $connect.DownloadIdentityAuth = $true }
    if ($ConcurrentSession)                                   { $connect.ConcurrentSession = $true }
    if ($SkipCertificateValidation)                           { $connect.SkipCertificateValidation = $true }

    $session = New-CASession @connect

    # ----- Process every row ----------------------------------------------
    $results = New-Object System.Collections.Generic.List[object]
    $counts  = [ordered]@{ Updated = 0; WouldUpdate = 0; NoChange = 0; NotFound = 0; Error = 0; Skipped = 0 }
    $rowNum  = 0

    foreach ($row in $rows) {
        $rowNum++
        $accountId   = ([string]$row.$AccountIDColumn).Trim()
        $name        = ''
        $safe        = ''
        $status      = ''
        $errorText   = ''
        $changedText = ''

        if ([string]::IsNullOrEmpty($accountId)) {
            $status = 'Skipped'; $errorText = 'Missing AccountID'; $counts.Skipped++
            Write-CALog -Type Warning -Message "Row ${rowNum}: skipped - missing AccountID."
            $results.Add([pscustomobject]@{ Row = $rowNum; AccountID = ''; Name = ''; SafeName = ''; ChangedFields = ''; Status = $status; Error = $errorText })
            continue
        }

        try {
            $account = Get-CAAccountDetail -Session $session -AccountID $accountId
            $name = [string]$account.name
            $safe = [string]$account.safeName

            $updates = @{}
            foreach ($prop in $effectiveProps) { $updates[$prop] = [string]$row.$prop }

            $plan = Get-CAPropertyUpdateOperation -Account $account -Updates $updates -RemoveEmptyValues:$RemoveEmptyValues

            if ($plan.Operations.Count -eq 0) {
                $status = 'NoChange'; $counts.NoChange++
                Write-CALog -Type Info -Message "Row ${rowNum} [$accountId] $name ($safe): no changes needed."
            } else {
                $changedText = ($plan.ChangedFields -join ' | ')
                $target = "Account $accountId ($name in safe '$safe')"
                $action = "Update account fields: $changedText"

                # Platform changes are consequential - make them stand out in the log.
                $platformOp = @($plan.Operations | Where-Object { $_.path -eq '/platformId' })
                if ($platformOp.Count -gt 0) {
                    Write-CALog -Type Warning -Message "Row ${rowNum} [$accountId] $name ($safe): PLATFORM CHANGE -> '$($platformOp[0].value)'. Verify the account is manageable on the new platform afterwards."
                }

                # $WhatIfPreference is $true for both -WhatIf and -DryRun. Check it
                # first so a dry-run can never issue a PATCH, regardless of how
                # ShouldProcess resolves the preference.
                if ($WhatIfPreference) {
                    $status = 'WouldUpdate'; $counts.WouldUpdate++
                    Write-CALog -Type Info -Message "Row ${rowNum} [$accountId] $name ($safe): WOULD update -> $changedText"
                } elseif ($PSCmdlet.ShouldProcess($target, $action)) {
                    $jsonBody = ConvertTo-CAJsonArray -Items $plan.Operations
                    $null = Invoke-CARest -Session $session -Method Patch -Uri "$($session.ApiBase)/Accounts/$accountId" -Body $jsonBody
                    $status = 'Updated'; $counts.Updated++
                    Write-CALog -Type Success -Message "Row ${rowNum} [$accountId] $name ($safe): updated -> $changedText"
                } else {
                    $status = 'Skipped'; $counts.Skipped++
                    Write-CALog -Type Warning -Message "Row ${rowNum} [$accountId] $name ($safe): skipped by user (declined at confirmation)."
                }
            }
        } catch {
            $statusCode = $null
            try { $statusCode = $_.Exception.Data['StatusCode'] } catch { }
            if ($statusCode -eq 404) {
                $status = 'NotFound'; $counts.NotFound++
                Write-CALog -Type Error -Message "Row ${rowNum} [$accountId]: account not found (HTTP 404)."
            } else {
                $status = 'Error'; $counts.Error++
                Write-CALog -Type Error -Message "Row ${rowNum} [$accountId] $name ($safe): $($_.Exception.Message)"
            }
            $errorText = $_.Exception.Message
        }

        $results.Add([pscustomobject]@{
                Row           = $rowNum
                AccountID     = $accountId
                Name          = $name
                SafeName      = $safe
                ChangedFields = $changedText
                Status        = $status
                Error         = $errorText
            })

        if ($ThrottleMs -gt 0) { Start-Sleep -Milliseconds $ThrottleMs }
    }

    # ----- Write results + summary ----------------------------------------
    try {
        $results | Export-Csv -LiteralPath $resultsCsv -NoTypeInformation -Encoding UTF8 -Force
        Write-CALog -Type Info -Message "Per-row results written to: $resultsCsv"
    } catch {
        Write-CALog -Type Warning -Message "Could not write results CSV: $($_.Exception.Message)"
    }

    $summaryValues = @($counts.Updated, $counts.WouldUpdate, $counts.NoChange, $counts.NotFound, $counts.Error, $counts.Skipped, $rows.Count)
    $summary = "Summary: Updated={0} WouldUpdate={1} NoChange={2} NotFound={3} Error={4} Skipped={5} (Total={6})" -f $summaryValues
    Write-CALog -Type Header -Message $summary
    Write-CALog -Type Header -Message "=== Import complete ==="

    # Emit results to the pipeline for further processing
    $results
} catch {
    Write-CALog -Type Error -Message "Import failed: $($_.Exception.Message)"
    throw
} finally {
    if ($null -ne $session) {
        Close-CASession -Session $session
    }
}

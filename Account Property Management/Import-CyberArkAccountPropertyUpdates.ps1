<#
.SYNOPSIS
    Updates custom File Categories / account properties of existing CyberArk
    accounts from a (hand-edited) CSV, keyed by the CyberArk account ID.

.DESCRIPTION
    Reads a CSV produced by Export-CyberArkAccountsToCsv.ps1 (or any CSV that
    contains an AccountID column plus the property columns you want to update),
    then for each row:

      1. Looks up the existing account by its CyberArk account ID
         (the authoritative key - Name/Safe/etc. are informational only).
      2. Compares the desired value of each property named in
         -PropertiesToUpdate against the current account value.
      3. Issues a single PATCH containing only the changed File Categories
         (add for new, replace for changed; remove when -RemoveEmptyValues
         is set and the cell is blank).

    Only /platformAccountProperties/* paths are modified. The script never
    changes the secret, never reads passwords, and never triggers password
    management. Notes is handled as an ordinary custom File Category.

    Supports -WhatIf / -DryRun, logs one result line per row, and continues
    processing when an individual row fails.

.PARAMETER PVWAUrl
    Privilege Cloud / PVWA base URL (see Export script for examples).

.PARAMETER InputCsv
    Path to the CSV to import. Must contain an account-ID column (see
    -AccountIDColumn) and one column per property in -PropertiesToUpdate.

.PARAMETER PropertiesToUpdate
    Names of the custom File Categories / account properties to update,
    e.g. -PropertiesToUpdate Notes,CustomField1. These must match CSV column
    headers and are applied as platformAccountProperties.

.PARAMETER Credential
    Authentication credential (see Export script).

.PARAMETER LogonToken
    Pre-obtained authorization token / Identity header (see Export script).

.PARAMETER AuthType
    OAuth (default), CyberArk, LDAP or RADIUS. Ignored when -LogonToken is used.

.PARAMETER IdentityTenantURL
    Optional CyberArk Identity tenant URL (OAuth only; auto-discovered if omitted).

.PARAMETER LogPath
    Text log file path. A per-row results CSV is written alongside it
    (<logname>.results.csv). Defaults are timestamped in the current directory.

.PARAMETER AccountIDColumn
    CSV column holding the CyberArk account ID. Defaults to AccountID (falls
    back to 'id' if AccountID is not present).

.PARAMETER RemoveEmptyValues
    When set, a blank cell for a managed property removes that property from the
    account. By default, blank cells are left untouched (no accidental clearing).

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
    # Dry run - show exactly what would change, write nothing
    $cred = Get-Credential
    .\Import-CyberArkAccountPropertyUpdates.ps1 `
        -PVWAUrl 'https://mytenant.privilegecloud.cyberark.cloud/PasswordVault' `
        -Credential $cred `
        -InputCsv .\accounts-edited.csv `
        -PropertiesToUpdate Notes,Environment `
        -WhatIf

.EXAMPLE
    # Real run
    .\Import-CyberArkAccountPropertyUpdates.ps1 `
        -PVWAUrl 'https://mytenant.privilegecloud.cyberark.cloud/PasswordVault' `
        -Credential $cred `
        -InputCsv .\accounts-edited.csv `
        -PropertiesToUpdate Notes,Environment `
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
    [ValidateSet('OAuth', 'CyberArk', 'LDAP', 'RADIUS')]
    [string]$AuthType = 'OAuth',

    [Parameter()]
    [string]$IdentityTenantURL,

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

    # Reject base/identity columns and read-only metadata - updates are limited
    # to custom File Categories (platformAccountProperties).
    $baseColumns     = Get-CABaseColumns
    $protectedProps  = @($PropertiesToUpdate | Where-Object { ($baseColumns -contains $_) -or ($_ -like '_*') -or ($_ -ieq $AccountIDColumn) })
    if ($protectedProps.Count -gt 0) {
        Write-CALog -Type Warning -Message "These -PropertiesToUpdate entries are identity/metadata columns and will NOT be updated: $($protectedProps -join ', ')"
    }

    $candidateProps = @($PropertiesToUpdate | Where-Object { -not (($baseColumns -contains $_) -or ($_ -like '_*') -or ($_ -ieq $AccountIDColumn)) })
    $missingProps   = @($candidateProps | Where-Object { $headers -notcontains $_ })
    $effectiveProps = @($candidateProps | Where-Object { $headers -contains $_ })
    if ($missingProps.Count -gt 0) {
        Write-CALog -Type Warning -Message "These -PropertiesToUpdate columns are not in the CSV and will be skipped: $($missingProps -join ', ')"
    }
    if ($effectiveProps.Count -eq 0) {
        throw "None of the -PropertiesToUpdate columns are updatable custom properties present in the CSV. Nothing to update."
    }
    Write-CALog -Type Info -Message "Effective properties: $($effectiveProps -join ', ')"

    # ----- Authenticate ---------------------------------------------------
    if (-not $Credential -and ($null -eq $LogonToken)) {
        $Credential = Get-Credential -Message "Enter CyberArk credentials (OAuth client ID/secret, or vault username/password)"
    }

    $connect = @{ PVWAUrl = $PVWAUrl; AuthType = $AuthType }
    if ($Credential)                          { $connect.Credential = $Credential }
    if ($null -ne $LogonToken)                { $connect.LogonToken = $LogonToken }
    if (-not [string]::IsNullOrEmpty($IdentityTenantURL)) { $connect.IdentityTenantURL = $IdentityTenantURL }
    if ($ConcurrentSession)                   { $connect.ConcurrentSession = $true }
    if ($SkipCertificateValidation)           { $connect.SkipCertificateValidation = $true }

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
                $action = "Update File Categories: $changedText"

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

<#
.SYNOPSIS
    Exports all CyberArk Privilege Cloud (ISPSS) / PVWA accounts, including
    every custom File Category (platformAccountProperties), to a CSV file.

.DESCRIPTION
    Authenticates to CyberArk, lists all accounts using pagination, optionally
    retrieves full per-account details so that all custom File Categories are
    captured, flattens each account into a row, and writes a single CSV whose
    columns are the union of all properties found across every account.

    The CSV is designed to be hand-edited (to add/update custom property
    values such as Notes) and then fed to Import-CyberArkAccountPropertyUpdates.ps1.

    Notes is treated as an ordinary custom File Category / account property,
    NOT as a special password field.

    This script only reads account metadata. It never retrieves passwords and
    never triggers password management actions.

.PARAMETER PVWAUrl
    Privilege Cloud / PVWA base URL.
    ISPSS example:  https://<subdomain>.privilegecloud.cyberark.cloud/PasswordVault
    Self-hosted:    https://<host>/PasswordVault

.PARAMETER Credential
    Credential used to authenticate.
      * AuthType OAuth (default, ISPSS): username = OAuth client ID,
        password = OAuth client secret.
      * AuthType CyberArk/LDAP/RADIUS: vault username and password.
    If neither -Credential nor -LogonToken is supplied you will be prompted.

.PARAMETER LogonToken
    A pre-obtained authorization token. Accepts the header hashtable returned
    by the Identity Authentication module's Get-IdentityHeader (recommended for
    interactive / MFA logins) or a raw token string. When supplied, the session
    is NOT logged off.

.PARAMETER AuthType
    OAuth (default), CyberArk, LDAP or RADIUS. Ignored when -LogonToken is used.

.PARAMETER IdentityTenantURL
    Optional CyberArk Identity tenant URL (e.g. https://abc1234.id.cyberark.cloud).
    Auto-discovered from PVWAUrl when omitted (OAuth only).

.PARAMETER OutputCsv
    Destination CSV path. Defaults to .\CyberArkAccounts-<timestamp>.csv.

.PARAMETER LogPath
    Text log file path. Defaults to .\CyberArk-AccountProperties-<timestamp>.log.

.PARAMETER SafeName
    Optional: export only accounts in this safe.

.PARAMETER Search
    Optional: free-text search filter passed to the API.

.PARAMETER PageSize
    Accounts per API page (1-1000, default 1000).

.PARAMETER FetchAccountDetails
    Auto (default) retrieves full details per account only when the list payload
    does not already include platformAccountProperties. Always forces a detail
    call for every account. Never uses only the list payload (fastest).

.PARAMETER ConcurrentSession
    Classic logon only: allow a concurrent vault session.

.PARAMETER SkipCertificateValidation
    Ignore TLS certificate errors (self-hosted labs with self-signed certs).

.PARAMETER Delimiter
    CSV delimiter. Defaults to comma (round-trips cleanly with the importer).

.EXAMPLE
    $cred = Get-Credential   # OAuth client ID + secret
    .\Export-CyberArkAccountsToCsv.ps1 `
        -PVWAUrl 'https://mytenant.privilegecloud.cyberark.cloud/PasswordVault' `
        -Credential $cred `
        -OutputCsv .\accounts.csv

.EXAMPLE
    # Interactive / MFA login using the Identity Authentication module
    Import-Module '..\Identity Authentication\IdentityAuth.psm1'
    $hdr = Get-IdentityHeader -IdentityUserName 'me@corp.com' -PCloudURL 'https://mytenant.privilegecloud.cyberark.cloud/PasswordVault'
    .\Export-CyberArkAccountsToCsv.ps1 -PVWAUrl 'https://mytenant.privilegecloud.cyberark.cloud/PasswordVault' -LogonToken $hdr -OutputCsv .\accounts.csv
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [Alias('url', 'PCloudURL')]
    [string]$PVWAUrl,

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
    [Alias('path', 'OutPath')]
    [string]$OutputCsv,

    [Parameter()]
    [string]$LogPath,

    [Parameter()]
    [Alias('Safe')]
    [string]$SafeName,

    [Parameter()]
    [string]$Search,

    [Parameter()]
    [ValidateRange(1, 1000)]
    [int]$PageSize = 1000,

    [Parameter()]
    [ValidateSet('Auto', 'Always', 'Never')]
    [string]$FetchAccountDetails = 'Auto',

    [Parameter()]
    [switch]$ConcurrentSession,

    [Parameter()]
    [switch]$SkipCertificateValidation,

    [Parameter()]
    [string]$Delimiter
)

$ErrorActionPreference = 'Stop'

# Import the shared helper module (sits next to this script)
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Import-Module (Join-Path $scriptRoot 'CyberArkAccountProperties.psm1') -Force

if ([string]::IsNullOrEmpty($OutputCsv)) {
    $OutputCsv = Join-Path -Path (Get-Location).Path -ChildPath ("CyberArkAccounts-{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
}

$resolvedLog = Initialize-CALog -LogPath $LogPath
Write-CALog -Type Header  -Message "=== Export-CyberArkAccountsToCsv started ==="
Write-CALog -Type Info    -Message "PVWA URL           : $PVWAUrl"
Write-CALog -Type Info    -Message "Output CSV         : $OutputCsv"
Write-CALog -Type Info    -Message "Log file           : $resolvedLog"
Write-CALog -Type Info    -Message "Fetch detail mode  : $FetchAccountDetails"

$session = $null
try {
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

    # ----- List accounts (paginated) --------------------------------------
    Write-CALog -Type Info -Message "Retrieving account list..."
    $list = @(Get-CAAccountList -Session $session -SafeName $SafeName -Search $Search -PageSize $PageSize)
    Write-CALog -Type Success -Message "Total accounts retrieved: $($list.Count)"

    if ($list.Count -eq 0) {
        Write-CALog -Type Warning -Message "No accounts found - nothing to export."
        return
    }

    # ----- Flatten (fetching per-account detail as needed) ----------------
    $baseColumns = Get-CABaseColumns
    $metaColumns = Get-CAMetadataColumns

    $records       = New-Object System.Collections.Generic.List[object]
    $customColumns = [ordered]@{}   # lower-case name -> canonical name
    $detailFetched = 0
    $index         = 0
    $total         = $list.Count

    foreach ($account in $list) {
        $index++
        $full      = $account
        $needDetail = $false

        if ($FetchAccountDetails -eq 'Always') {
            $needDetail = $true
        } elseif ($FetchAccountDetails -eq 'Auto') {
            $hasProps = ($account.PSObject.Properties['platformAccountProperties'] -and $account.platformAccountProperties)
            if (-not $hasProps) { $needDetail = $true }
        }

        if ($needDetail) {
            try {
                $full = Get-CAAccountDetail -Session $session -AccountID $account.id
                $detailFetched++
            } catch {
                Write-CALog -Type Warning -Message "Could not fetch details for $($account.id) ($($account.name)): $($_.Exception.Message). Using list data."
                $full = $account
            }
        }

        $record = ConvertTo-CAFlatRecord -Account $full

        foreach ($key in $record.Keys) {
            if ($baseColumns -contains $key) { continue }
            if ($key -like '_*')             { continue }
            $lower = $key.ToLower()
            if (-not $customColumns.Contains($lower)) {
                $customColumns[$lower] = $key
            }
        }

        $records.Add($record)

        if ($index % 50 -eq 0 -or $index -eq $total) {
            Write-CALog -Type Info -Message "Processed $index / $total accounts..."
        }
    }

    # ----- Build the unified column layout --------------------------------
    $sortedCustom = @($customColumns.Values | Sort-Object)
    $allColumns   = @()
    $allColumns  += $baseColumns
    $allColumns  += $sortedCustom
    $allColumns  += $metaColumns

    # ----- Normalise every record to the full column set ------------------
    $objects = New-Object System.Collections.Generic.List[object]
    foreach ($record in $records) {
        $ordered = [ordered]@{}
        foreach ($column in $allColumns) {
            if ($record.Contains($column)) {
                $ordered[$column] = $record[$column]
            } else {
                $ordered[$column] = ''
            }
        }
        $objects.Add([pscustomobject]$ordered)
    }

    # ----- Write the CSV --------------------------------------------------
    $exportParams = @{
        Path              = $OutputCsv
        NoTypeInformation = $true
        Encoding          = 'UTF8'
        Force             = $true
    }
    if (-not [string]::IsNullOrEmpty($Delimiter)) {
        $exportParams.Delimiter = $Delimiter
    }
    $objects | Export-Csv @exportParams

    Write-CALog -Type Success -Message "Exported $($objects.Count) accounts to $OutputCsv"
    Write-CALog -Type Info    -Message "Per-account detail calls made: $detailFetched"
    if ($sortedCustom.Count -gt 0) {
        Write-CALog -Type Info -Message ("Custom File Category columns ({0}): {1}" -f $sortedCustom.Count, ($sortedCustom -join ', '))
    } else {
        Write-CALog -Type Warning -Message "No custom File Category columns were found across the exported accounts."
    }
    Write-CALog -Type Header -Message "=== Export complete ==="
} catch {
    Write-CALog -Type Error -Message "Export failed: $($_.Exception.Message)"
    throw
} finally {
    if ($null -ne $session) {
        Close-CASession -Session $session
    }
}

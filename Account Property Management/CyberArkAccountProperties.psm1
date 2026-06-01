<#
.SYNOPSIS
    Shared helper functions for exporting and importing CyberArk account
    metadata (custom File Categories / platformAccountProperties).

.DESCRIPTION
    This module wraps the CyberArk Privilege Cloud (ISPSS) / self-hosted PVWA
    REST API patterns used by the EPV-API-SCRIPTS repository:

      * Authentication
          - CyberArk Identity OAuth2 client-credentials (ISPSS, headless)
          - A pre-obtained logon token / header (interactive / MFA via the
            Identity Authentication module)
          - Classic PVWA logon (CyberArk / LDAP / RADIUS) for self-hosted
      * Listing accounts with pagination (GET /api/Accounts)
      * Retrieving full account details (GET /api/Accounts/{id})
      * Flattening custom File Categories (platformAccountProperties) to a
        flat record suitable for CSV
      * Building JSON-Patch operations to add/replace/remove custom File
        Categories (PATCH /api/Accounts/{id})

    The module never retrieves passwords and never triggers password
    management actions. Updates are restricted to /platformAccountProperties/*.

    Compatible with Windows PowerShell 5.1 and PowerShell 7+.
#>

# Prefer TLS 1.2 (required by most CyberArk tenants on older hosts)
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
} catch {
    Write-Verbose "Unable to set TLS 1.2 security protocol: $($_.Exception.Message)"
}

# ---------------------------------------------------------------------------
# Module state
# ---------------------------------------------------------------------------
$script:CALogPath     = $null
$script:CABaseColumns = @('AccountID', 'Name', 'Username', 'Address', 'SafeName', 'PlatformID', 'SecretType')
$script:CAMetaColumns = @('_AutomaticManagementEnabled', '_ManualManagementReason', '_CreatedTime', '_CategoryModificationTime')

# Top-level account fields the import is allowed to update, mapped to their
# JSON-Patch path and the account property used to read the current value.
$script:CAEditableBaseFields = [ordered]@{
    'Address'  = @{ Path = '/address';  Property = 'address' }
    'Username' = @{ Path = '/userName'; Property = 'userName' }
    'Name'     = @{ Path = '/name';     Property = 'name' }
}

# Columns that are identity/keys and must never be updated by the import.
$script:CAProtectedColumns = @('AccountID', 'SafeName', 'PlatformID', 'SecretType')

function Get-CABaseColumns {
    <#.SYNOPSIS Returns the fixed base (identity/locator) column names.#>
    return $script:CABaseColumns
}

function Get-CAMetadataColumns {
    <#.SYNOPSIS Returns the read-only metadata column names (underscore-prefixed).#>
    return $script:CAMetaColumns
}

function Get-CAEditableBaseFieldMap {
    <#.SYNOPSIS Returns the map of updatable top-level fields (column -> path/property).#>
    return $script:CAEditableBaseFields
}

function Get-CAProtectedColumns {
    <#.SYNOPSIS Returns column names that the import must never change.#>
    return $script:CAProtectedColumns
}

function Get-CABaseFieldValue {
    <#.SYNOPSIS Case-insensitively reads a top-level account field value (or $null).#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Account,

        [Parameter(Mandatory)]
        [string]$Property
    )
    foreach ($prop in $Account.PSObject.Properties) {
        if ($prop.Name -ieq $Property) { return $prop.Value }
    }
    return $null
}

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
function Initialize-CALog {
    <#
    .SYNOPSIS
        Initialise the log file used by Write-CALog.
    .PARAMETER LogPath
        Path to the text log file. If empty, a timestamped file is created in
        the current directory.
    .OUTPUTS
        The resolved log file path.
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$LogPath
    )

    if ([string]::IsNullOrEmpty($LogPath)) {
        $LogPath = Join-Path -Path (Get-Location).Path -ChildPath ("CyberArk-AccountProperties-{0}.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
    }

    $dir = Split-Path -Parent $LogPath
    if (-not [string]::IsNullOrEmpty($dir) -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    $script:CALogPath = $LogPath
    return $LogPath
}

function Write-CALog {
    <#
    .SYNOPSIS
        Write a timestamped, coloured message to the console and the log file.
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [ValidateSet('Info', 'Warning', 'Error', 'Success', 'Debug', 'Header')]
        [string]$Type = 'Info',

        [Parameter(Mandatory)]
        [string]$Message
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "{0} [{1}] {2}" -f $timestamp, $Type.ToUpper(), $Message

    switch ($Type) {
        'Error'   { Write-Host $line -ForegroundColor Red }
        'Warning' { Write-Host $line -ForegroundColor Yellow }
        'Success' { Write-Host $line -ForegroundColor Green }
        'Header'  { Write-Host $line -ForegroundColor Cyan }
        'Debug'   { Write-Verbose $line }
        default   { Write-Host $line }
    }

    if (-not [string]::IsNullOrEmpty($script:CALogPath)) {
        try {
            Add-Content -LiteralPath $script:CALogPath -Value $line -ErrorAction Stop
        } catch {
            # Never let logging failures abort processing
            Write-Verbose "Failed to write to log file: $($_.Exception.Message)"
        }
    }
}

# ---------------------------------------------------------------------------
# Value conversion helpers
# ---------------------------------------------------------------------------
function ConvertTo-CAStringValue {
    <#.SYNOPSIS Converts an account property value to a stable string for CSV.#>
    [CmdletBinding()]
    param(
        [Parameter()]
        [AllowNull()]
        $Value
    )

    if ($null -eq $Value) { return '' }
    if ($Value -is [bool]) {
        if ($Value) { return 'true' } else { return 'false' }
    }
    if ($Value -is [System.Array]) { return ($Value -join ';') }
    return [string]$Value
}

function Convert-CAEpoch {
    <#.SYNOPSIS Converts a CyberArk epoch (seconds or milliseconds) to UTC text.#>
    [CmdletBinding()]
    param(
        [Parameter()]
        [AllowNull()]
        $Epoch
    )

    $text = "$Epoch"
    if ([string]::IsNullOrWhiteSpace($text)) { return '' }

    try {
        if ($text.Length -gt 11) {
            $dt = [DateTimeOffset]::FromUnixTimeMilliseconds([long]$text).UtcDateTime
        } else {
            $dt = [DateTimeOffset]::FromUnixTimeSeconds([long]$text).UtcDateTime
        }
        return $dt.ToString('yyyy-MM-dd HH:mm:ss') + 'Z'
    } catch {
        return $text
    }
}

# ---------------------------------------------------------------------------
# URL / certificate helpers
# ---------------------------------------------------------------------------
function Get-CAApiBaseUrl {
    <#.SYNOPSIS Normalises a PVWA/PCloud URL to its /api base.#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$PVWAUrl
    )

    $base = $PVWAUrl.Trim().TrimEnd('/')
    if ($base -match '(?i)/api$') {
        return $base
    }
    return "$base/api"
}

function Set-CACertificateValidation {
    <#
    .SYNOPSIS
        Disables server certificate validation on Windows PowerShell 5.1.
    .DESCRIPTION
        Intended for self-hosted lab environments with self-signed certs.
        On PowerShell 7+ the per-request -SkipCertificateCheck switch is used
        instead (see Invoke-CARest), so this is a no-op there.
    #>
    [CmdletBinding()]
    param(
        [switch]$Skip
    )

    if (-not $Skip) { return }
    if ($PSVersionTable.PSEdition -eq 'Core') { return }

    try {
        if (-not ([System.Management.Automation.PSTypeName]'CACertPolicy').Type) {
            Add-Type -TypeDefinition @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class CACertPolicy : ICertificatePolicy {
    public bool CheckValidationResult(ServicePoint sp, X509Certificate cert, WebRequest req, int problem) {
        return true;
    }
}
"@ -ErrorAction Stop
        }
        [System.Net.ServicePointManager]::CertificatePolicy = New-Object CACertPolicy
        Write-CALog -Type Warning -Message "Server certificate validation is DISABLED for this session."
    } catch {
        Write-CALog -Type Warning -Message "Could not disable certificate validation: $($_.Exception.Message)"
    }
}

# ---------------------------------------------------------------------------
# REST error handling + wrapper (with retry/back-off)
# ---------------------------------------------------------------------------
function Get-CARestError {
    <#.SYNOPSIS Extracts status code + CyberArk error body from an error record.#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $ErrorRecord
    )

    $result = [ordered]@{
        StatusCode   = $null
        ErrorCode    = $null
        ErrorMessage = $null
        Raw          = $null
    }

    $ex = $ErrorRecord.Exception

    if ($ex -and $ex.Response) {
        try { $result.StatusCode = [int]$ex.Response.StatusCode } catch { }
    }

    # PowerShell 7 usually exposes the response body here
    if ($ErrorRecord.ErrorDetails -and -not [string]::IsNullOrEmpty($ErrorRecord.ErrorDetails.Message)) {
        $result.Raw = $ErrorRecord.ErrorDetails.Message
    } elseif ($ex -and $ex.Response -and $ex.Response.GetResponseStream) {
        # Windows PowerShell 5.1 path
        try {
            $stream = $ex.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($stream)
            $result.Raw = $reader.ReadToEnd()
            $reader.Close()
            $stream.Close()
        } catch { }
    }

    if (-not [string]::IsNullOrEmpty($result.Raw)) {
        try {
            $obj = $result.Raw | ConvertFrom-Json -ErrorAction Stop
            if ($obj.PSObject.Properties['ErrorCode'])    { $result.ErrorCode    = $obj.ErrorCode }
            if ($obj.PSObject.Properties['ErrorMessage']) { $result.ErrorMessage = $obj.ErrorMessage }
        } catch { }
    }

    return [pscustomobject]$result
}

function Invoke-CARest {
    <#
    .SYNOPSIS
        Invoke a CyberArk REST call using a session, with retry/back-off.
    .DESCRIPTION
        Retries on HTTP 429, 5xx and transient network errors using
        exponential back-off (2s, 4s, 8s ...). 4xx errors (other than 429)
        are not retried. Throws a System.Exception whose .Data['StatusCode']
        holds the HTTP status code when available.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Session,

        [Parameter(Mandatory)]
        [ValidateSet('Get', 'Post', 'Patch', 'Put', 'Delete')]
        [string]$Method,

        [Parameter(Mandatory)]
        [string]$Uri,

        [Parameter()]
        [AllowNull()]
        $Body,

        [Parameter()]
        [int]$MaxRetries = 3,

        [Parameter()]
        [int]$TimeoutSec = 600
    )

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            $params = @{
                Method      = $Method
                Uri         = $Uri
                Headers     = $Session.Headers
                ContentType = 'application/json'
                TimeoutSec  = $TimeoutSec
                ErrorAction = 'Stop'
            }
            if ($null -ne $Body) {
                if ($Body -is [string]) {
                    $params.Body = $Body
                } else {
                    $params.Body = ($Body | ConvertTo-Json -Depth 10)
                }
            }
            if ($Session.SkipCertCheck -and $PSVersionTable.PSEdition -eq 'Core') {
                $params.SkipCertificateCheck = $true
            }

            return Invoke-RestMethod @params
        } catch {
            $err  = Get-CARestError -ErrorRecord $_
            $code = $err.StatusCode

            $isRetryable = ($code -eq 429) -or ($code -ge 500 -and $code -le 599) -or ($null -eq $code)
            if ($isRetryable -and $attempt -le $MaxRetries) {
                $wait = [int][math]::Pow(2, $attempt)   # 2, 4, 8 ...
                $codeText = if ($null -eq $code) { 'network error' } else { "HTTP $code" }
                Write-CALog -Type Warning -Message ("{0} on {1} {2} (attempt {3}/{4}); retrying in {5}s..." -f $codeText, $Method, $Uri, $attempt, $MaxRetries, $wait)
                Start-Sleep -Seconds $wait
                continue
            }

            $msg = "REST $Method $Uri failed"
            if ($code) { $msg += " (HTTP $code)" }
            if (-not [string]::IsNullOrEmpty($err.ErrorCode) -or -not [string]::IsNullOrEmpty($err.ErrorMessage)) {
                $msg += " - {0}: {1}" -f $err.ErrorCode, $err.ErrorMessage
            } elseif (-not [string]::IsNullOrEmpty($err.Raw)) {
                $msg += " - $($err.Raw)"
            } else {
                $msg += " - $($_.Exception.Message)"
            }

            $exception = New-Object System.Exception($msg)
            if ($code) { $exception.Data['StatusCode'] = $code }
            throw $exception
        }
    }
}

# ---------------------------------------------------------------------------
# Authentication
# ---------------------------------------------------------------------------
function Import-CAIdentityModule {
    <#
    .SYNOPSIS
        Ensures the repository's IdentityAuth.psm1 module (Get-IdentityHeader)
        is loaded.
    .DESCRIPTION
        Looks for an already-loaded Get-IdentityHeader, then for IdentityAuth.psm1
        next to this module, in the sibling "Identity Authentication" repo folder,
        or in the current directory. With -AllowDownload it falls back to fetching
        the module from the epv-api-scripts repo (same pattern other scripts use).
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$ModulePath,

        [switch]$AllowDownload
    )

    if (Get-Command -Name Get-IdentityHeader -ErrorAction SilentlyContinue) {
        return
    }

    $candidates = New-Object System.Collections.Generic.List[string]
    if (-not [string]::IsNullOrEmpty($ModulePath)) { $candidates.Add($ModulePath) }
    if ($PSScriptRoot) {
        $candidates.Add((Join-Path $PSScriptRoot '..\Identity Authentication\IdentityAuth.psm1'))
        $candidates.Add((Join-Path $PSScriptRoot 'IdentityAuth.psm1'))
    }
    $candidates.Add((Join-Path (Get-Location).Path 'IdentityAuth.psm1'))

    foreach ($candidate in $candidates) {
        if (-not [string]::IsNullOrEmpty($candidate) -and (Test-Path -LiteralPath $candidate)) {
            $resolved = (Resolve-Path -LiteralPath $candidate).Path
            Import-Module $resolved -Force -Global -ErrorAction Stop
            Write-CALog -Type Info -Message "Loaded Identity Authentication module: $resolved"
            if (Get-Command -Name Get-IdentityHeader -ErrorAction SilentlyContinue) { return }
        }
    }

    if ($AllowDownload) {
        try {
            $dest = Join-Path (Get-Location).Path 'IdentityAuth.psm1'
            Write-CALog -Type Warning -Message "IdentityAuth.psm1 not found locally; downloading it from the epv-api-scripts repository..."
            Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/cyberark/epv-api-scripts/main/Identity%20Authentication/IdentityAuth.psm1' -OutFile $dest -UseBasicParsing -ErrorAction Stop
            Import-Module $dest -Force -Global -ErrorAction Stop
            if (Get-Command -Name Get-IdentityHeader -ErrorAction SilentlyContinue) { return }
        } catch {
            throw "Failed to download/import IdentityAuth.psm1: $($_.Exception.Message)"
        }
    }

    throw "Could not locate IdentityAuth.psm1. Run these scripts from inside the epv-api-scripts repo (so '..\Identity Authentication\IdentityAuth.psm1' resolves), pass -IdentityAuthModulePath, use -DownloadIdentityAuth, or supply a -LogonToken obtained from Get-IdentityHeader."
}

function Get-CAIdentityHeader {
    <#
    .SYNOPSIS
        Authenticates to CyberArk Identity by delegating to the repository's
        Get-IdentityHeader (IdentityAuth.psm1).
    .DESCRIPTION
        Imports IdentityAuth.psm1 and calls Get-IdentityHeader, adapting to the
        parameters the installed module version exposes (PCloudURL vs
        PCloudTenantAPIURL, optional psPASFormat/IdentityTenantURL). Supports
        OAuth client credentials, username/password, and interactive (username)
        flows - the module itself handles any MFA / push / SAML+PIN challenges.
    .OUTPUTS
        Whatever Get-IdentityHeader returns (a header hashtable or a token string);
        normalise it with ConvertTo-CAHeader.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$PVWAUrl,

        [Parameter()]
        [ValidateSet('Identity', 'OAuth')]
        [string]$Mode = 'Identity',

        [Parameter()]
        [System.Management.Automation.PSCredential]$Credential,

        [Parameter()]
        [string]$IdentityUserName,

        [Parameter()]
        [string]$IdentityTenantURL,

        [Parameter()]
        [string]$ModulePath,

        [switch]$AllowDownload,

        [switch]$ForceNewSession
    )

    Import-CAIdentityModule -ModulePath $ModulePath -AllowDownload:$AllowDownload

    $command   = Get-Command -Name Get-IdentityHeader -ErrorAction Stop
    $supported = $command.Parameters.Keys
    $splat     = @{}

    # The PCloud URL parameter name differs between module versions.
    if ($supported -contains 'PCloudURL') {
        $splat['PCloudURL'] = $PVWAUrl
    } elseif ($supported -contains 'PCloudTenantAPIURL') {
        $splat['PCloudTenantAPIURL'] = $PVWAUrl
    } else {
        throw "The loaded Get-IdentityHeader does not expose a recognised PCloud URL parameter (PCloudURL / PCloudTenantAPIURL)."
    }

    if (-not [string]::IsNullOrEmpty($IdentityTenantURL) -and ($supported -contains 'IdentityTenantURL')) {
        $splat['IdentityTenantURL'] = $IdentityTenantURL
    }
    # Older module versions need -psPASFormat to return a header hashtable.
    if ($supported -contains 'psPASFormat') {
        $splat['psPASFormat'] = $true
    }
    if ($ForceNewSession -and ($supported -contains 'ForceNewSession')) {
        $splat['ForceNewSession'] = $true
    }

    if ($Mode -eq 'OAuth') {
        if (-not $Credential) { throw "OAuth Identity authentication requires -Credential (OAuth client ID as the username, client secret as the password)." }
        if (-not ($supported -contains 'OAuthCreds')) { throw "The loaded Get-IdentityHeader does not support -OAuthCreds." }
        $splat['OAuthCreds'] = $Credential
    } elseif ($Credential) {
        if (-not ($supported -contains 'UPCreds')) { throw "The loaded Get-IdentityHeader does not support -UPCreds." }
        $splat['UPCreds'] = $Credential
    } elseif (-not [string]::IsNullOrEmpty($IdentityUserName)) {
        if (-not ($supported -contains 'IdentityUserName')) { throw "The loaded Get-IdentityHeader does not support -IdentityUserName." }
        $splat['IdentityUserName'] = $IdentityUserName
    } else {
        throw "Identity authentication requires one of -Credential, -IdentityUserName, or -LogonToken."
    }

    Write-CALog -Type Info -Message "Authenticating via IdentityAuth.psm1 (Get-IdentityHeader, mode: $Mode)..."
    $result = Get-IdentityHeader @splat
    if ($null -eq $result) {
        throw "Get-IdentityHeader returned nothing - authentication failed."
    }
    return $result
}

function ConvertTo-CAHeader {
    <#.SYNOPSIS Normalises a supplied logon token (hashtable or string) to a header hashtable.#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $LogonToken
    )

    if ($LogonToken -is [System.Collections.IDictionary]) {
        if (-not ($LogonToken.Keys -contains 'Authorization')) {
            throw "The supplied -LogonToken hashtable does not contain an 'Authorization' key."
        }
        # Return a shallow copy so callers cannot mutate the original
        $copy = @{}
        foreach ($k in $LogonToken.Keys) { $copy[$k] = $LogonToken[$k] }
        return $copy
    }

    $token = [string]$LogonToken
    if ([string]::IsNullOrWhiteSpace($token)) {
        throw "The supplied -LogonToken is empty."
    }

    if ($token -match '^(?i)bearer ') {
        return @{ 'Authorization' = $token; 'X-IDAP-NATIVE-CLIENT' = 'true' }
    }
    # JWT-looking tokens are CyberArk Identity tokens and need the Bearer prefix
    if ($token -match '^eyJ') {
        return @{ 'Authorization' = "Bearer $token"; 'X-IDAP-NATIVE-CLIENT' = 'true' }
    }
    # Otherwise assume a classic PVWA session token (used raw)
    return @{ 'Authorization' = $token }
}

function Get-CAClassicLogonHeader {
    <#.SYNOPSIS Classic PVWA logon (CyberArk / LDAP / RADIUS) for self-hosted vaults.#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ApiBase,

        [Parameter(Mandatory)]
        [System.Management.Automation.PSCredential]$Credential,

        [Parameter()]
        [ValidateSet('CyberArk', 'LDAP', 'RADIUS')]
        [string]$AuthType = 'CyberArk',

        [Parameter()]
        [string]$RadiusOTP,

        [switch]$ConcurrentSession,

        [switch]$SkipCertificateValidation
    )

    $logonUrl = "$ApiBase/auth/$AuthType/Logon"
    $password = $Credential.GetNetworkCredential().Password
    if (-not [string]::IsNullOrEmpty($RadiusOTP)) {
        $password = "$password,$RadiusOTP"
    }

    $bodyHash = @{
        username = $Credential.UserName.Replace('\', '')
        password = $password
    }
    if ($ConcurrentSession) { $bodyHash.concurrentSession = $true }

    $irmParams = @{
        Uri         = $logonUrl
        Method      = 'Post'
        Body        = ($bodyHash | ConvertTo-Json)
        ContentType = 'application/json'
        ErrorAction = 'Stop'
    }
    if ($SkipCertificateValidation -and $PSVersionTable.PSEdition -eq 'Core') {
        $irmParams.SkipCertificateCheck = $true
    }

    try {
        $token = Invoke-RestMethod @irmParams
    } catch {
        throw "Logon to $logonUrl failed: $($_.Exception.Message)"
    }

    if ([string]::IsNullOrEmpty($token)) {
        throw "Logon to $logonUrl returned an empty token."
    }

    return @{ 'Authorization' = $token }
}

function New-CASession {
    <#
    .SYNOPSIS
        Authenticates to CyberArk and returns a reusable session object.
    .DESCRIPTION
        Resolution order:
          1. -LogonToken supplied            -> use it as-is (no logoff performed)
          2. -AuthType Identity (default)     -> CyberArk Identity via IdentityAuth.psm1
                                                 (interactive / username+password,
                                                 with MFA handled by the module)
          3. -AuthType OAuth                  -> CyberArk Identity OAuth client
                                                 credentials via IdentityAuth.psm1
          4. -AuthType CyberArk/LDAP/RADIUS   -> classic PVWA logon (self-hosted)
    .OUTPUTS
        PSCustomObject with: PVWAUrl, ApiBase, Headers, AuthType, CanLogoff,
        SkipCertCheck.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$PVWAUrl,

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

        [switch]$DownloadIdentityAuth,

        [Parameter()]
        [string]$RadiusOTP,

        [switch]$ConcurrentSession,

        [switch]$SkipCertificateValidation
    )

    $apiBase = Get-CAApiBaseUrl -PVWAUrl $PVWAUrl
    if ($SkipCertificateValidation) {
        Set-CACertificateValidation -Skip
    }

    $headers      = $null
    $canLogoff    = $false
    $resolvedAuth = $AuthType

    if ($null -ne $LogonToken) {
        $headers      = ConvertTo-CAHeader -LogonToken $LogonToken
        $resolvedAuth = 'Token'
        $canLogoff    = $false
        Write-CALog -Type Info -Message "Using supplied logon token (the session will NOT be logged off)."
    } elseif ($AuthType -eq 'Identity' -or $AuthType -eq 'OAuth') {
        $mode = 'Identity'
        if ($AuthType -eq 'OAuth') { $mode = 'OAuth' }
        $result = Get-CAIdentityHeader -PVWAUrl $PVWAUrl -Mode $mode -Credential $Credential `
            -IdentityUserName $IdentityUserName -IdentityTenantURL $IdentityTenantURL `
            -ModulePath $IdentityAuthModulePath -AllowDownload:$DownloadIdentityAuth
        $headers      = ConvertTo-CAHeader -LogonToken $result
        $resolvedAuth = 'Identity'
        $canLogoff    = $false   # the Identity module manages its own session/token lifetime
        Write-CALog -Type Success -Message "Authenticated to CyberArk Identity via IdentityAuth.psm1."
    } else {
        if (-not $Credential) {
            throw "$AuthType authentication requires -Credential."
        }
        $headers   = Get-CAClassicLogonHeader -ApiBase $apiBase -Credential $Credential -AuthType $AuthType -RadiusOTP $RadiusOTP -ConcurrentSession:$ConcurrentSession -SkipCertificateValidation:$SkipCertificateValidation
        $canLogoff = $true
        Write-CALog -Type Success -Message "Authenticated to CyberArk ($AuthType logon)."
    }

    return [pscustomobject]@{
        PVWAUrl       = $PVWAUrl.Trim().TrimEnd('/')
        ApiBase       = $apiBase
        Headers       = $headers
        AuthType      = $resolvedAuth
        CanLogoff     = $canLogoff
        SkipCertCheck = [bool]$SkipCertificateValidation
    }
}

function Close-CASession {
    <#.SYNOPSIS Logs off a classic PVWA session. No-op for token / OAuth sessions.#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Session
    )

    if ($Session.CanLogoff) {
        try {
            Invoke-CARest -Session $Session -Method Post -Uri "$($Session.ApiBase)/auth/Logoff" | Out-Null
            Write-CALog -Type Info -Message "Logged off the CyberArk session."
        } catch {
            Write-CALog -Type Warning -Message "Logoff failed: $($_.Exception.Message)"
        }
    } else {
        Write-CALog -Type Debug -Message "Identity / OAuth / token session - no logoff performed."
    }
}

# ---------------------------------------------------------------------------
# Account retrieval
# ---------------------------------------------------------------------------
function Get-CAAccountList {
    <#
    .SYNOPSIS
        Retrieves all accounts using offset/limit pagination.
    .DESCRIPTION
        Pages through GET /api/Accounts using limit/offset until the API stops
        returning a nextLink or a full page. Optional Safe and free-text search
        filters are supported.
    .OUTPUTS
        System.Collections.Generic.List[object] of account list items.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Session,

        [Parameter()]
        [string]$SafeName,

        [Parameter()]
        [string]$Search,

        [Parameter()]
        [ValidateRange(1, 1000)]
        [int]$PageSize = 1000
    )

    $accounts = New-Object System.Collections.Generic.List[object]
    $offset   = 0
    $baseUrl  = "$($Session.ApiBase)/Accounts"

    while ($true) {
        $query = "?limit=$PageSize&offset=$offset"
        if (-not [string]::IsNullOrEmpty($SafeName)) {
            $query += "&filter=" + [uri]::EscapeDataString("safeName eq $SafeName")
        }
        if (-not [string]::IsNullOrEmpty($Search)) {
            $query += "&search=" + [uri]::EscapeDataString($Search)
        }

        $uri  = $baseUrl + $query
        $resp = Invoke-CARest -Session $Session -Method Get -Uri $uri

        $pageItems = @()
        if ($resp.PSObject.Properties['value'] -and $resp.value) {
            $pageItems = @($resp.value)
        }
        foreach ($item in $pageItems) { $accounts.Add($item) }

        $got = $pageItems.Count
        Write-CALog -Type Info -Message ("Retrieved {0} accounts (running total: {1})." -f $got, $accounts.Count)

        $hasNext = $false
        if ($resp.PSObject.Properties['nextLink'] -and -not [string]::IsNullOrEmpty($resp.nextLink)) {
            $hasNext = $true
        }

        if ($got -eq 0) { break }
        if (-not $hasNext -and $got -lt $PageSize) { break }

        $offset += $PageSize
    }

    return $accounts
}

function Get-CAAccountDetail {
    <#.SYNOPSIS Retrieves full account details (incl. platformAccountProperties) by ID.#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Session,

        [Parameter(Mandatory)]
        [string]$AccountID
    )

    $uri = "$($Session.ApiBase)/Accounts/$AccountID"
    return Invoke-CARest -Session $Session -Method Get -Uri $uri
}

# ---------------------------------------------------------------------------
# Flattening + diffing custom File Categories
# ---------------------------------------------------------------------------
function ConvertTo-CAFlatRecord {
    <#
    .SYNOPSIS
        Flattens a CyberArk account object into an ordered hashtable of CSV cells.
    .DESCRIPTION
        Produces base/identity columns, one column per custom File Category
        (platformAccountProperties), and a few read-only metadata columns
        (underscore-prefixed). Custom property names that collide with a base
        column name are prefixed with 'FC_'.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Account
    )

    $record = [ordered]@{}
    $record['AccountID']  = ConvertTo-CAStringValue $Account.id
    $record['Name']       = ConvertTo-CAStringValue $Account.name
    $record['Username']   = ConvertTo-CAStringValue $Account.userName
    $record['Address']    = ConvertTo-CAStringValue $Account.address
    $record['SafeName']   = ConvertTo-CAStringValue $Account.safeName
    $record['PlatformID'] = ConvertTo-CAStringValue $Account.platformId
    $record['SecretType'] = ConvertTo-CAStringValue $Account.secretType

    # Custom File Categories (platformAccountProperties). Notes lives here too.
    if ($Account.PSObject.Properties['platformAccountProperties'] -and $Account.platformAccountProperties) {
        foreach ($prop in $Account.platformAccountProperties.PSObject.Properties) {
            $column = $prop.Name
            if ($script:CABaseColumns -contains $column) {
                $column = "FC_$column"
            }
            $record[$column] = ConvertTo-CAStringValue $prop.Value
        }
    }

    # Read-only metadata (informational; ignored by the import script)
    $autoMgmt     = $null
    $manualReason = $null
    if ($Account.PSObject.Properties['secretManagement'] -and $Account.secretManagement) {
        if ($Account.secretManagement.PSObject.Properties['automaticManagementEnabled']) {
            $autoMgmt = $Account.secretManagement.automaticManagementEnabled
        }
        if ($Account.secretManagement.PSObject.Properties['manualManagementReason']) {
            $manualReason = $Account.secretManagement.manualManagementReason
        }
    }
    $record['_AutomaticManagementEnabled'] = ConvertTo-CAStringValue $autoMgmt
    $record['_ManualManagementReason']     = ConvertTo-CAStringValue $manualReason

    $createdTime = $null
    if ($Account.PSObject.Properties['createdTime']) { $createdTime = $Account.createdTime }
    $record['_CreatedTime'] = Convert-CAEpoch $createdTime

    $catModTime = $null
    if ($Account.PSObject.Properties['categoryModificationTime']) { $catModTime = $Account.categoryModificationTime }
    $record['_CategoryModificationTime'] = Convert-CAEpoch $catModTime

    return $record
}

function Get-CAAccountProperty {
    <#.SYNOPSIS Case-insensitively returns a platformAccountProperties property (Name+Value) or $null.#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Account,

        [Parameter(Mandatory)]
        [string]$Name
    )

    if (-not ($Account.PSObject.Properties['platformAccountProperties'] -and $Account.platformAccountProperties)) {
        return $null
    }
    foreach ($prop in $Account.platformAccountProperties.PSObject.Properties) {
        if ($prop.Name -ieq $Name) { return $prop }
    }
    return $null
}

function Get-CAPropertyUpdateOperation {
    <#
    .SYNOPSIS
        Computes the JSON-Patch operations needed to bring an account's
        properties in line with the desired values.
    .DESCRIPTION
        Each desired property is handled as either an editable top-level field
        (Address -> /address, Username -> /userName, Name -> /name) or a custom
        File Category (-> /platformAccountProperties/<Name>).

        For custom File Categories:
          * absent on account + non-empty value -> 'add'
          * present + different value            -> 'replace'
          * present + same value                 -> no operation
          * empty value                          -> skipped, unless
            -RemoveEmptyValues, in which case an existing property is 'remove'd

        For top-level fields:
          * different value -> 'replace'; same value -> no operation
          * empty value     -> skipped (top-level fields are never cleared/removed)

        Secrets and secret management are never touched.
    .OUTPUTS
        PSCustomObject with Operations (list) and ChangedFields (list of text).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Account,

        [Parameter(Mandatory)]
        [hashtable]$Updates,

        [switch]$RemoveEmptyValues
    )

    $operations    = New-Object System.Collections.Generic.List[object]
    $changed       = New-Object System.Collections.Generic.List[string]
    $editableBase  = Get-CAEditableBaseFieldMap

    foreach ($name in $Updates.Keys) {
        $desired = $Updates[$name]
        if ($null -ne $desired) { $desired = ([string]$desired).Trim() }
        $isEmpty = [string]::IsNullOrEmpty($desired)

        # --- Editable top-level account field (e.g. Address -> /address) ---
        $baseKey = $null
        foreach ($key in $editableBase.Keys) {
            if ($key -ieq $name) { $baseKey = $key; break }
        }
        if ($null -ne $baseKey) {
            if ($isEmpty) { continue }   # never clear a top-level field implicitly
            $map         = $editableBase[$baseKey]
            $currentBase = ConvertTo-CAStringValue (Get-CABaseFieldValue -Account $Account -Property $map.Property)
            if ($currentBase -ne $desired) {
                $operations.Add([ordered]@{ op = 'replace'; path = $map.Path; value = $desired })
                $changed.Add("{0}: '{1}' -> '{2}'" -f $baseKey, $currentBase, $desired)
            }
            continue
        }

        # --- Custom File Category (platformAccountProperties) ---
        $existing  = Get-CAAccountProperty -Account $Account -Name $name
        $exists    = ($null -ne $existing)
        $canonical = $name
        $current   = $null
        if ($exists) {
            $canonical = $existing.Name
            $current   = ConvertTo-CAStringValue $existing.Value
        }

        if ($isEmpty) {
            if ($RemoveEmptyValues -and $exists) {
                $operations.Add([ordered]@{ op = 'remove'; path = "/platformAccountProperties/$canonical" })
                $changed.Add("{0}: '{1}' -> (removed)" -f $canonical, $current)
            }
            continue
        }

        if ($exists) {
            if ($current -ne $desired) {
                $operations.Add([ordered]@{ op = 'replace'; path = "/platformAccountProperties/$canonical"; value = $desired })
                $changed.Add("{0}: '{1}' -> '{2}'" -f $canonical, $current, $desired)
            }
        } else {
            $operations.Add([ordered]@{ op = 'add'; path = "/platformAccountProperties/$name"; value = $desired })
            $changed.Add("{0}: (none) -> '{1}'" -f $name, $desired)
        }
    }

    return [pscustomobject]@{
        Operations    = $operations
        ChangedFields = $changed
    }
}

function ConvertTo-CAJsonArray {
    <#.SYNOPSIS Serialises a list of operations to a JSON array (always bracketed).#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Items,

        [Parameter()]
        [int]$Depth = 6
    )

    $json = ConvertTo-Json -InputObject @($Items) -Depth $Depth
    if (-not $json.TrimStart().StartsWith('[')) {
        $json = "[$json]"
    }
    return $json
}

Export-ModuleMember -Function @(
    'Get-CABaseColumns',
    'Get-CAMetadataColumns',
    'Get-CAEditableBaseFieldMap',
    'Get-CAProtectedColumns',
    'Get-CABaseFieldValue',
    'Initialize-CALog',
    'Write-CALog',
    'ConvertTo-CAStringValue',
    'Convert-CAEpoch',
    'Get-CAApiBaseUrl',
    'Set-CACertificateValidation',
    'Get-CARestError',
    'Invoke-CARest',
    'Import-CAIdentityModule',
    'Get-CAIdentityHeader',
    'ConvertTo-CAHeader',
    'Get-CAClassicLogonHeader',
    'New-CASession',
    'Close-CASession',
    'Get-CAAccountList',
    'Get-CAAccountDetail',
    'ConvertTo-CAFlatRecord',
    'Get-CAAccountProperty',
    'Get-CAPropertyUpdateOperation',
    'ConvertTo-CAJsonArray'
)

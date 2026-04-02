# Release/Check-in all accounts from the CSV on the SOURCE (old) environment
# Run this from the "Migration via REST" directory

# Force .NET to use WinHttpHandler (trusts Windows cert store like your browser)
$env:DOTNET_SYSTEM_NET_HTTP_USESOCKETSHTTPHANDLER = '0'

# --- Config ---
$PVWAURL   = "https://prdcapw1-esaw2a.uprising.t-mobile.com/PasswordVault"
$AuthType  = "cyberark"
$CSV       = ".\ExportOfAccounts.csv"

# --- Authenticate ---
Write-Host "Authenticating to $PVWAURL ..." -ForegroundColor Cyan
$cred = Get-Credential -Message "Enter your $AuthType credentials"
$body = @{
    username          = $cred.UserName
    password          = $cred.GetNetworkCredential().Password
    concurrentSession = $true
} | ConvertTo-Json

$token = Invoke-RestMethod -Uri "$PVWAURL/api/auth/$AuthType/Logon" -Method Post -Body $body -ContentType "application/json"
$headers = @{ Authorization = $token }

Write-Host "Logged in." -ForegroundColor Green

# --- Load accounts from CSV ---
$accounts = Import-Csv $CSV
Write-Host "Loaded $($accounts.Count) accounts from CSV.`n" -ForegroundColor Cyan

$released = 0
$skipped  = 0

foreach ($acct in $accounts) {
    $id   = $acct.id
    $name = $acct.name

    try {
        Invoke-RestMethod -Uri "$PVWAURL/api/Accounts/$id/CheckIn" -Method Post -Headers $headers -ContentType "application/json" | Out-Null
        Write-Host "[OK]   $name ($id)" -ForegroundColor Green
        $released++
    }
    catch {
        # Not checked out, or already released — just skip
        $skipped++
    }
}

Write-Host "`nDone! Released: $released  |  Skipped (not checked out): $skipped" -ForegroundColor Cyan

# --- Logoff ---
try { Invoke-RestMethod -Uri "$PVWAURL/api/auth/Logoff" -Method Post -Headers $headers | Out-Null } catch {}

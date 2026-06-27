$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()

$appDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$envPath = Join-Path $appDir ".env"

Write-Host "Set an app password for the web form."
Write-Host "Username will be: user"
Write-Host "Password input will not be displayed."
$secure = Read-Host "APP_PASSWORD" -AsSecureString
$bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
try {
    $password = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
}
finally {
    if ($bstr -ne [IntPtr]::Zero) {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
}

if ([string]::IsNullOrWhiteSpace($password)) {
    Write-Host "Password is empty. Canceled."
    Read-Host "Press Enter to close"
    exit 1
}

$lines = @()
if (Test-Path $envPath) {
    $lines = Get-Content -LiteralPath $envPath -Encoding UTF8 | Where-Object { $_ -notmatch '^APP_PASSWORD=' }
}
$lines += "APP_PASSWORD=$password"
$lines | Set-Content -LiteralPath $envPath -Encoding UTF8

Write-Host "Done. Restart the app to apply the password."
Read-Host "Press Enter to close"

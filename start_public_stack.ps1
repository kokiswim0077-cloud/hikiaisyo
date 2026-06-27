param(
    [Parameter(Mandatory=$true)]
    [string]$Domain
)

$ErrorActionPreference = "Stop"
$appDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$envPath = Join-Path $appDir ".env"
$caddy = Get-ChildItem -Path "$env:LOCALAPPDATA\Microsoft\WinGet\Packages" -Recurse -Filter caddy.exe -ErrorAction SilentlyContinue |
    Select-Object -First 1 -ExpandProperty FullName

if (-not $caddy) {
    throw "caddy.exe was not found. Install Caddy first."
}

if (-not (Test-Path $envPath)) {
    throw ".env was not found. Set GEMINI_API_KEY and APP_PASSWORD first."
}

$envText = Get-Content -LiteralPath $envPath -Encoding UTF8
if (-not ($envText -match '^APP_PASSWORD=.+')) {
    throw "APP_PASSWORD is not set. Run set_app_password.ps1 first."
}
if (-not ($envText -match '^GEMINI_API_KEY=.+')) {
    throw "GEMINI_API_KEY is not set. Run set_gemini_key.ps1 first."
}

$caddyfile = Join-Path $appDir "Caddyfile.production"
@"
$Domain {
    encode gzip

    reverse_proxy 127.0.0.1:8765 {
        header_up X-Forwarded-Proto {scheme}
        header_up X-Forwarded-Host {host}
        header_up X-Real-IP {remote_host}
    }

    header {
        X-Content-Type-Options nosniff
        X-Frame-Options DENY
        Referrer-Policy no-referrer
    }
}
"@ | Set-Content -LiteralPath $caddyfile -Encoding UTF8

Set-Location $appDir

$env:PUBLIC_MODE = "true"
$env:REQUIRE_HTTPS = "true"
$env:ENABLE_HSTS = "true"
$env:TRUST_PROXY_HEADERS = "true"
$env:MAX_UPLOAD_MB = "12"
$env:OUTPUT_RETENTION_MINUTES = "120"
$env:DOWNLOAD_TOKEN_MINUTES = "60"

$waitressArgs = @('-m','waitress','--listen=127.0.0.1:8765','--threads=6','app:app')
Start-Process -WindowStyle Hidden python -ArgumentList $waitressArgs -WorkingDirectory $appDir
Start-Sleep -Seconds 2

& $caddy validate --config $caddyfile
& $caddy run --config $caddyfile

$ErrorActionPreference = "Stop"
$appDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$publicIp = (Invoke-RestMethod -Uri "https://api.ipify.org" -TimeoutSec 10).Trim()
if (-not ($publicIp -match '^\d{1,3}(\.\d{1,3}){3}$')) {
    throw "Could not determine a valid IPv4 public IP."
}

$domain = ($publicIp -replace '\.', '-') + ".sslip.io"
Write-Host "Public IP: $publicIp"
Write-Host "Domain: https://$domain"
Write-Host ""
Write-Host "Before this can work from the internet, forward router TCP 443 to this PC."
Write-Host "Then keep this window open."
Write-Host ""

& powershell -ExecutionPolicy Bypass -File (Join-Path $appDir "start_public_stack.ps1") -Domain $domain

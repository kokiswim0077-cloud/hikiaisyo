$ErrorActionPreference = "Stop"
$appDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $appDir

$env:REQUIRE_HTTPS = "true"
$env:ENABLE_HSTS = "true"
$env:TRUST_PROXY_HEADERS = "true"

python -m waitress --listen=127.0.0.1:8765 --threads=6 app:app

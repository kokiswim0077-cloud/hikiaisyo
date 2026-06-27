$ErrorActionPreference = "Stop"
$appDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $appDir
Start-Process -WindowStyle Hidden python -ArgumentList @((Join-Path $appDir "app.py"))
Start-Sleep -Seconds 2
Start-Process "http://127.0.0.1:8765"

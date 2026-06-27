$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()

$appDir = "C:\Users\koki0\inquiry_voice_form"
$envPath = Join-Path $appDir ".env"

Write-Host "Paste your Gemini API key, then press Enter."
Write-Host "The typed key will not be displayed."
$secure = Read-Host "GEMINI_API_KEY" -AsSecureString
$bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
try {
    $key = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
}
finally {
    if ($bstr -ne [IntPtr]::Zero) {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
}

if ([string]::IsNullOrWhiteSpace($key)) {
    Write-Host "API key is empty. Canceled."
    Read-Host "Press Enter to close"
    exit 1
}

"GEMINI_API_KEY=$key" | Set-Content -LiteralPath $envPath -Encoding UTF8
[Environment]::SetEnvironmentVariable("GEMINI_API_KEY", $key, "User")

$line = netstat -ano | Select-String ":8765\s" | Select-Object -First 1
if ($line) {
    $pidText = (($line.ToString() -split "\s+") | Select-Object -Last 1)
    Stop-Process -Id ([int]$pidText) -Force
    Start-Sleep -Seconds 1
}

Start-Process -WindowStyle Hidden python -ArgumentList @((Join-Path $appDir "app.py"))
Start-Sleep -Seconds 2
Start-Process "http://127.0.0.1:8765"

Write-Host ""
Write-Host "Done. The app has been restarted and opened in your browser."
Read-Host "Press Enter to close"

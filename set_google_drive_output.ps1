param(
    [Parameter(Mandatory = $true)]
    [string]$GoogleDriveFolder
)

$resolved = Resolve-Path -LiteralPath $GoogleDriveFolder -ErrorAction Stop
$outputDir = Join-Path $resolved.Path "引合書_見積書"
New-Item -ItemType Directory -Force -Path $outputDir | Out-Null

[Environment]::SetEnvironmentVariable("OUTPUT_DIR", $outputDir, "User")
Write-Host "保存先を設定しました: $outputDir"
Write-Host "アプリを起動中の場合は、いったん閉じて再起動してください。"

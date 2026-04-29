$proxyDir = $PSScriptRoot
$caddyfile = Join-Path $proxyDir "Caddyfile"
$logDir = Join-Path $proxyDir "logs"
$stdoutLog = Join-Path $logDir "caddy_stdout.log"
$stderrLog = Join-Path $logDir "caddy_stderr.log"

if (-not (Test-Path $caddyfile)) {
    Write-Host "Caddyfile nao encontrado: $caddyfile" -ForegroundColor Red
    exit 1
}

$caddyExe = $null
$cmd = Get-Command caddy -ErrorAction SilentlyContinue
if ($cmd) {
    $caddyExe = $cmd.Source
}
if (-not $caddyExe) {
    $fallback = "C:\Program Files\Caddy\caddy.exe"
    if (Test-Path $fallback) {
        $caddyExe = $fallback
    }
}
if (-not $caddyExe) {
    $wingetCaddy = Get-ChildItem "$env:LOCALAPPDATA\Microsoft\WinGet\Packages" -Filter "caddy.exe" -Recurse -ErrorAction SilentlyContinue |
        Select-Object -First 1 -ExpandProperty FullName
    if ($wingetCaddy -and (Test-Path $wingetCaddy)) {
        $caddyExe = $wingetCaddy
    }
}
if (-not $caddyExe) {
    Write-Host "Caddy nao encontrado no sistema." -ForegroundColor Red
    exit 1
}

New-Item -ItemType Directory -Path $logDir -Force | Out-Null

Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -match '^caddy(\.exe)?$' } |
    ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }

Start-Process -FilePath $caddyExe `
    -ArgumentList @("run", "--config", $caddyfile, "--adapter", "caddyfile") `
    -WorkingDirectory $proxyDir `
    -WindowStyle Hidden `
    -RedirectStandardOutput $stdoutLog `
    -RedirectStandardError $stderrLog

Write-Host "Proxy reverso iniciado com Caddy." -ForegroundColor Green

$proxyDir = $PSScriptRoot
$caddyfile = Join-Path $proxyDir "Caddyfile"
$logDir = Join-Path $proxyDir "logs"
$stdoutLog = Join-Path $logDir "caddy_stdout.log"
$stderrLog = Join-Path $logDir "caddy_stderr.log"
$localCaddyExe = Join-Path $proxyDir "caddy.exe"

function Get-ProcessSnapshot {
    if (Get-Command Get-CimInstance -ErrorAction SilentlyContinue) {
        return Get-CimInstance Win32_Process -ErrorAction SilentlyContinue
    }

    if (Get-Command Get-WmiObject -ErrorAction SilentlyContinue) {
        return Get-WmiObject Win32_Process -ErrorAction SilentlyContinue
    }

    return @()
}

if (-not (Test-Path $caddyfile)) {
    Write-Host "Caddyfile nao encontrado: $caddyfile" -ForegroundColor Red
    exit 1
}

$caddyExe = $null
$cmd = $null
$localCmd = Get-Command $localCaddyExe -ErrorAction SilentlyContinue
if ($localCmd) {
    $caddyExe = $localCmd.Source
}
if (-not $caddyExe) {
    $cmd = Get-Command caddy -ErrorAction SilentlyContinue
}
if ($cmd -and -not $caddyExe) {
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

Get-ProcessSnapshot |
    Where-Object { $_.Name -match '^caddy(\.exe)?$' } |
    ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }

Start-Process -FilePath $caddyExe `
    -ArgumentList @("run", "--config", $caddyfile, "--adapter", "caddyfile") `
    -WorkingDirectory $proxyDir `
    -WindowStyle Hidden `
    -RedirectStandardOutput $stdoutLog `
    -RedirectStandardError $stderrLog

Write-Host "Proxy reverso iniciado com Caddy." -ForegroundColor Green

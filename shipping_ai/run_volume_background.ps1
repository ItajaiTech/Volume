# Inicia o Volume em segundo plano, sem depender de uma janela aberta do PowerShell.
$appPath = "C:\Volume\shipping_ai"
$launcherScript = "$appPath\start_volume.ps1"
$logDir = "$appPath\logs"
$stdoutLog = "$logDir\volume_stdout.log"
$stderrLog = "$logDir\volume_stderr.log"

if (-not (Test-Path $launcherScript)) {
    Write-Error "Script principal nao encontrado: $launcherScript"
    exit 1
}

New-Item -ItemType Directory -Path $logDir -Force | Out-Null

$env:VOLUME_OPEN_BROWSER = "0"
$env:VOLUME_DEBUG = "0"
$env:VOLUME_USE_WAITRESS = "1"
if (-not $env:VOLUME_BIND_HOST) {
    $env:VOLUME_BIND_HOST = "0.0.0.0"
}
if (-not $env:VOLUME_PUBLIC_HOST) {
    $env:VOLUME_PUBLIC_HOST = "util.local"
}

$existingListener = Get-NetTCPConnection -LocalPort 6100 -State Listen -ErrorAction SilentlyContinue
if ($existingListener) {
    Write-Host "Volume ja esta em execucao na porta 6100."
    exit 0
}

Start-Process powershell.exe `
    -WorkingDirectory $appPath `
    -ArgumentList @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", $launcherScript
    ) `
    -WindowStyle Hidden `
    -RedirectStandardOutput $stdoutLog `
    -RedirectStandardError $stderrLog
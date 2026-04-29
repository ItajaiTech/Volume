# Inicia o Volume em segundo plano, sem depender de uma janela aberta do PowerShell.
$appPath = $PSScriptRoot
if (-not $appPath) {
    $appPath = Split-Path -Parent $MyInvocation.MyCommand.Path
}
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
    $env:VOLUME_PUBLIC_HOST = "volume.local"
}
if (-not $env:VOLUME_PORT) {
    $env:VOLUME_PORT = "6100"
}

$existingListener = Get-NetTCPConnection -LocalPort ([int]$env:VOLUME_PORT) -State Listen -ErrorAction SilentlyContinue
if ($existingListener) {
    $listenerPid = $existingListener[0].OwningProcess
    $listenerProcess = Get-CimInstance Win32_Process -Filter "ProcessId = $listenerPid" -ErrorAction SilentlyContinue
    $isVolumeProcess = $listenerProcess -and $listenerProcess.CommandLine -match "app_volum\.py"

    if ($isVolumeProcess) {
        Write-Host "Volume ja esta em execucao na porta $($env:VOLUME_PORT)."
        exit 0
    }

    Write-Host "Porta $($env:VOLUME_PORT) ocupada por outro processo (PID $listenerPid). Nao foi possivel iniciar o Volume em background." -ForegroundColor Yellow
    exit 1
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
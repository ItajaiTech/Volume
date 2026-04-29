# Inicia o sistema Volume em ambiente virtual isolado
$appPath = "C:\Volume\shipping_ai"
$venvDir = "$appPath\.venv"
$venvPython = "$venvDir\Scripts\python.exe"

function New-VolumeVenv {
    param(
        [string]$TargetPath,
        [string]$ExpectedPython
    )

    $attempts = @(
        @{ Label = "C:\python314\python.exe"; Command = "C:\python314\python.exe"; Args = @("-m", "venv", $TargetPath) },
        @{ Label = "py -3"; Command = "py"; Args = @("-3", "-m", "venv", $TargetPath) },
        @{ Label = "python"; Command = "python"; Args = @("-m", "venv", $TargetPath) }
    )

    foreach ($attempt in $attempts) {
        $cmd = Get-Command $attempt.Command -ErrorAction SilentlyContinue
        if (-not $cmd) {
            continue
        }

        Write-Host "Tentando criar .venv com $($attempt.Label)..." -ForegroundColor DarkYellow
        & $attempt.Command @($attempt.Args)

        if ($LASTEXITCODE -eq 0 -and (Test-Path $ExpectedPython)) {
            return $true
        }
    }

    return $false
}

if (-not (Test-Path $venvPython)) {
    Write-Host "Ambiente virtual local nao encontrado. Criando em $venvDir ..." -ForegroundColor Yellow
    $created = New-VolumeVenv -TargetPath $venvDir -ExpectedPython $venvPython
    if (-not $created) {
        Write-Host "Nao foi possivel criar o .venv do Volume." -ForegroundColor Red
        Write-Host "Instale Python 3 e tente novamente." -ForegroundColor Red
        exit 1
    }
}

$env:VOLUME_BIND_HOST = "0.0.0.0"
$env:VOLUME_PUBLIC_HOST = "volume.local"
$env:VOLUME_PORT = "6100"
if (-not $env:VOLUME_URL_SCHEME) {
    $env:VOLUME_URL_SCHEME = "https"
}
if (-not $env:VOLUME_SERVER_NAME) {
    $env:VOLUME_SERVER_NAME = $env:VOLUME_PUBLIC_HOST
}
if (-not $env:VOLUME_TRUSTED_HOSTS) {
    $env:VOLUME_TRUSTED_HOSTS = "util.local,util.local:6100,volume.local,volume.local:6100"
}

if (-not $env:VOLUME_DEBUG) {
    $env:VOLUME_DEBUG = "0"
}
if (-not $env:VOLUME_USE_WAITRESS) {
    $env:VOLUME_USE_WAITRESS = "1"
}

$existingListener = Get-NetTCPConnection -LocalPort ([int]$env:VOLUME_PORT) -State Listen -ErrorAction SilentlyContinue
if ($existingListener) {
    Write-Host "Volume ja esta em execucao na porta $($env:VOLUME_PORT)." -ForegroundColor Yellow
    Start-Process "$($env:VOLUME_URL_SCHEME)://$($env:VOLUME_PUBLIC_HOST):$($env:VOLUME_PORT)"
    exit 0
}

Write-Host "Atualizando dependencias do Volume..." -ForegroundColor Yellow
& $venvPython -m pip install --disable-pip-version-check -r "$appPath\requirements.txt"
if ($LASTEXITCODE -ne 0) {
    Write-Host "Falha ao instalar dependencias do Volume." -ForegroundColor Red
    exit 1
}

Set-Location $appPath
Write-Host "Iniciando Volume em $($env:VOLUME_URL_SCHEME)://$($env:VOLUME_PUBLIC_HOST):$($env:VOLUME_PORT)" -ForegroundColor Green
Write-Host "Bind interno: $($env:VOLUME_URL_SCHEME)://$($env:VOLUME_BIND_HOST):$($env:VOLUME_PORT)" -ForegroundColor DarkGray

if ($env:VOLUME_OPEN_BROWSER -ne "0") {
    Start-Process "$($env:VOLUME_URL_SCHEME)://$($env:VOLUME_PUBLIC_HOST):$($env:VOLUME_PORT)"
}

& $venvPython "$appPath\app_volum.py"

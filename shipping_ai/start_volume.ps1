# Inicia o sistema Volume em ambiente virtual isolado
$appPath = $PSScriptRoot
if (-not $appPath) {
    $appPath = Split-Path -Parent $MyInvocation.MyCommand.Path
}
$venvDir = "$appPath\.venv"
$venvPython = "$venvDir\Scripts\python.exe"

function Test-VolumePython {
    param(
        [string]$PythonPath
    )

    if (-not (Test-Path $PythonPath)) {
        return $false
    }

    & $PythonPath --version *> $null
    return ($LASTEXITCODE -eq 0)
}

function New-VolumeVenv {
    param(
        [string]$TargetPath,
        [string]$ExpectedPython
    )

    $localPython = Join-Path $env:LocalAppData "Programs\Python\Python312\python.exe"

    $attempts = @(
        @{ Label = "LocalAppData Python312"; Command = $localPython; Args = @("-m", "venv", $TargetPath) },
        @{ Label = "C:\python314\python.exe"; Command = "C:\python314\python.exe"; Args = @("-m", "venv", $TargetPath) },
        @{ Label = "py -3"; Command = "py"; Args = @("-3", "-m", "venv", $TargetPath) },
        @{ Label = "python"; Command = "python"; Args = @("-m", "venv", $TargetPath) },
        @{ Label = "uv venv"; Command = "uv"; Args = @("venv", $TargetPath) }
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

if (-not (Test-VolumePython -PythonPath $venvPython)) {
    if (Test-Path $venvDir) {
        Write-Host "Ambiente virtual existente esta invalido. Recriando em $venvDir ..." -ForegroundColor Yellow
        Remove-Item -Path $venvDir -Recurse -Force -ErrorAction SilentlyContinue
    } else {
        Write-Host "Ambiente virtual local nao encontrado. Criando em $venvDir ..." -ForegroundColor Yellow
    }

    $created = New-VolumeVenv -TargetPath $venvDir -ExpectedPython $venvPython
    if (-not $created -or -not (Test-VolumePython -PythonPath $venvPython)) {
        Write-Host "Nao foi possivel criar o .venv do Volume." -ForegroundColor Red
        Write-Host "Instale Python 3 e tente novamente." -ForegroundColor Red
        exit 1
    }
}

if (-not $env:VOLUME_BIND_HOST) {
    $env:VOLUME_BIND_HOST = "0.0.0.0"
}
if (-not $env:VOLUME_PUBLIC_HOST) {
    $env:VOLUME_PUBLIC_HOST = "volume.local"
}
if (-not $env:VOLUME_PORT) {
    $env:VOLUME_PORT = "6100"
}
if (-not $env:VOLUME_URL_SCHEME) {
    $env:VOLUME_URL_SCHEME = "https"
}
if (-not $env:VOLUME_SSL_CERT_FILE) {
    $env:VOLUME_SSL_CERT_FILE = Join-Path $appPath "certs\volume.local.crt"
}
if (-not $env:VOLUME_SSL_KEY_FILE) {
    $env:VOLUME_SSL_KEY_FILE = Join-Path $appPath "certs\volume.local.key"
}
if (-not $env:VOLUME_SERVER_NAME) {
    $env:VOLUME_SERVER_NAME = $env:VOLUME_PUBLIC_HOST
}
if (-not $env:VOLUME_TRUSTED_HOSTS) {
    $env:VOLUME_TRUSTED_HOSTS = "volume.local,volume.local:6100"
}

if (-not $env:VOLUME_DEBUG) {
    $env:VOLUME_DEBUG = "0"
}
if (-not $env:VOLUME_USE_WAITRESS) {
    $env:VOLUME_USE_WAITRESS = "1"
}

$httpsRequested = $env:VOLUME_URL_SCHEME -eq "https"
$hasCertFiles = (Test-Path $env:VOLUME_SSL_CERT_FILE) -and (Test-Path $env:VOLUME_SSL_KEY_FILE)
if ($httpsRequested -and $hasCertFiles -and $env:VOLUME_USE_WAITRESS -eq "1") {
    Write-Host "HTTPS com certificado detectado. Desativando Waitress para usar TLS nativo do Flask." -ForegroundColor Yellow
    $env:VOLUME_USE_WAITRESS = "0"
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

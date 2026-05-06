$ErrorActionPreference = 'Stop'

$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$certFiles = @(
    'ponto.crt',
    'ponto_mobile.crt',
    'ponto_mobile_complete.pem',
    'util.local.crt',
    'volume.local.crt'
)

$stores = @(
    'Cert:\CurrentUser\Root',
    'Cert:\CurrentUser\CA'
)

$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if ($isAdmin) {
    $stores += 'Cert:\LocalMachine\Root'
    $stores += 'Cert:\LocalMachine\CA'
}

foreach ($file in $certFiles) {
    $fullPath = Join-Path $baseDir $file
    if (-not (Test-Path $fullPath)) {
        Write-Warning "Arquivo nao encontrado: $fullPath"
        continue
    }

    foreach ($store in $stores) {
        try {
            Import-Certificate -FilePath $fullPath -CertStoreLocation $store | Out-Null
            Write-Host "OK: $file -> $store" -ForegroundColor Green
        }
        catch {
            Write-Warning "Falha ao importar $file em ${store}: $($_.Exception.Message)"
        }
    }
}

Write-Host 'Instalacao concluida.' -ForegroundColor Cyan
if (-not $isAdmin) {
    Write-Host 'Dica: execute como Administrador para instalar tambem no LocalMachine.' -ForegroundColor Yellow
}

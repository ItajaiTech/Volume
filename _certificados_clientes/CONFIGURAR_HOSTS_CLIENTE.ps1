param(
    [string]$ServerIp = "192.168.1.10",
    [string[]]$Domains = @("volume.local", "ponto.local", "ponto.admin", "util.local")
)

$ErrorActionPreference = "Stop"
$hostsPath = Join-Path $env:windir "System32\drivers\etc\hosts"

$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "Execute este script como Administrador." -ForegroundColor Yellow
    exit 1
}

if (-not [System.Net.IPAddress]::TryParse($ServerIp, [ref]([System.Net.IPAddress]$null))) {
    Write-Host "IP invalido: $ServerIp" -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $hostsPath)) {
    Write-Host "Arquivo hosts nao encontrado: $hostsPath" -ForegroundColor Red
    exit 1
}

$lines = Get-Content -Path $hostsPath -ErrorAction Stop
$updated = @()

foreach ($line in $lines) {
    $skip = $false
    foreach ($domain in $Domains) {
        if ($line -match "(^|\s)" + [regex]::Escape($domain) + "(\s|$)") {
            $skip = $true
            break
        }
    }
    if (-not $skip) {
        $updated += $line
    }
}

$updated += ""
$updated += "# Local domains - managed by CONFIGURAR_HOSTS_CLIENTE.ps1"
foreach ($domain in $Domains) {
    $updated += "$ServerIp $domain"
}

Set-Content -Path $hostsPath -Value $updated -Encoding ASCII
Write-Host "Hosts atualizado com sucesso." -ForegroundColor Green
Write-Host "Entradas aplicadas:" -ForegroundColor Cyan
foreach ($domain in $Domains) {
    Write-Host "  $ServerIp $domain"
}

ipconfig /flushdns | Out-Null
Write-Host "Cache DNS limpo (ipconfig /flushdns)." -ForegroundColor Green

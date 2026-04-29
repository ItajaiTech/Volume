param(
    [string]$HostName = "volume.local",
    [string]$TargetIp = "127.0.0.1"
)

$hostsPath = "C:\Windows\System32\drivers\etc\hosts"
$entry = "$TargetIp $HostName"

if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Execute este script como Administrador para alterar o arquivo hosts." -ForegroundColor Yellow
    Write-Host "Entrada desejada: $entry" -ForegroundColor Yellow
    exit 1
}

if (-not (Test-Path $hostsPath)) {
    Write-Host "Arquivo hosts nao encontrado: $hostsPath" -ForegroundColor Red
    exit 1
}

$content = Get-Content $hostsPath -ErrorAction Stop
$pattern = "(^|\s)" + [regex]::Escape($HostName) + "($|\s)"

if ($content | Select-String -Pattern $pattern) {
    Write-Host "$HostName ja possui entrada no hosts. Revise manualmente se quiser trocar o IP." -ForegroundColor Yellow
    exit 0
}

Add-Content -Path $hostsPath -Value "`r`n$entry"
Write-Host "Entrada adicionada com sucesso: $entry" -ForegroundColor Green
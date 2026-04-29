# Inicia RelogioPonto e Volume em janelas separadas e ambientes isolados
$relogioScript = "C:\RelogioPonto\Ponto\start_relogioponto.ps1"
$workspaceRoot = $PSScriptRoot
if (-not $workspaceRoot) {
    $workspaceRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
}
$volumeScript = Join-Path $workspaceRoot "shipping_ai\start_volume.ps1"
$relogioStarted = $false
$volumeStarted = $false

if (-not (Test-Path $relogioScript)) {
    Write-Host "Script do RelogioPonto nao encontrado: $relogioScript" -ForegroundColor Yellow
} else {
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$relogioScript`"" -WorkingDirectory (Split-Path -Parent $relogioScript)
    $relogioStarted = $true
}

if (-not (Test-Path $volumeScript)) {
    Write-Host "Script do Volume nao encontrado: $volumeScript" -ForegroundColor Red
} else {
    Start-Sleep -Seconds 2
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$volumeScript`"" -WorkingDirectory (Split-Path -Parent $volumeScript)
    $volumeStarted = $true
}

if (-not $relogioStarted -and -not $volumeStarted) {
    Write-Host "Nenhuma aplicacao foi iniciada." -ForegroundColor Red
    exit 1
}

if ($relogioStarted -and $volumeStarted) {
    Write-Host "Aplicacoes iniciadas em processos separados." -ForegroundColor Green
} elseif ($volumeStarted) {
    Write-Host "Volume iniciado. RelogioPonto nao foi iniciado." -ForegroundColor Yellow
} else {
    Write-Host "RelogioPonto iniciado. Volume nao foi iniciado." -ForegroundColor Yellow
}

if ($relogioStarted) {
    Write-Host "RelogioPonto: https://ponto.local:5000 (ou https://ponto.admin:5050)"
}
if ($volumeStarted) {
    Write-Host "Volume: http://volume.local:6100"
}

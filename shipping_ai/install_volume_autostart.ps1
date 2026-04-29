# Registra uma tarefa agendada para manter o Volume disponivel apos reboot e logon.
$taskName = "VolumeAppBackground"
$appPath = $PSScriptRoot
if (-not $appPath) {
    $appPath = Split-Path -Parent $MyInvocation.MyCommand.Path
}
$scriptPath = Join-Path $appPath "run_volume_background.ps1"

if (-not (Test-Path $scriptPath)) {
    Write-Host "Script de background nao encontrado: $scriptPath" -ForegroundColor Red
    exit 1
}

$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`""
$startupTrigger = New-ScheduledTaskTrigger -AtStartup
$logonTrigger = New-ScheduledTaskTrigger -AtLogOn
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -MultipleInstances IgnoreNew `
    -RestartCount 999 `
    -RestartInterval (New-TimeSpan -Minutes 1) `
    -StartWhenAvailable

$registeredMode = $null

function Start-VolumeBackgroundNow {
    Start-Process powershell.exe -WindowStyle Hidden -ArgumentList @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", $scriptPath
    ) | Out-Null
}

try {
    $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
    Register-ScheduledTask `
        -TaskName $taskName `
        -Action $action `
        -Trigger @($startupTrigger, $logonTrigger) `
        -Settings $settings `
        -Principal $principal `
        -Force `
        -ErrorAction Stop | Out-Null
    $registeredMode = "system"
} catch {
    Write-Host "Sem permissao para registrar como SYSTEM. Aplicando fallback para o usuario atual..." -ForegroundColor Yellow

    try {
        $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        $principal = New-ScheduledTaskPrincipal -UserId $currentUser -LogonType Interactive -RunLevel Limited

        Register-ScheduledTask `
            -TaskName $taskName `
            -Action $action `
            -Trigger $logonTrigger `
            -Settings $settings `
            -Principal $principal `
            -Force `
            -ErrorAction Stop | Out-Null
        $registeredMode = "user"
    } catch {
        Write-Host "Agendador indisponivel para este usuario. Aplicando fallback via inicializacao do Windows..." -ForegroundColor Yellow

        $runKeyPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
        $runValue = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`""
        New-Item -Path $runKeyPath -Force | Out-Null
        Set-ItemProperty -Path $runKeyPath -Name $taskName -Value $runValue -Force
        $registeredMode = "registry"
    }
}

if (-not $registeredMode) {
    Write-Host "Nao foi possivel registrar a tarefa agendada do Volume." -ForegroundColor Red
    exit 1
}

if ($registeredMode -in @("system", "user")) {
    Start-ScheduledTask -TaskName $taskName -ErrorAction Stop
} else {
    Start-VolumeBackgroundNow
}

if ($registeredMode -eq "system") {
    Write-Host "Tarefa agendada registrada com sucesso: $taskName" -ForegroundColor Green
    Write-Host "O Volume sera iniciado automaticamente no boot e no logon." -ForegroundColor Green
} elseif ($registeredMode -eq "user") {
    Write-Host "Tarefa agendada registrada com sucesso para o usuario atual: $taskName" -ForegroundColor Green
    Write-Host "O Volume sera iniciado automaticamente quando este usuario fizer logon." -ForegroundColor Yellow
} elseif ($registeredMode -eq "registry") {
    Write-Host "Inicializacao automatica registrada com sucesso no perfil do usuario." -ForegroundColor Green
    Write-Host "O Volume sera iniciado automaticamente quando este usuario fizer logon." -ForegroundColor Yellow
}
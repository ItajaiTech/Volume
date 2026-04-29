$taskName = "VolumePontoReverseProxy"
$proxyDir = $PSScriptRoot
$scriptPath = Join-Path $proxyDir "start_reverse_proxy.ps1"

if (-not (Test-Path $scriptPath)) {
    Write-Host "Script de start do proxy nao encontrado: $scriptPath" -ForegroundColor Red
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

try {
    $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger @($startupTrigger, $logonTrigger) -Settings $settings -Principal $principal -Force -ErrorAction Stop | Out-Null
    $registeredMode = "system"
} catch {
    try {
        $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        $principal = New-ScheduledTaskPrincipal -UserId $currentUser -LogonType Interactive -RunLevel Limited
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $logonTrigger -Settings $settings -Principal $principal -Force -ErrorAction Stop | Out-Null
        $registeredMode = "user"
    } catch {
        $runKeyPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
        $runValue = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`""
        New-Item -Path $runKeyPath -Force | Out-Null
        Set-ItemProperty -Path $runKeyPath -Name $taskName -Value $runValue -Force
        $registeredMode = "registry"
    }
}

if (-not $registeredMode) {
    Write-Host "Nao foi possivel registrar autostart do proxy." -ForegroundColor Red
    exit 1
}

if ($registeredMode -in @("system", "user")) {
    Start-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
} else {
    Start-Process powershell.exe -WindowStyle Hidden -ArgumentList @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $scriptPath) | Out-Null
}

Write-Host "Autostart do proxy configurado no modo: $registeredMode" -ForegroundColor Green

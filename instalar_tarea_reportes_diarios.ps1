$ErrorActionPreference = "Stop"

$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$pythonExe = "python"
$scriptPath = Join-Path $baseDir "enviar_reportes_diarios.py"

if (-not (Test-Path $scriptPath)) {
    throw "No se encontro el script: $scriptPath"
}

# Ejecuta todos los dias a las 07:00.
$taskName = "WES_Envio_Reportes_Diarios"
$action = New-ScheduledTaskAction -Execute $pythonExe -Argument "`"$scriptPath`"" -WorkingDirectory $baseDir
$trigger = New-ScheduledTaskTrigger -Daily -At 07:00
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

Register-ScheduledTask `
    -TaskName $taskName `
    -Action $action `
    -Trigger $trigger `
    -Principal $principal `
    -Settings $settings `
    -Force | Out-Null

Write-Host "[OK] Tarea instalada/actualizada: $taskName"
Write-Host "[OK] Hora: 07:00 diaria"
Write-Host "[OK] Script: $scriptPath"


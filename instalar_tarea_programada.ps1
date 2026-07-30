# Script PowerShell para instalar tarea programada de monitoreo de correos WES
# Usa el Programador de Tareas de Windows (no requiere NSSM)
# Debe ejecutarse como Administrador

$ErrorActionPreference = "Stop"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "INSTALACION DE TAREA PROGRAMADA WES" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Verificar que se ejecuta como Administrador
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "[ERROR] Este script debe ejecutarse como Administrador" -ForegroundColor Red
    Write-Host "Haz clic derecho y selecciona 'Ejecutar como administrador'" -ForegroundColor Yellow
    pause
    exit 1
}

# Parámetros
$taskName = "WESMonitoreoCorreos"
$taskDescription = "Monitorea correos automaticamente y genera reportes de consumo de agua"

# Rutas
$scriptPath = "C:\Users\joseo\Desktop\wes-scripts\monitorear_correos_y_generar_reportes.py"
$pythonExe = "C:\Users\joseo\AppData\Local\Programs\Python\Python314\python.exe"
$workingDir = "C:\Users\joseo\Desktop\wes-scripts"

# Verificar que existen los archivos
if (-not (Test-Path $scriptPath)) {
    Write-Host "[ERROR] No se encontro el script: $scriptPath" -ForegroundColor Red
    pause
    exit 1
}

if (-not (Test-Path $pythonExe)) {
    Write-Host "[ERROR] No se encontro Python: $pythonExe" -ForegroundColor Red
    pause
    exit 1
}

Write-Host "[INFO] Instalando tarea programada..." -ForegroundColor Cyan
Write-Host "  Nombre: $taskName" -ForegroundColor White
Write-Host "  Python: $pythonExe" -ForegroundColor White
Write-Host "  Script: $scriptPath" -ForegroundColor White
Write-Host ""

# Eliminar tarea si ya existe
$existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($existingTask) {
    Write-Host "[INFO] La tarea ya existe. Eliminando version anterior..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    Start-Sleep -Seconds 2
}

# Crear accion (comando a ejecutar)
$action = New-ScheduledTaskAction -Execute $pythonExe -Argument "`"$scriptPath`" --continuo --intervalo 5" -WorkingDirectory $workingDir

# Crear trigger (inicio al arrancar el sistema)
$trigger = New-ScheduledTaskTrigger -AtStartup

# Configurar para ejecutar aunque el usuario no este conectado
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType ServiceAccount -RunLevel Highest

# Configuracion adicional
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)

# Registrar la tarea
try {
    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description $taskDescription | Out-Null
    
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "TAREA INSTALADA EXITOSAMENTE" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "La tarea '$taskName' esta configurada para:" -ForegroundColor Green
    Write-Host "  - Iniciarse automaticamente al encender el PC" -ForegroundColor White
    Write-Host "  - Ejecutarse en segundo plano" -ForegroundColor White
    Write-Host "  - Monitorear correos cada 5 minutos" -ForegroundColor White
    Write-Host "  - Reiniciarse automaticamente si falla" -ForegroundColor White
    Write-Host ""
    Write-Host "Comandos utiles:" -ForegroundColor Cyan
    Write-Host "  Iniciar:   Start-ScheduledTask -TaskName `"$taskName`"" -ForegroundColor White
    Write-Host "  Detener:   Stop-ScheduledTask -TaskName `"$taskName`"" -ForegroundColor White
    Write-Host "  Estado:    Get-ScheduledTask -TaskName `"$taskName`"" -ForegroundColor White
    Write-Host "  Eliminar:  Unregister-ScheduledTask -TaskName `"$taskName`" -Confirm:`$false" -ForegroundColor White
    Write-Host ""
    Write-Host "Iniciando la tarea ahora..." -ForegroundColor Cyan
    Start-ScheduledTask -TaskName $taskName
    Start-Sleep -Seconds 2
    
    $task = Get-ScheduledTask -TaskName $taskName
    if ($task.State -eq "Running") {
        Write-Host "[OK] Tarea iniciada y corriendo" -ForegroundColor Green
    } else {
        Write-Host "[INFO] Tarea registrada. Se iniciara al reiniciar el PC" -ForegroundColor Yellow
    }
    
} catch {
    Write-Host "[ERROR] Fallo la instalacion" -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
    pause
    exit 1
}

Write-Host ""
pause

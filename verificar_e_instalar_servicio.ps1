# Script simplificado para verificar e instalar el servicio
# Ejecutar como Administrador

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "VERIFICACION E INSTALACION DE SERVICIO" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Verificar permisos de administrador
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "[ERROR] Debes ejecutar como Administrador" -ForegroundColor Red
    Write-Host "Haz clic derecho -> Ejecutar como administrador" -ForegroundColor Yellow
    pause
    exit 1
}

$taskName = "WESMonitoreoCorreos"
$scriptPath = "C:\Users\joseo\Desktop\wes-scripts\monitorear_correos_y_generar_reportes.py"
$pythonExe = "C:\Users\joseo\AppData\Local\Programs\Python\Python314\python.exe"
$workingDir = "C:\Users\joseo\Desktop\wes-scripts"

# Verificar archivos
Write-Host "Verificando archivos..." -ForegroundColor Cyan
if (-not (Test-Path $scriptPath)) {
    Write-Host "[ERROR] Script no encontrado: $scriptPath" -ForegroundColor Red
    pause
    exit 1
}
Write-Host "  [OK] Script encontrado" -ForegroundColor Green

if (-not (Test-Path $pythonExe)) {
    Write-Host "[ERROR] Python no encontrado: $pythonExe" -ForegroundColor Red
    pause
    exit 1
}
Write-Host "  [OK] Python encontrado" -ForegroundColor Green
Write-Host ""

# Verificar si la tarea existe
Write-Host "Verificando tarea existente..." -ForegroundColor Cyan
$existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($existingTask) {
    Write-Host "  [INFO] Tarea existe. Estado: $($existingTask.State)" -ForegroundColor Yellow
    Write-Host "  Eliminando para reinstalar..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
}

# Crear tarea
Write-Host "Creando tarea programada..." -ForegroundColor Cyan
try {
    $action = New-ScheduledTaskAction -Execute $pythonExe -Argument "`"$scriptPath`" --continuo --intervalo 5" -WorkingDirectory $workingDir
    $trigger = New-ScheduledTaskTrigger -AtStartup
    $principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType ServiceAccount -RunLevel Highest
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)
    
    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description "Monitorea correos y genera reportes WES" | Out-Null
    
    Write-Host "  [OK] Tarea creada exitosamente" -ForegroundColor Green
    Write-Host ""
    
    # Iniciar tarea
    Write-Host "Iniciando tarea..." -ForegroundColor Cyan
    Start-ScheduledTask -TaskName $taskName
    Start-Sleep -Seconds 3
    
    # Verificar estado
    $task = Get-ScheduledTask -TaskName $taskName
    Write-Host "  Estado: $($task.State)" -ForegroundColor $(if ($task.State -eq "Running") { "Green" } else { "Yellow" })
    
    if ($task.State -eq "Running") {
        Write-Host ""
        Write-Host "========================================" -ForegroundColor Green
        Write-Host "SERVICIO INSTALADO Y CORRIENDO" -ForegroundColor Green
        Write-Host "========================================" -ForegroundColor Green
        Write-Host ""
        Write-Host "El servicio esta corriendo en segundo plano" -ForegroundColor Green
        Write-Host "Monitoreara correos cada 5 minutos automaticamente" -ForegroundColor Green
    } else {
        Write-Host ""
        Write-Host "[INFO] Tarea instalada. Se iniciara al reiniciar el PC" -ForegroundColor Yellow
    }
    
} catch {
    Write-Host "[ERROR] Error al crear tarea: $_" -ForegroundColor Red
    pause
    exit 1
}

Write-Host ""
pause








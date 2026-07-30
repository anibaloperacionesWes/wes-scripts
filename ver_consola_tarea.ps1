# Script para ver la consola/salida de la tarea programada de monitoreo
# Muestra los logs en tiempo real

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "CONSOLA DE MONITOREO DE CORREOS WES" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

$logFile = "logs\monitoreo_correos.log"

if (-not (Test-Path $logFile)) {
    Write-Host "[ADVERTENCIA] No se encontró el archivo de log: $logFile" -ForegroundColor Yellow
    Write-Host "[INFO] La tarea puede no haber iniciado aún o no estar configurada para generar logs" -ForegroundColor Yellow
    Write-Host ""
    
    # Verificar si la tarea está corriendo
    $task = Get-ScheduledTask -TaskName "WESMonitoreoCorreos" -ErrorAction SilentlyContinue
    if ($task) {
        Write-Host "Estado de la tarea: $($task.State)" -ForegroundColor White
    } else {
        Write-Host "[ERROR] La tarea WESMonitoreoCorreos no existe" -ForegroundColor Red
    }
    
    # Verificar procesos Python
    $processes = Get-Process python -ErrorAction SilentlyContinue | Where-Object {$_.Path -like "*Python314*"}
    if ($processes) {
        Write-Host ""
        Write-Host "Procesos Python corriendo:" -ForegroundColor Green
        $processes | ForEach-Object {
            Write-Host "  - PID: $($_.Id), Inicio: $($_.StartTime)" -ForegroundColor White
        }
    } else {
        Write-Host ""
        Write-Host "[INFO] No hay procesos Python corriendo" -ForegroundColor Yellow
    }
    
    Write-Host ""
    Write-Host "Opciones:" -ForegroundColor Cyan
    Write-Host "1. Ejecutar el script manualmente para ver la salida en tiempo real" -ForegroundColor White
    Write-Host "2. Verificar que la tarea esté configurada correctamente" -ForegroundColor White
    Write-Host ""
    pause
    exit
}

Write-Host "[INFO] Mostrando últimas 50 líneas del log..." -ForegroundColor Yellow
Write-Host "[INFO] Presiona Ctrl+C para salir" -ForegroundColor Yellow
Write-Host ""
Write-Host "=" * 70 -ForegroundColor Gray
Write-Host ""

# Mostrar últimas líneas
Get-Content $logFile -Tail 50

Write-Host ""
Write-Host "=" * 70 -ForegroundColor Gray
Write-Host ""
Write-Host "[INFO] Para ver el log en tiempo real, ejecuta:" -ForegroundColor Cyan
Write-Host "  Get-Content logs\monitoreo_correos.log -Wait -Tail 20" -ForegroundColor White
Write-Host ""
Write-Host "[INFO] O ejecuta el script manualmente para ver la salida:" -ForegroundColor Cyan
Write-Host "  python monitorear_correos_y_generar_reportes.py --continuo --intervalo 5" -ForegroundColor White
Write-Host ""

pause




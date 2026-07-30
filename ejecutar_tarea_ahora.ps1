Write-Host "Forzando ejecucion inmediata de la tarea..." -ForegroundColor Yellow
$tarea = Get-ScheduledTask | Where-Object { $_.TaskName -like '*WESMonitoreoCorreos*' }
if ($tarea) {
    # Detener la tarea si está corriendo
    if ($tarea.State -eq 'Running') {
        Write-Host "Deteniendo tarea actual..." -ForegroundColor Yellow
        Stop-ScheduledTask -TaskName $tarea.TaskName -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
    }
    
    # Iniciar la tarea inmediatamente
    Write-Host "Iniciando tarea ahora..." -ForegroundColor Yellow
    Start-ScheduledTask -TaskName $tarea.TaskName
    Start-Sleep -Seconds 2
    
    $estado = (Get-ScheduledTask -TaskName $tarea.TaskName).State
    Write-Host "Estado: $estado" -ForegroundColor Green
    Write-Host "[OK] La tarea se esta ejecutando ahora" -ForegroundColor Green
    Write-Host ""
    Write-Host "Monitoreando logs en tiempo real (presiona Ctrl+C para salir)..." -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor Gray
    
    # Mostrar logs en tiempo real
    Get-Content logs\monitoreo_correos.log -Wait -Tail 5
} else {
    Write-Host "No se encontro la tarea WESMonitoreoCorreos" -ForegroundColor Red
}

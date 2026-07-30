$tarea = Get-ScheduledTask | Where-Object { $_.TaskName -like '*WESMonitoreoCorreos*' }
if ($tarea) {
    Write-Host "Forzando ejecucion inmediata..." -ForegroundColor Yellow
    
    # Detener si está corriendo
    if ($tarea.State -eq 'Running') {
        Write-Host "Deteniendo tarea actual..." -ForegroundColor Yellow
        Stop-ScheduledTask -TaskName $tarea.TaskName -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
    }
    
    # Iniciar inmediatamente
    Write-Host "Iniciando tarea ahora..." -ForegroundColor Yellow
    Start-ScheduledTask -TaskName $tarea.TaskName
    Start-Sleep -Seconds 2
    
    $estado = (Get-ScheduledTask -TaskName $tarea.TaskName).State
    Write-Host "[OK] Tarea ejecutada. Estado: $estado" -ForegroundColor Green
} else {
    Write-Host "No se encontro la tarea" -ForegroundColor Red
}

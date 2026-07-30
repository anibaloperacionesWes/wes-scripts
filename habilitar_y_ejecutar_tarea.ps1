$tarea = Get-ScheduledTask | Where-Object { $_.TaskName -like '*WESMonitoreoCorreos*' }
if ($tarea) {
    Write-Host "Habilitando tarea..." -ForegroundColor Yellow
    Enable-ScheduledTask -TaskName $tarea.TaskName
    Start-Sleep -Seconds 1
    
    Write-Host "Deteniendo tarea si esta corriendo..." -ForegroundColor Yellow
    if ($tarea.State -eq 'Running') {
        Stop-ScheduledTask -TaskName $tarea.TaskName -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
    }
    
    Write-Host "Iniciando tarea ahora..." -ForegroundColor Yellow
    Start-ScheduledTask -TaskName $tarea.TaskName
    Start-Sleep -Seconds 2
    
    $estado = (Get-ScheduledTask -TaskName $tarea.TaskName).State
    Write-Host "[OK] Tarea habilitada y ejecutada. Estado: $estado" -ForegroundColor Green
} else {
    Write-Host "No se encontro la tarea" -ForegroundColor Red
}

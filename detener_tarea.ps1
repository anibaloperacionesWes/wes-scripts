Write-Host "Deteniendo la tarea..." -ForegroundColor Yellow
$tarea = Get-ScheduledTask | Where-Object { $_.TaskName -like '*WESMonitoreoCorreos*' }
if ($tarea) {
    if ($tarea.State -eq 'Running') {
        Write-Host "La tarea esta corriendo. Deteniendo..." -ForegroundColor Yellow
        Stop-ScheduledTask -TaskName $tarea.TaskName
        Start-Sleep -Seconds 2
        $estado = (Get-ScheduledTask -TaskName $tarea.TaskName).State
        Write-Host "[OK] Tarea detenida. Estado actual: $estado" -ForegroundColor Green
    } else {
        Write-Host "La tarea no esta corriendo. Estado: $($tarea.State)" -ForegroundColor Yellow
    }
} else {
    Write-Host "No se encontro la tarea WESMonitoreoCorreos" -ForegroundColor Red
}

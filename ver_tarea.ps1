$tarea = Get-ScheduledTask | Where-Object { $_.TaskName -like '*WESMonitoreoCorreos*' }
if ($tarea) {
    $info = Get-ScheduledTaskInfo -TaskName $tarea.TaskName
    
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  INFORMACION DE LA TAREA" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    
    Write-Host "Nombre: " -NoNewline
    Write-Host $tarea.TaskName -ForegroundColor Cyan
    
    Write-Host "Estado: " -NoNewline
    $colorEstado = if ($tarea.State -eq 'Running') { 'Green' } elseif ($tarea.State -eq 'Ready') { 'Yellow' } else { 'Red' }
    Write-Host $tarea.State -ForegroundColor $colorEstado
    
    Write-Host "Ultima ejecucion: " -NoNewline
    Write-Host $info.LastRunTime -ForegroundColor Cyan
    
    Write-Host "Proxima ejecucion: " -NoNewline
    Write-Host $info.NextRunTime -ForegroundColor Cyan
    
    Write-Host "Ultimo resultado: " -NoNewline
    if ($info.LastTaskResult -eq 0) {
        Write-Host "Exitoso (0x0)" -ForegroundColor Green
    } else {
        $resultadoHex = "0x{0:X}" -f $info.LastTaskResult
        Write-Host "Error ($resultadoHex)" -ForegroundColor Red
    }
    
    Write-Host ""
    Write-Host "Accion: " -NoNewline
    Write-Host $tarea.Actions.Execute -ForegroundColor Cyan
    
    Write-Host "Argumentos: " -NoNewline
    Write-Host $tarea.Actions.Arguments -ForegroundColor Cyan
    
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
} else {
    Write-Host "No se encontro la tarea WESMonitoreoCorreos" -ForegroundColor Red
}

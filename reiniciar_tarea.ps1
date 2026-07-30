# Script para reiniciar la tarea de Windows de monitoreo de correos
# Ejecutar como administrador

Write-Host "Buscando tareas relacionadas..." -ForegroundColor Yellow

# Buscar tareas
$tareas = Get-ScheduledTask | Where-Object {
    $taskName = $_.TaskName.ToLower()
    $taskName -like '*correo*' -or 
    $taskName -like '*reporte*' -or 
    $taskName -like '*monitoreo*' -or 
    $taskName -like '*wes*' -or
    $taskName -like '*monitorearcorreos*'
}

# Si no se encuentran por nombre, buscar por accion
if ($tareas.Count -eq 0) {
    $allTasks = Get-ScheduledTask
    foreach ($task in $allTasks) {
        $action = $task.Actions.Execute
        if ($action -like '*python*' -and ($task.Actions.Arguments -like '*monitorear*' -or $task.Actions.Arguments -like '*correo*')) {
            $tareas = @($task)
            break
        }
    }
}

if ($tareas.Count -eq 0) {
    Write-Host "No se encontraron tareas. Listando todas:" -ForegroundColor Red
    Get-ScheduledTask | Format-Table TaskName, State -AutoSize
    exit 1
}

# Reiniciar cada tarea encontrada
foreach ($tarea in $tareas) {
    Write-Host "Tarea encontrada: $($tarea.TaskName)" -ForegroundColor Green
    Write-Host "Estado actual: $($tarea.State)" -ForegroundColor Cyan
    
    try {
        # Detener si esta en ejecucion
        if ($tarea.State -eq 'Running') {
            Write-Host "Deteniendo tarea..." -ForegroundColor Yellow
            Stop-ScheduledTask -TaskName $tarea.TaskName -ErrorAction Stop
            Start-Sleep -Seconds 3
        }
        
        # Iniciar la tarea
        Write-Host "Iniciando tarea..." -ForegroundColor Yellow
        Start-ScheduledTask -TaskName $tarea.TaskName -ErrorAction Stop
        Start-Sleep -Seconds 2
        
        # Verificar estado
        $estado = (Get-ScheduledTask -TaskName $tarea.TaskName).State
        Write-Host "Tarea reiniciada. Nuevo estado: $estado" -ForegroundColor Green
        
    } catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host "Proceso completado." -ForegroundColor Green






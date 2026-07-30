# Script para verificar y reiniciar la tarea de Windows de monitoreo de correos
# Ejecutar como administrador: powershell -ExecutionPolicy Bypass -File verificar_y_reiniciar_tarea.ps1

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  VERIFICACIÓN Y REINICIO DE TAREA" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Buscar todas las tareas programadas
Write-Host "Buscando tareas relacionadas con monitoreo de correos..." -ForegroundColor Yellow
$tareas = Get-ScheduledTask | Where-Object {
    $taskName = $_.TaskName.ToLower()
    $taskName -like '*correo*' -or 
    $taskName -like '*reporte*' -or 
    $taskName -like '*monitoreo*' -or 
    $taskName -like '*wes*' -or
    $taskName -like '*monitorearcorreos*'
}

if ($tareas.Count -eq 0) {
    Write-Host "[INFO] No se encontraron tareas con nombres relacionados." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Listando todas las tareas programadas:" -ForegroundColor Yellow
    $allTasks = Get-ScheduledTask
    foreach ($task in $allTasks) {
        $action = $task.Actions.Execute
        if ($action -like '*python*' -or $action -like '*monitorear*' -or $action -like '*correo*') {
            Write-Host "  - $($task.TaskName): $($task.State) [Acción: $action]" -ForegroundColor Cyan
            $tareas += $task
        }
    }
    Write-Host ""
}

if ($tareas.Count -eq 0) {
    Write-Host "[ADVERTENCIA] No se encontró ninguna tarea relacionada." -ForegroundColor Red
    Write-Host ""
    Write-Host "Por favor, verifica manualmente el nombre de la tarea en:" -ForegroundColor Yellow
    Write-Host "  - Programador de tareas de Windows (taskschd.msc)" -ForegroundColor Yellow
    Write-Host "  - O ejecuta: Get-ScheduledTask | Format-Table TaskName, State" -ForegroundColor Yellow
    exit 1
}

Write-Host "Tareas encontradas:" -ForegroundColor Green
foreach ($tarea in $tareas) {
    Write-Host "  - $($tarea.TaskName): $($tarea.State)" -ForegroundColor Cyan
}
Write-Host ""

# Reiniciar cada tarea encontrada
foreach ($tarea in $tareas) {
    Write-Host "Reiniciando tarea: $($tarea.TaskName)..." -ForegroundColor Yellow
    
    try {
        # Detener la tarea si está en ejecución
        if ($tarea.State -eq 'Running') {
            Write-Host "  [1/3] Deteniendo tarea..." -ForegroundColor Yellow
            Stop-ScheduledTask -TaskName $tarea.TaskName -ErrorAction Stop
            Write-Host "  [OK] Tarea detenida" -ForegroundColor Green
            Start-Sleep -Seconds 3
        } else {
            Write-Host "  [INFO] La tarea no estaba en ejecución" -ForegroundColor Gray
        }
        
        # Iniciar la tarea
        Write-Host "  [2/3] Iniciando tarea..." -ForegroundColor Yellow
        Start-ScheduledTask -TaskName $tarea.TaskName -ErrorAction Stop
        Write-Host "  [OK] Tarea iniciada" -ForegroundColor Green
        
        Start-Sleep -Seconds 2
        
        # Verificar estado
        Write-Host "  [3/3] Verificando estado..." -ForegroundColor Yellow
        $estado = (Get-ScheduledTask -TaskName $tarea.TaskName).State
        Write-Host "  [OK] Estado actual: $estado" -ForegroundColor Green
        
        Write-Host ""
        Write-Host "[OK] Tarea reiniciada exitosamente: $($tarea.TaskName)" -ForegroundColor Green
        Write-Host ""
    }
    catch {
        Write-Host "[ERROR] No se pudo reiniciar la tarea: $($tarea.TaskName)" -ForegroundColor Red
        Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""
    }
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  PROCESO COMPLETADO" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "El código de generar_reporte_word.py está actualizado con:" -ForegroundColor Green
Write-Host "  - Gráfica de anillo: 4.5 pulgadas" -ForegroundColor Gray
Write-Host "  - Título: fontsize=14, pad=5" -ForegroundColor Gray
Write-Host "  - Leyenda: bbox_to_anchor=(0.5, 0.02)" -ForegroundColor Gray
Write-Host "  - Recomendación en conclusiones cuando hay análisis de filtración" -ForegroundColor Gray
Write-Host ""







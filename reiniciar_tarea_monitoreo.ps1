# Script para verificar y reiniciar la tarea de Windows de monitoreo de correos

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  VERIFICACIÓN Y REINICIO DE TAREA" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Buscar tareas relacionadas con monitoreo, correos o reportes
Write-Host "Buscando tareas relacionadas..." -ForegroundColor Yellow
$tareas = Get-ScheduledTask | Where-Object {
    $_.TaskName -like '*correo*' -or 
    $_.TaskName -like '*reporte*' -or 
    $_.TaskName -like '*monitoreo*' -or 
    $_.TaskName -like '*wes*' -or
    $_.TaskName -like '*MonitoreoCorreos*' -or
    $_.TaskName -like '*MonitorearCorreos*'
}

if ($tareas.Count -eq 0) {
    Write-Host "No se encontraron tareas relacionadas." -ForegroundColor Red
    Write-Host ""
    Write-Host "Listando todas las tareas programadas:" -ForegroundColor Yellow
    Get-ScheduledTask | Format-Table TaskName, State -AutoSize
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
            Write-Host "  Deteniendo tarea..." -ForegroundColor Yellow
            Stop-ScheduledTask -TaskName $tarea.TaskName -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
        }
        
        # Iniciar la tarea
        Write-Host "  Iniciando tarea..." -ForegroundColor Yellow
        Start-ScheduledTask -TaskName $tarea.TaskName
        
        Start-Sleep -Seconds 1
        
        # Verificar estado
        $estado = (Get-ScheduledTask -TaskName $tarea.TaskName).State
        Write-Host "  Estado actual: $estado" -ForegroundColor Green
        
        Write-Host "[OK] Tarea reiniciada exitosamente: $($tarea.TaskName)" -ForegroundColor Green
    }
    catch {
        Write-Host "[ERROR] No se pudo reiniciar la tarea: $($tarea.TaskName)" -ForegroundColor Red
        Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host ""
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  PROCESO COMPLETADO" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

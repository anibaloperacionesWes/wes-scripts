# Script para verificar el estado de la tarea de Windows de monitoreo de correos

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  VERIFICACIÓN DE ESTADO DE TAREA" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Buscar tareas relacionadas
Write-Host "Buscando tareas relacionadas con monitoreo de correos..." -ForegroundColor Yellow
$tareas = Get-ScheduledTask | Where-Object {
    $taskName = $_.TaskName.ToLower()
    $taskName -like '*correo*' -or 
    $taskName -like '*reporte*' -or 
    $taskName -like '*monitoreo*' -or 
    $taskName -like '*wes*'
}

# Si no se encuentran por nombre, buscar por acción
if ($tareas.Count -eq 0) {
    Write-Host "Buscando por acción ejecutada..." -ForegroundColor Yellow
    $tareas = Get-ScheduledTask | Where-Object {
        $action = $_.Actions.Execute
        $args = $_.Actions.Arguments
        ($action -like '*python*' -and ($args -like '*monitorear*' -or $args -like '*correo*')) -or
        ($action -like '*monitorear_correos*')
    }
}

if ($tareas.Count -eq 0) {
    Write-Host "[ADVERTENCIA] No se encontraron tareas relacionadas." -ForegroundColor Red
    Write-Host ""
    Write-Host "Listando todas las tareas programadas:" -ForegroundColor Yellow
    Get-ScheduledTask | Format-Table TaskName, State -AutoSize
    exit 1
}

Write-Host "Tareas encontradas: $($tareas.Count)" -ForegroundColor Green
Write-Host ""

foreach ($tarea in $tareas) {
    $info = Get-ScheduledTaskInfo -TaskName $tarea.TaskName
    
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Tarea: $($tarea.TaskName)" -ForegroundColor White
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Estado: " -NoNewline
    
    if ($tarea.State -eq 'Running') {
        Write-Host "$($tarea.State)" -ForegroundColor Green
    } elseif ($tarea.State -eq 'Ready') {
        Write-Host "$($tarea.State) (Lista para ejecutarse)" -ForegroundColor Yellow
    } else {
        Write-Host "$($tarea.State)" -ForegroundColor Red
    }
    
    Write-Host "Última ejecución: " -NoNewline
    if ($info.LastRunTime) {
        Write-Host "$($info.LastRunTime)" -ForegroundColor Cyan
    } else {
        Write-Host "Nunca ejecutada" -ForegroundColor Gray
    }
    
    Write-Host "Próxima ejecución: " -NoNewline
    if ($info.NextRunTime) {
        Write-Host "$($info.NextRunTime)" -ForegroundColor Cyan
    } else {
        Write-Host "No programada" -ForegroundColor Gray
    }
    
    Write-Host "Último resultado: " -NoNewline
    if ($info.LastTaskResult -eq 0) {
        Write-Host "Exitoso (0x0)" -ForegroundColor Green
    } elseif ($info.LastTaskResult) {
        Write-Host "Error (0x$($info.LastTaskResult.ToString('X')))" -ForegroundColor Red
    } else {
        Write-Host "N/A" -ForegroundColor Gray
    }
    
    Write-Host "Acción: $($tarea.Actions.Execute) $($tarea.Actions.Arguments)" -ForegroundColor Gray
    Write-Host ""
}

# Verificar procesos de Python relacionados
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  PROCESOS DE PYTHON EN EJECUCIÓN" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

$pythonProcesses = Get-Process python -ErrorAction SilentlyContinue | Where-Object {
    $_.Path -like '*wes-scripts*' -or 
    $_.CommandLine -like '*monitorear*' -or
    $_.CommandLine -like '*correo*'
}

if ($pythonProcesses) {
    Write-Host "Procesos encontrados: $($pythonProcesses.Count)" -ForegroundColor Green
    foreach ($proc in $pythonProcesses) {
        Write-Host "  - PID: $($proc.Id) | Ruta: $($proc.Path)" -ForegroundColor Cyan
    }
} else {
    Write-Host "No se encontraron procesos de Python relacionados." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  VERIFICACIÓN COMPLETADA" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan







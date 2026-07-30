# Script para abrir el proceso de monitoreo en una ventana visible

$taskName = "WESMonitoreoCorreos"

Write-Host "Obteniendo informacion de la tarea..." -ForegroundColor Yellow

try {
    $task = Get-ScheduledTask -TaskName $taskName -ErrorAction Stop
    $action = $task.Actions[0]
    
    $pythonExe = $action.Execute
    $arguments = $action.Arguments
    
    Write-Host "Python: $pythonExe" -ForegroundColor Cyan
    Write-Host "Argumentos: $arguments" -ForegroundColor Cyan
    Write-Host ""
    
    Write-Host "Abriendo ventana de Python..." -ForegroundColor Green
    
    # Abrir en una nueva ventana de PowerShell visible
    Start-Process powershell -ArgumentList "-NoExit", "-Command", "cd '$PWD'; Write-Host 'Ejecutando monitoreo de correos...' -ForegroundColor Green; Write-Host ''; & '$pythonExe' $arguments; Write-Host ''; Write-Host 'Presiona cualquier tecla para cerrar...' -ForegroundColor Yellow; $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')"
    
    Write-Host "Ventana abierta. El proceso se ejecutara en la nueva ventana." -ForegroundColor Green
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Ejecutando directamente el script conocido..." -ForegroundColor Yellow
    
    # Ejecutar directamente el script conocido
    $scriptPath = Join-Path $PWD "monitorear_correos_y_generar_reportes.py"
    
    if (Test-Path $scriptPath) {
        Start-Process powershell -ArgumentList "-NoExit", "-Command", "cd '$PWD'; Write-Host 'Ejecutando monitoreo de correos...' -ForegroundColor Green; Write-Host ''; python '$scriptPath' --continuo --intervalo 5; Write-Host ''; Write-Host 'Presiona cualquier tecla para cerrar...' -ForegroundColor Yellow; $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')"
        Write-Host "Ventana abierta con el script: $scriptPath" -ForegroundColor Green
    } else {
        Write-Host "No se encontro el script: $scriptPath" -ForegroundColor Red
    }
}






# Script para ver la consola de la tarea en tiempo real
# Muestra los logs actualizándose automáticamente

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "CONSOLA EN TIEMPO REAL - MONITOREO WES" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "[INFO] Mostrando logs en tiempo real..." -ForegroundColor Yellow
Write-Host "[INFO] Presiona Ctrl+C para salir" -ForegroundColor Yellow
Write-Host ""
Write-Host "=" * 70 -ForegroundColor Gray
Write-Host ""

$logFile = "logs\monitoreo_correos.log"

if (-not (Test-Path $logFile)) {
    Write-Host "[ADVERTENCIA] No se encontró el archivo de log: $logFile" -ForegroundColor Yellow
    Write-Host "[INFO] Esperando a que se cree el archivo..." -ForegroundColor Yellow
    
    # Esperar hasta que exista el archivo
    $timeout = 30
    $elapsed = 0
    while (-not (Test-Path $logFile) -and $elapsed -lt $timeout) {
        Start-Sleep -Seconds 1
        $elapsed++
        Write-Host "." -NoNewline
    }
    Write-Host ""
    
    if (-not (Test-Path $logFile)) {
        Write-Host "[ERROR] El archivo de log no se creó después de $timeout segundos" -ForegroundColor Red
        Write-Host "[INFO] Verifica que la tarea esté corriendo" -ForegroundColor Yellow
        pause
        exit
    }
}

# Mostrar últimas 20 líneas y luego seguir en tiempo real
Get-Content $logFile -Tail 20
Write-Host ""
Write-Host "=" * 70 -ForegroundColor Gray
Write-Host "[INFO] Esperando nuevas entradas... (Ctrl+C para salir)" -ForegroundColor Cyan
Write-Host ""

# Seguir mostrando nuevas líneas en tiempo real
Get-Content $logFile -Wait -Tail 20




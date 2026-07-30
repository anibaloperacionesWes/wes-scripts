# Script para ejecutar el monitoreo manualmente y ver la salida en consola
# Útil para debugging y ver qué está haciendo el sistema

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "EJECUTANDO MONITOREO MANUALMENTE" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "[INFO] Ejecutando monitoreo de correos..." -ForegroundColor Yellow
Write-Host "[INFO] Presiona Ctrl+C para detener" -ForegroundColor Yellow
Write-Host ""
Write-Host "=" * 70 -ForegroundColor Gray
Write-Host ""

$pythonExe = "C:\Users\joseo\AppData\Local\Programs\Python\Python314\python.exe"
$scriptPath = "C:\Users\joseo\Desktop\wes-scripts\monitorear_correos_y_generar_reportes.py"
$workingDir = "C:\Users\joseo\Desktop\wes-scripts"

if (-not (Test-Path $pythonExe)) {
    Write-Host "[ERROR] No se encontró Python en: $pythonExe" -ForegroundColor Red
    pause
    exit 1
}

if (-not (Test-Path $scriptPath)) {
    Write-Host "[ERROR] No se encontró el script en: $scriptPath" -ForegroundColor Red
    pause
    exit 1
}

# Ejecutar el script y mostrar la salida en tiempo real
Set-Location $workingDir
& $pythonExe $scriptPath --continuo --intervalo 5




# Script para ejecutar directamente el script de monitoreo de correos
# Esto es útil para probar o ejecutar manualmente

$scriptPath = Join-Path $PSScriptRoot "monitorear_correos_y_generar_reportes.py"
$pythonPath = "C:\Users\joseo\AppData\Local\Programs\Python\Python314\python.exe"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  EJECUCION DIRECTA DEL SCRIPT" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Verificar que el script existe
if (-not (Test-Path $scriptPath)) {
    Write-Host "[ERROR] No se encontro el script: $scriptPath" -ForegroundColor Red
    exit 1
}

Write-Host "[OK] Script encontrado: $scriptPath" -ForegroundColor Green
Write-Host "[OK] Python: $pythonPath" -ForegroundColor Green
Write-Host ""

# Cambiar al directorio del script
Set-Location $PSScriptRoot

Write-Host "[INFO] Ejecutando script..." -ForegroundColor Yellow
Write-Host "[INFO] El script se ejecutara en segundo plano" -ForegroundColor Yellow
Write-Host "[INFO] Usa '.\ver_log.ps1 -Seguir' para ver el log en tiempo real" -ForegroundColor Cyan
Write-Host ""

# Ejecutar el script
try {
    Start-Process -FilePath $pythonPath -ArgumentList "`"$scriptPath`"" -WorkingDirectory $PSScriptRoot -WindowStyle Hidden
    Write-Host "[OK] Script iniciado" -ForegroundColor Green
    Write-Host ""
    
    Start-Sleep -Seconds 3
    
    # Verificar procesos
    $pythonProcesses = Get-Process python* -ErrorAction SilentlyContinue | Where-Object { $_.Path -eq $pythonPath }
    if ($pythonProcesses) {
        Write-Host "[OK] Proceso de Python ejecutandose:" -ForegroundColor Green
        $pythonProcesses | Format-Table Id, ProcessName, StartTime -AutoSize
    }
} catch {
    Write-Host "[ERROR] No se pudo ejecutar el script: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  PROCESO COMPLETADO" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""


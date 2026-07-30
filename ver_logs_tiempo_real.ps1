$logPath = "logs\monitoreo_correos.log"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  LOGS EN TIEMPO REAL" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "Ultimas 20 lineas:" -ForegroundColor Yellow
Write-Host "----------------------------------------" -ForegroundColor Gray
Get-Content $logPath -Tail 20

Write-Host ""
Write-Host "----------------------------------------" -ForegroundColor Gray
Write-Host "Esperando nuevas entradas... (Ctrl+C para salir)" -ForegroundColor Yellow
Write-Host ""

Get-Content $logPath -Wait -Tail 1

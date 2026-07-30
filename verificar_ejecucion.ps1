Start-Sleep -Seconds 5
Write-Host "Verificando estado y logs..." -ForegroundColor Yellow
Write-Host ""

$tarea = Get-ScheduledTask | Where-Object { $_.TaskName -like '*WESMonitoreoCorreos*' }
Write-Host "Estado de la tarea: " -NoNewline
Write-Host $tarea.State -ForegroundColor Cyan
Write-Host ""

Write-Host "Ultimas 15 lineas del log:" -ForegroundColor Yellow
Write-Host "----------------------------------------" -ForegroundColor Gray
Get-Content logs\monitoreo_correos.log -Tail 15 | ForEach-Object {
    if ($_ -match '\[ERROR\]') {
        Write-Host $_ -ForegroundColor Red
    } elseif ($_ -match '\[OK\]|\[SUCCESS\]') {
        Write-Host $_ -ForegroundColor Green
    } elseif ($_ -match '\[INFO\]') {
        Write-Host $_ -ForegroundColor Cyan
    } else {
        Write-Host $_
    }
}

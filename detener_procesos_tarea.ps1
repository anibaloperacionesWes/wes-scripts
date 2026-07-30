Write-Host "Deteniendo procesos relacionados con la tarea..." -ForegroundColor Yellow
Write-Host ""

# Detener y desactivar la tarea programada (evita que vuelva a iniciarse sola)
$nombresTarea = @('WESMonitoreoCorreos')
foreach ($nombre in $nombresTarea) {
    $t = Get-ScheduledTask -TaskName $nombre -ErrorAction SilentlyContinue
    if ($t) {
        Write-Host "1. Deteniendo tarea programada: $nombre..." -ForegroundColor Yellow
        Stop-ScheduledTask -TaskName $nombre -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
        try {
            Disable-ScheduledTask -TaskName $nombre -ErrorAction Stop | Out-Null
            Write-Host "   [OK] Tarea deshabilitada (no se reiniciará sola)" -ForegroundColor Green
        } catch {
            Write-Host "   [INFO] No se pudo deshabilitar la tarea (¿sin admin?): $($_.Exception.Message)" -ForegroundColor Cyan
        }
        $estado = (Get-ScheduledTask -TaskName $nombre -ErrorAction SilentlyContinue).State
        Write-Host "   Estado: $estado" -ForegroundColor Cyan
    }
}
# Otras variantes de nombre (si existieran)
foreach ($t2 in (Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object { $_.TaskName -like '*WESMonitoreoCorreos*' })) {
    if ($nombresTarea -contains $t2.TaskName) { continue }
    Write-Host "1b. Deteniendo tarea: $($t2.TaskName)..." -ForegroundColor Yellow
    Stop-ScheduledTask -TaskName $t2.TaskName -ErrorAction SilentlyContinue
    try {
        Disable-ScheduledTask -TaskName $t2.TaskName -ErrorAction Stop | Out-Null
    } catch { }
}

# Detener procesos Python relacionados
Write-Host ""
Write-Host "2. Buscando procesos Python relacionados..." -ForegroundColor Yellow
$procesos = Get-Process python -ErrorAction SilentlyContinue | Where-Object {
    $cmd = (Get-CimInstance Win32_Process -Filter "ProcessId = $($_.Id)" -ErrorAction SilentlyContinue).CommandLine
    $cmd -like '*monitorear_correos*' -or $cmd -like '*wes-scripts*'
}

if ($procesos) {
    Write-Host "   Encontrados $($procesos.Count) proceso(s)" -ForegroundColor Yellow
    foreach ($proc in $procesos) {
        Write-Host "   Deteniendo proceso PID: $($proc.Id)..." -ForegroundColor Yellow
        Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
    }
    Write-Host "   [OK] Procesos detenidos" -ForegroundColor Green
} else {
    Write-Host "   No se encontraron procesos Python relacionados" -ForegroundColor Cyan
}

Write-Host ""
Write-Host "[OK] Tarea y procesos detenidos" -ForegroundColor Green

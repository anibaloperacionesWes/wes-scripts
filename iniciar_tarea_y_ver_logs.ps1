# Verificar e iniciar la tarea
Write-Host "Verificando estado de la tarea..." -ForegroundColor Yellow
$tarea = Get-ScheduledTask | Where-Object { $_.TaskName -like '*WESMonitoreoCorreos*' }
if ($tarea) {
    if ($tarea.State -ne 'Running') {
        Write-Host "La tarea NO esta corriendo. Iniciando..." -ForegroundColor Yellow
        Start-ScheduledTask -TaskName $tarea.TaskName
        Start-Sleep -Seconds 2
        $estado = (Get-ScheduledTask -TaskName $tarea.TaskName).State
        Write-Host "Estado actual: $estado" -ForegroundColor Green
    } else {
        Write-Host "La tarea YA esta corriendo" -ForegroundColor Green
    }
} else {
    Write-Host "No se encontro la tarea" -ForegroundColor Red
}

Write-Host ""
Write-Host "Ultimas 20 lineas del log:" -ForegroundColor Yellow
Write-Host "----------------------------------------" -ForegroundColor Gray
$logPath = "logs\monitoreo_correos.log"
if (Test-Path $logPath) {
    Get-Content $logPath -Tail 20 | ForEach-Object {
        if ($_ -match '\[ERROR\]') {
            Write-Host $_ -ForegroundColor Red
        } elseif ($_ -match '\[OK\]|\[SUCCESS\]') {
            Write-Host $_ -ForegroundColor Green
        } elseif ($_ -match '\[INFO\]') {
            Write-Host $_ -ForegroundColor Cyan
        } elseif ($_ -match '\[ADVERTENCIA\]|\[WARNING\]') {
            Write-Host $_ -ForegroundColor Yellow
        } else {
            Write-Host $_
        }
    }
} else {
    Write-Host "No se encontro el archivo de log" -ForegroundColor Red
}

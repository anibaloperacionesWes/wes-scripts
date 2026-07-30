Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  VERIFICACION DE TAREA Y CONEXION" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Verificar estado de la tarea
Write-Host "1. ESTADO DE LA TAREA:" -ForegroundColor Yellow
Write-Host "----------------------------------------" -ForegroundColor Gray
$tarea = Get-ScheduledTask | Where-Object { $_.TaskName -like '*WESMonitoreoCorreos*' }
if ($tarea) {
    $info = Get-ScheduledTaskInfo -TaskName $tarea.TaskName
    Write-Host "Nombre: " -NoNewline
    Write-Host $tarea.TaskName -ForegroundColor Cyan
    Write-Host "Estado: " -NoNewline
    if ($tarea.State -eq 'Running') {
        Write-Host $tarea.State -ForegroundColor Green
    } else {
        Write-Host $tarea.State -ForegroundColor Yellow
    }
    Write-Host "Ultima ejecucion: " -NoNewline
    Write-Host $info.LastRunTime -ForegroundColor Cyan
    Write-Host "Ultimo resultado: " -NoNewline
    if ($info.LastTaskResult -eq 0) {
        Write-Host "Exitoso (0x0)" -ForegroundColor Green
    } else {
        Write-Host "Error (0x$($info.LastTaskResult.ToString('X')))" -ForegroundColor Red
    }
} else {
    Write-Host "No se encontro la tarea WESMonitoreoCorreos" -ForegroundColor Red
}
Write-Host ""

# Verificar procesos Python
Write-Host "2. PROCESOS PYTHON RELACIONADOS:" -ForegroundColor Yellow
Write-Host "----------------------------------------" -ForegroundColor Gray
$procs = Get-Process python -ErrorAction SilentlyContinue | Where-Object { 
    $_.Path -like '*wes-scripts*' -or 
    $_.CommandLine -like '*monitorear*' -or
    $_.Path -like '*Python*'
}
if ($procs) {
    $procs | ForEach-Object {
        Write-Host "  PID: $($_.Id)" -NoNewline
        Write-Host " | Ruta: $($_.Path)" -ForegroundColor Green
    }
} else {
    Write-Host "  No hay procesos Python relacionados ejecutandose" -ForegroundColor Yellow
}
Write-Host ""

# Verificar conexion a internet
Write-Host "3. VERIFICACION DE CONEXION A INTERNET:" -ForegroundColor Yellow
Write-Host "----------------------------------------" -ForegroundColor Gray
$testSites = @(
    @{Name="Google DNS"; Address="8.8.8.8"},
    @{Name="Google"; Address="google.com"},
    @{Name="OAuth2 Google"; Address="oauth2.googleapis.com"}
)

foreach ($site in $testSites) {
    Write-Host "Probando $($site.Name) ($($site.Address))... " -NoNewline
    try {
        $result = Test-Connection -ComputerName $site.Address -Count 1 -Quiet -ErrorAction Stop
        if ($result) {
            Write-Host "OK" -ForegroundColor Green
        } else {
            Write-Host "FALLO" -ForegroundColor Red
        }
    } catch {
        Write-Host "FALLO" -ForegroundColor Red
        Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}
Write-Host ""

# Ver ultimas lineas del log
Write-Host "4. ULTIMAS LINEAS DEL LOG:" -ForegroundColor Yellow
Write-Host "----------------------------------------" -ForegroundColor Gray
$logPath = "logs\monitoreo_correos.log"
if (Test-Path $logPath) {
    $ultimas = Get-Content $logPath -Tail 15
    $ultimas | ForEach-Object {
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
    Write-Host "  No se encontro el archivo de log" -ForegroundColor Red
}
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan

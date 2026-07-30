# Script para ver los logs de la tarea de monitoreo
param(
    [int]$UltimasLineas = 50,
    [switch]$TiempoReal,
    [string]$Buscar = "",
    [switch]$Ayuda
)

$logPath = "logs\monitoreo_correos.log"

if ($Ayuda) {
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  AYUDA - VER LOGS DE LA TAREA" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Uso:" -ForegroundColor Yellow
    Write-Host "  .\ver_logs_tarea.ps1                    # Ver ultimas 50 lineas"
    Write-Host "  .\ver_logs_tarea.ps1 -UltimasLineas 100  # Ver ultimas 100 lineas"
    Write-Host "  .\ver_logs_tarea.ps1 -TiempoReal         # Ver en tiempo real (como tail -f)"
    Write-Host "  .\ver_logs_tarea.ps1 -Buscar 'ERROR'     # Buscar texto en el log"
    Write-Host "  .\ver_logs_tarea.ps1 -Ayuda              # Mostrar esta ayuda"
    Write-Host ""
    exit 0
}

if (-not (Test-Path $logPath)) {
    Write-Host "[ERROR] No se encontro el archivo de log: $logPath" -ForegroundColor Red
    exit 1
}

# Informacion del archivo
$fileInfo = Get-Item $logPath
$fileSize = [math]::Round($fileInfo.Length / 1MB, 2)
$lineCount = (Get-Content $logPath | Measure-Object -Line).Lines

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  LOGS DE MONITOREO DE CORREOS" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Archivo: " -NoNewline
Write-Host $logPath -ForegroundColor Cyan
Write-Host "Tamaño: " -NoNewline
Write-Host "$fileSize MB" -ForegroundColor Cyan
Write-Host "Total de lineas: " -NoNewline
Write-Host $lineCount -ForegroundColor Cyan
Write-Host ""

if ($TiempoReal) {
    Write-Host "Modo tiempo real (presiona Ctrl+C para salir)..." -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    
    Get-Content $logPath -Wait -Tail 10
} elseif ($Buscar -ne "") {
    Write-Host "Buscando: " -NoNewline
    Write-Host "'$Buscar'" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    
    $resultados = Get-Content $logPath | Select-String -Pattern $Buscar -CaseSensitive:$false
    $count = ($resultados | Measure-Object).Count
    
    Write-Host "Resultados encontrados: $count" -ForegroundColor Green
    Write-Host ""
    
    if ($count -gt 0) {
        $resultados | ForEach-Object {
            Write-Host $_.Line
        }
    }
} else {
    Write-Host "Ultimas $UltimasLineas lineas:" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    
    Get-Content $logPath -Tail $UltimasLineas | ForEach-Object {
        # Colorear por tipo de mensaje
        if ($_ -match '\[ERROR\]') {
            Write-Host $_ -ForegroundColor Red
        } elseif ($_ -match '\[ADVERTENCIA\]|\[WARNING\]') {
            Write-Host $_ -ForegroundColor Yellow
        } elseif ($_ -match '\[OK\]|\[SUCCESS\]') {
            Write-Host $_ -ForegroundColor Green
        } elseif ($_ -match '\[INFO\]') {
            Write-Host $_ -ForegroundColor Cyan
        } else {
            Write-Host $_
        }
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan

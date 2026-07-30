# Script para ver logs del monitoreo de correos
# Uso: .\ver_logs.ps1 [opcion]
# Opciones: tiempo-real, errores, estadisticas, ultimas [n], buscar [texto]

param(
    [string]$Modo = "tiempo-real",
    [int]$Lineas = 50,
    [string]$Buscar = ""
)

$LogFile = "logs\monitoreo_correos.log"

if (-not (Test-Path $LogFile)) {
    Write-Host "[ERROR] El archivo de log no existe: $LogFile" -ForegroundColor Red
    exit 1
}

switch ($Modo.ToLower()) {
    "tiempo-real" {
        Write-Host "=== MONITOREO DE LOG EN TIEMPO REAL ===" -ForegroundColor Cyan
        Write-Host "Presiona Ctrl+C para detener`n" -ForegroundColor Yellow
        Get-Content $LogFile -Wait -Tail $Lineas -Encoding UTF8
    }
    
    "errores" {
        Write-Host "=== BUSCANDO ERRORES Y ADVERTENCIAS ===" -ForegroundColor Yellow
        Write-Host ""
        Get-Content $LogFile -Encoding UTF8 | Select-String -Pattern 'ERROR|ADVERTENCIA|Exception|Traceback' | ForEach-Object {
            if ($_ -match 'ERROR') {
                Write-Host $_ -ForegroundColor Red
            } elseif ($_ -match 'ADVERTENCIA') {
                Write-Host $_ -ForegroundColor Yellow
            } else {
                Write-Host $_ -ForegroundColor Magenta
            }
        }
    }
    
    "ultimas" {
        Write-Host "=== ULTIMAS $Lineas LINEAS DEL LOG ===" -ForegroundColor Cyan
        Write-Host ""
        Get-Content $LogFile -Tail $Lineas -Encoding UTF8
    }
    
    "buscar" {
        if ($Buscar -eq "") {
            Write-Host "[ERROR] Debes especificar el texto a buscar" -ForegroundColor Red
            Write-Host "Uso: .\ver_logs.ps1 buscar 'texto a buscar'" -ForegroundColor Yellow
            exit 1
        }
        Write-Host "=== BUSCANDO: '$Buscar' ===" -ForegroundColor Cyan
        Write-Host ""
        Get-Content $LogFile -Encoding UTF8 | Select-String -Pattern $Buscar | ForEach-Object {
            Write-Host $_
        }
    }
    
    "estadisticas" {
        Write-Host "=== ESTADISTICAS DEL LOG ===" -ForegroundColor Cyan
        Write-Host ""
        $log = Get-Content $LogFile -Encoding UTF8
        $tamano = (Get-Item $LogFile).Length / 1MB
        $errores = ($log | Select-String -Pattern '\[ERROR\]').Count
        $advertencias = ($log | Select-String -Pattern '\[ADVERTENCIA\]').Count
        $exitos = ($log | Select-String -Pattern '\[OK\]').Count
        $info = ($log | Select-String -Pattern '\[INFO\]').Count
        
        Write-Host "Tamaño del archivo: $([math]::Round($tamano, 2)) MB" -ForegroundColor White
        Write-Host "Total de lineas: $($log.Count)" -ForegroundColor White
        Write-Host ""
        Write-Host "Errores: $errores" -ForegroundColor Red
        Write-Host "Advertencias: $advertencias" -ForegroundColor Yellow
        Write-Host "Exitos: $exitos" -ForegroundColor Green
        Write-Host "Informacion: $info" -ForegroundColor Cyan
    }
    
    "colores" {
        Write-Host "=== LOG CON COLORES (TIEMPO REAL) ===" -ForegroundColor Cyan
        Write-Host "Presiona Ctrl+C para detener`n" -ForegroundColor Yellow
        Get-Content $LogFile -Wait -Tail $Lineas -Encoding UTF8 | ForEach-Object {
            if ($_ -match '\[ERROR\]') {
                Write-Host $_ -ForegroundColor Red
            } elseif ($_ -match '\[ADVERTENCIA\]') {
                Write-Host $_ -ForegroundColor Yellow
            } elseif ($_ -match '\[OK\]|CORRECTO') {
                Write-Host $_ -ForegroundColor Green
            } elseif ($_ -match '\[INFO\]') {
                Write-Host $_ -ForegroundColor Cyan
            } elseif ($_ -match 'DEBUG') {
                Write-Host $_ -ForegroundColor Gray
            } else {
                Write-Host $_ -ForegroundColor White
            }
        }
    }
    
    default {
        Write-Host "=== GUIA DE USO ===" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Uso: .\ver_logs.ps1 [modo] [opciones]`n" -ForegroundColor Yellow
        Write-Host "Modos disponibles:" -ForegroundColor White
        Write-Host "  tiempo-real    - Ver log en tiempo real (por defecto)" -ForegroundColor Gray
        Write-Host "  errores        - Mostrar solo errores y advertencias" -ForegroundColor Gray
        Write-Host "  ultimas [n]    - Ver ultimas n lineas (ej: ultimas 100)" -ForegroundColor Gray
        Write-Host "  buscar [texto] - Buscar texto en el log" -ForegroundColor Gray
        Write-Host "  estadisticas   - Mostrar estadisticas del log" -ForegroundColor Gray
        Write-Host "  colores        - Ver log en tiempo real con colores" -ForegroundColor Gray
        Write-Host ""
        Write-Host "Ejemplos:" -ForegroundColor White
        Write-Host "  .\ver_logs.ps1                    # Tiempo real (50 lineas)" -ForegroundColor Green
        Write-Host "  .\ver_logs.ps1 tiempo-real 100    # Tiempo real (100 lineas)" -ForegroundColor Green
        Write-Host "  .\ver_logs.ps1 errores            # Solo errores" -ForegroundColor Green
        Write-Host "  .\ver_logs.ps1 ultimas 200        # Ultimas 200 lineas" -ForegroundColor Green
        Write-Host "  .\ver_logs.ps1 buscar 'PPT'       # Buscar 'PPT'" -ForegroundColor Green
        Write-Host "  .\ver_logs.ps1 estadisticas       # Estadisticas" -ForegroundColor Green
        Write-Host "  .\ver_logs.ps1 colores            # Tiempo real con colores" -ForegroundColor Green
    }
}

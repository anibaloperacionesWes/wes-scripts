# Script para ver el log de monitoreo de correos
# Uso: .\ver_log.ps1 [opciones]

param(
    [Parameter(Mandatory=$false)]
    [int]$Lineas = 50,
    
    [Parameter(Mandatory=$false)]
    [switch]$Seguir,
    
    [Parameter(Mandatory=$false)]
    [switch]$Todo,
    
    [Parameter(Mandatory=$false)]
    [string]$Buscar
)

$logPath = "logs\monitoreo_correos.log"

if (-not (Test-Path $logPath)) {
    Write-Host "[ERROR] No se encontró el archivo de log: $logPath" -ForegroundColor Red
    exit 1
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  VISOR DE LOG - MONITOREO CORREOS" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

if ($Todo) {
    Write-Host "Mostrando todo el log..." -ForegroundColor Yellow
    Write-Host ""
    Get-Content $logPath
} elseif ($Seguir) {
    Write-Host "Siguiendo el log en tiempo real (Ctrl+C para salir)..." -ForegroundColor Yellow
    Write-Host ""
    Get-Content $logPath -Wait -Tail 20
} elseif ($Buscar) {
    Write-Host "Buscando: '$Buscar'..." -ForegroundColor Yellow
    Write-Host ""
    Get-Content $logPath | Select-String -Pattern $Buscar -Context 2,2
} else {
    Write-Host "Últimas $Lineas líneas del log:" -ForegroundColor Yellow
    Write-Host ""
    Get-Content $logPath -Tail $Lineas
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  FIN DEL LOG" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan


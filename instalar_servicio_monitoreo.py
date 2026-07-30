"""
Script para instalar el servicio de Windows que monitorea correos automáticamente.
Requiere ejecutarse como administrador.
"""

import sys
import os
from pathlib import Path

def instalar_servicio():
    """Instala el servicio de Windows para monitoreo de correos."""
    
    script_path = Path(__file__).parent / "monitorear_correos_y_generar_reportes.py"
    python_exe = sys.executable
    
    # Verificar que el script existe
    if not script_path.exists():
        print(f"[ERROR] No se encontró el script: {script_path}")
        return False
    
    # Comando para ejecutar el servicio
    comando = f'"{python_exe}" "{script_path}" --continuo --intervalo 5'
    
    print("=" * 70)
    print("INSTALACIÓN DE SERVICIO DE WINDOWS")
    print("=" * 70)
    print()
    print("Este script creará un servicio de Windows que:")
    print("  - Correrá en segundo plano automáticamente")
    print("  - Se iniciará al encender el PC")
    print("  - Monitoreará correos cada 5 minutos")
    print("  - Generará y enviará reportes automáticamente")
    print()
    print("IMPORTANTE: Debes ejecutar este script como Administrador")
    print()
    
    # Usar NSSM (Non-Sucking Service Manager) o pywin32
    try:
        import win32serviceutil
        import win32service
        import servicemanager
        
        print("[INFO] Usando pywin32 para crear el servicio...")
        print("[INFO] Esto requiere configuración manual adicional")
        print()
        print("ALTERNATIVA RECOMENDADA: Usar NSSM (más simple)")
        print()
        print("Pasos para instalar con NSSM:")
        print("1. Descarga NSSM desde: https://nssm.cc/download")
        print("2. Extrae nssm.exe a una carpeta (ej: C:\\nssm)")
        print("3. Ejecuta como Administrador:")
        print(f'   C:\\nssm\\nssm.exe install WESMonitoreoCorreos "{python_exe}" "{script_path}" --continuo --intervalo 5')
        print("4. Inicia el servicio:")
        print("   C:\\nssm\\nssm.exe start WESMonitoreoCorreos")
        print()
        print("O usa el script PowerShell que se generará...")
        
    except ImportError:
        print("[INFO] pywin32 no está instalado")
        print("[INFO] Generando script PowerShell alternativo...")
    
    # Generar script PowerShell para instalar con NSSM
    crear_script_powershell(comando)
    
    return True


def crear_script_powershell(comando):
    """Crea un script PowerShell para instalar el servicio usando NSSM."""
    
    script_ps = Path(__file__).parent / "instalar_servicio_monitoreo.ps1"
    
    contenido = f"""# Script PowerShell para instalar servicio de monitoreo de correos WES
# Debe ejecutarse como Administrador

$ErrorActionPreference = "Stop"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "INSTALACIÓN DE SERVICIO WES MONITOREO" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Verificar que se ejecuta como Administrador
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {{
    Write-Host "[ERROR] Este script debe ejecutarse como Administrador" -ForegroundColor Red
    Write-Host "Haz clic derecho y selecciona 'Ejecutar como administrador'" -ForegroundColor Yellow
    pause
    exit 1
}}

# Ruta de NSSM
$nssmPath = "C:\\nssm\\nssm.exe"
$nssmUrl = "https://nssm.cc/release/nssm-2.24.zip"

# Verificar si NSSM existe
if (-not (Test-Path $nssmPath)) {{
    Write-Host "[INFO] NSSM no encontrado. Descargando..." -ForegroundColor Yellow
    Write-Host "[INFO] Por favor, descarga NSSM manualmente desde: https://nssm.cc/download" -ForegroundColor Yellow
    Write-Host "[INFO] Extrae nssm.exe a C:\\nssm\\" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "O ejecuta estos comandos:" -ForegroundColor Cyan
    Write-Host "  mkdir C:\\nssm" -ForegroundColor White
    Write-Host "  Invoke-WebRequest -Uri `"$nssmUrl`" -OutFile `"C:\\nssm\\nssm.zip`"" -ForegroundColor White
    Write-Host "  Expand-Archive -Path `"C:\\nssm\\nssm.zip`" -DestinationPath `"C:\\nssm\\`"" -ForegroundColor White
    Write-Host "  Move-Item `"C:\\nssm\\nssm-2.24\\win64\\nssm.exe`" `"C:\\nssm\\nssm.exe`"" -ForegroundColor White
    pause
    exit 1
}}

# Parámetros del servicio
$serviceName = "WESMonitoreoCorreos"
$serviceDisplayName = "WES Monitoreo de Correos y Reportes"
$serviceDescription = "Monitorea correos automáticamente y genera reportes de consumo de agua"

# Ruta del script Python
$scriptPath = "{Path(__file__).parent / "monitorear_correos_y_generar_reportes.py"}".Replace("\\", "\\\\")
$pythonExe = "{sys.executable}".Replace("\\", "\\\\")

Write-Host "[INFO] Instalando servicio..." -ForegroundColor Cyan
Write-Host "  Nombre: $serviceName" -ForegroundColor White
Write-Host "  Python: $pythonExe" -ForegroundColor White
Write-Host "  Script: $scriptPath" -ForegroundColor White
Write-Host ""

# Eliminar servicio si ya existe
$existingService = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
if ($existingService) {{
    Write-Host "[INFO] El servicio ya existe. Eliminando versión anterior..." -ForegroundColor Yellow
    & $nssmPath stop $serviceName
    Start-Sleep -Seconds 2
    & $nssmPath remove $serviceName confirm
    Start-Sleep -Seconds 2
}}

# Instalar servicio
Write-Host "[INFO] Creando servicio..." -ForegroundColor Cyan
& $nssmPath install $serviceName $pythonExe "`"$scriptPath`" --continuo --intervalo 5"

if ($LASTEXITCODE -ne 0) {{
    Write-Host "[ERROR] Falló la instalación del servicio" -ForegroundColor Red
    pause
    exit 1
}}

# Configurar servicio
Write-Host "[INFO] Configurando servicio..." -ForegroundColor Cyan
& $nssmPath set $serviceName DisplayName "$serviceDisplayName"
& $nssmPath set $serviceName Description "$serviceDescription"
& $nssmPath set $serviceName Start SERVICE_AUTO_START
& $nssmPath set $serviceName AppStdout "C:\\nssm\\wes_monitoreo_stdout.log"
& $nssmPath set $serviceName AppStderr "C:\\nssm\\wes_monitoreo_stderr.log"

# Iniciar servicio
Write-Host "[INFO] Iniciando servicio..." -ForegroundColor Cyan
& $nssmPath start $serviceName

if ($LASTEXITCODE -eq 0) {{
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "SERVICIO INSTALADO EXITOSAMENTE" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "El servicio '$serviceDisplayName' está corriendo." -ForegroundColor Green
    Write-Host ""
    Write-Host "Comandos útiles:" -ForegroundColor Cyan
    Write-Host "  Iniciar:   net start $serviceName" -ForegroundColor White
    Write-Host "  Detener:   net stop $serviceName" -ForegroundColor White
    Write-Host "  Estado:    sc query $serviceName" -ForegroundColor White
    Write-Host "  Logs:      C:\\nssm\\wes_monitoreo_stdout.log" -ForegroundColor White
    Write-Host ""
}} else {{
    Write-Host "[ERROR] No se pudo iniciar el servicio" -ForegroundColor Red
    Write-Host "Revisa los logs en: C:\\nssm\\wes_monitoreo_stderr.log" -ForegroundColor Yellow
}}

pause
"""
    
    with open(script_ps, 'w', encoding='utf-8') as f:
        f.write(contenido)
    
    print(f"[OK] Script PowerShell creado: {script_ps}")
    print()
    print("Para instalar el servicio:")
    print("1. Haz clic derecho en 'instalar_servicio_monitoreo.ps1'")
    print("2. Selecciona 'Ejecutar con PowerShell' (como Administrador)")
    print()


if __name__ == "__main__":
    instalar_servicio()








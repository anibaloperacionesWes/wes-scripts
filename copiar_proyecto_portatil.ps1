# Script para copiar el proyecto WES a una carpeta portátil
# Uso: .\copiar_proyecto_portatil.ps1

param(
    [string]$Destino = "D:\WES_PROYECTO"
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "COPIANDO PROYECTO WES A CARPETA PORTÁTIL" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Obtener la ruta del script actual
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Origen = $ScriptDir

Write-Host "Origen: $Origen" -ForegroundColor Yellow
Write-Host "Destino: $Destino" -ForegroundColor Yellow
Write-Host ""

# Crear estructura de carpetas
Write-Host "Creando estructura de carpetas..." -ForegroundColor Green
New-Item -ItemType Directory -Force -Path "$Destino\scripts" | Out-Null
New-Item -ItemType Directory -Force -Path "$Destino\documentacion" | Out-Null
New-Item -ItemType Directory -Force -Path "$Destino\config" | Out-Null
New-Item -ItemType Directory -Force -Path "$Destino\scripts_auxiliares" | Out-Null

# Archivos principales (scripts)
Write-Host "Copiando scripts principales..." -ForegroundColor Green
$ScriptsPrincipales = @(
    "generar_reporte_word.py",
    "actualizar_empresas_nodos.py"
)

foreach ($script in $ScriptsPrincipales) {
    if (Test-Path "$Origen\$script") {
        Copy-Item "$Origen\$script" "$Destino\scripts\" -Force
        Write-Host "  ✓ $script" -ForegroundColor Gray
    }
}

# Documentación
Write-Host "Copiando documentación..." -ForegroundColor Green
$Documentacion = @(
    "README.md",
    "GUIA_CREACION_USUARIOS.md",
    "PROMPT_CREAR_USUARIO.md",
    "GUIA_ENVIO_CORREO.md",
    "INSTRUCCIONES_ENVIO_CORREO.md",
    "obtener_contraseña_aplicacion.md",
    "PROMPT_REPORTE_ALERTAS.md"
)

foreach ($doc in $Documentacion) {
    if (Test-Path "$Origen\$doc") {
        Copy-Item "$Origen\$doc" "$Destino\documentacion\" -Force
        Write-Host "  ✓ $doc" -ForegroundColor Gray
    }
}

# Configuración
Write-Host "Copiando archivos de configuración..." -ForegroundColor Green
if (Test-Path "$Origen\api-docs.json") {
    Copy-Item "$Origen\api-docs.json" "$Destino\config\" -Force
    Write-Host "  ✓ api-docs.json" -ForegroundColor Gray
}

# Scripts auxiliares
Write-Host "Copiando scripts auxiliares..." -ForegroundColor Green
$ScriptsAuxiliares = Get-ChildItem -Path $Origen -Filter "generar_y_enviar_reportes_*.py"
foreach ($script in $ScriptsAuxiliares) {
    Copy-Item $script.FullName "$Destino\scripts_auxiliares\" -Force
    Write-Host "  ✓ $($script.Name)" -ForegroundColor Gray
}

# Otros scripts auxiliares
$OtrosScripts = @(
    "generar_reportes_san_ignacio.py",
    "probar_envio_correo.py",
    "reenviar_reporte.py"
)

foreach ($script in $OtrosScripts) {
    if (Test-Path "$Origen\$script") {
        Copy-Item "$Origen\$script" "$Destino\scripts_auxiliares\" -Force
        Write-Host "  ✓ $script" -ForegroundColor Gray
    }
}

# Crear requirements.txt
Write-Host "Creando requirements.txt..." -ForegroundColor Green
$Requirements = @(
    "requests>=2.31.0",
    "python-docx>=1.1.0",
    "matplotlib>=3.7.0",
    "numpy>=1.24.0",
    "Pillow>=10.0.0"
)

$Requirements | Set-Content "$Destino\requirements.txt" -Encoding UTF8
Write-Host "  ✓ requirements.txt creado" -ForegroundColor Gray

# Crear SETUP_INSTRUCCIONES.md
Write-Host "Creando SETUP_INSTRUCCIONES.md..." -ForegroundColor Green
$SetupInstrucciones = @(
    "# Instrucciones de Instalación - Proyecto WES",
    "",
    "## Requisitos Previos",
    "",
    "1. **Python 3.10 o superior**",
    "   - Descargar desde: https://www.python.org/downloads/",
    "   - Durante la instalación, marcar 'Add Python to PATH'",
    "",
    "2. **Verificar instalación de Python**",
    "   Ejecutar en PowerShell:",
    "   python --version",
    "",
    "## Instalación de Dependencias",
    "",
    "1. Abrir PowerShell o CMD en esta carpeta",
    "2. Ejecutar:",
    "   pip install -r requirements.txt",
    "",
    "## Estructura del Proyecto",
    "",
    "* scripts/ - Scripts principales del proyecto",
    "* documentacion/ - Documentación y guías",
    "* config/ - Archivos de configuración",
    "* scripts_auxiliares/ - Scripts auxiliares y temporales",
    "",
    "## Uso Básico",
    "",
    "### Generar un reporte individual:",
    "python scripts\generar_reporte_word.py --company-id 000025 --node-id 000025-20 --start-date 01112025 --end-date 07122025",
    "",
    "### Generar reporte agregado:",
    "Ver scripts en scripts_auxiliares/ para ejemplos de reportes agregados.",
    "",
    "## Notas",
    "",
    "* Los reportes se generan en la carpeta reports/ (se crea automáticamente)",
    "* Asegúrate de tener conexión a internet para acceder a la API",
    "* Para envío de correos, configura las credenciales SMTP según la documentación",
    "",
    "## Solución de Problemas",
    "",
    "### Python no encontrado:",
    "* Verificar que Python esté en el PATH del sistema",
    "* Usar la ruta completa: C:\Users\[Usuario]\AppData\Local\Programs\Python\Python314\python.exe",
    "",
    "### Error al instalar dependencias:",
    "* Actualizar pip: python -m pip install --upgrade pip",
    "* Instalar dependencias una por una si es necesario"
)

$SetupInstrucciones | Set-Content "$Destino\SETUP_INSTRUCCIONES.md" -Encoding UTF8
Write-Host "  ✓ SETUP_INSTRUCCIONES.md creado" -ForegroundColor Gray

# Crear README principal
Write-Host "Creando README.md principal..." -ForegroundColor Green
$Readme = @(
    "# Proyecto WES - Generación de Reportes",
    "",
    "Sistema para generar reportes de consumo y fugas de agua desde la API WES.",
    "",
    "## Estructura",
    "",
    "* scripts/ - Scripts principales",
    "* documentacion/ - Documentación y guías",
    "* config/ - Configuración",
    "* scripts_auxiliares/ - Scripts auxiliares",
    "",
    "## Inicio Rápido",
    "",
    "1. Instalar dependencias: pip install -r requirements.txt",
    "2. Ver SETUP_INSTRUCCIONES.md para detalles",
    "3. Ver documentacion/README.md para documentación completa",
    "",
    "## Script Principal",
    "",
    "scripts\generar_reporte_word.py - Genera reportes individuales y agregados",
    "",
    "## Documentación",
    "",
    "Ver carpeta documentacion/ para:",
    "* Guías de uso",
    "* Prompts utilizados",
    "* Instrucciones de configuración"
)

$Readme | Set-Content "$Destino\README.md" -Encoding UTF8
Write-Host "  ✓ README.md creado" -ForegroundColor Gray

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "PROYECTO COPIADO EXITOSAMENTE" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Ubicación: $Destino" -ForegroundColor Yellow
Write-Host ""
Write-Host "Próximos pasos:" -ForegroundColor Cyan
Write-Host "1. Revisar la carpeta creada" -ForegroundColor White
Write-Host "2. Copiar esta carpeta a otro computador" -ForegroundColor White
Write-Host "3. Seguir las instrucciones en SETUP_INSTRUCCIONES.md" -ForegroundColor White
Write-Host ""


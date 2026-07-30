# WES Scripts - Gestión de Usuarios

Scripts para interactuar con la API WES (Water Efficiency System) para crear y gestionar usuarios.

## 📁 Archivos

### Scripts Principales
- **`crear_usuario.ps1`** - Script principal para crear usuarios (busca empresa por nombre automáticamente)
- **`get_company.ps1`** - Obtiene información de una compañía por ID
- **`generar_reporte_word.py`** - Genera reportes de consumo y fugas en Word a partir de la API ACL Node
- **`gestionar_incidentes.py`** - Analiza y gestiona incidentes relacionados con consumo, fugas y alertas
- **`reporte_incidentes_dia_anterior.py`** - Genera reporte Word de incidentes del día anterior (consumo cero, sin respuesta, datos incompletos)

### Documentación
- **`PROMPT_CREAR_USUARIO.md`** - Prompt estándar para usar con el asistente de IA
- **`PROMPT_INCIDENTES.md`** - Prompt completo para gestión de incidentes con acceso a todas las APIs
- **`PROMPT_REPORTE_ALERTAS.md`** - Prompt para generar reportes de consumo y alertas
- **`GUIA_CREACION_USUARIOS.md`** - Guía completa paso a paso para crear usuarios
- **`api-docs.json`** - Documentación completa de la API en formato JSON

## 🚀 Inicio Rápido

### Crear un Usuario

El script principal `crear_usuario.ps1` acepta el nombre de la empresa (y opcionalmente el ID):

```powershell
.\crear_usuario.ps1 -NombreEmpresa "Lo valledor" -Nombre "Jose Luis" -Apellido "Otarola" -Email "usuario@ejemplo.com"
```

**Con ID de empresa (opcional):**
```powershell
.\crear_usuario.ps1 -NombreEmpresa "Lo valledor" -CompanyId "000002" -Nombre "Jose Luis" -Apellido "Otarola" -Email "usuario@ejemplo.com"
```

### Obtener Información de una Compañía

```powershell
powershell -ExecutionPolicy Bypass -File get_company.ps1
```

### Generar Reporte de Consumo y Fugas

Requiere Python 3.10+ y las dependencias `requests`, `matplotlib` y `python-docx`:

```powershell
pip install requests matplotlib python-docx
python generar_reporte_word.py ^
  --company-id 000025 ^
  --node-id 000025-12 ^
  --start-date 2025-11-01 ^
  --end-date 2025-11-26
```

El archivo Word se almacena en la carpeta `reports/` con las gráficas y tablas generadas.

### Analizar Incidentes

Requiere Python 3.10+ y las dependencias `requests`:

```powershell
python gestionar_incidentes.py ^
  --node-id 000025-20 ^
  --start-date 01112025 ^
  --end-date 07122025 ^
  --output resultado.json
```

Este script analiza un incidente completo, obteniendo:
- Información del nodo y empresa
- Medidas de consumo del periodo
- Alertas de fuga
- Datos horarios para días con alertas
- Precio del agua (si está disponible)

### Generar Reporte de Incidentes del Día Anterior

Requiere Python 3.10+ y las dependencias `requests` y `python-docx`:

```powershell
python reporte_incidentes_dia_anterior.py
```

Este script:
- Analiza **todos los puntos** del sistema
- Obtiene datos del día anterior a la ejecución
- Clasifica puntos en:
  - Consumo cero
  - Sin respuesta (error de API)
  - Datos incompletos (menos de 24 horas, indicando horas faltantes)
- Genera un documento Word con tablas detalladas
- Guarda el reporte en `reports/incidentes/REPORTE_DIA_ANTERIOR_[FECHA]/`

## 📖 Documentación

- **Prompt para el Asistente**: [PROMPT_CREAR_USUARIO.md](PROMPT_CREAR_USUARIO.md)
- **Guía Completa**: [GUIA_CREACION_USUARIOS.md](GUIA_CREACION_USUARIOS.md)

## 🔗 Enlaces Útiles

- **Swagger UI**: http://104.248.53.141:7001/wes/api/acl-entities/v1/swagger-ui.html
- **API Docs**: http://104.248.53.141:7001/wes/api/acl-entities/v1/v2/api-docs

## ⚡ Proceso de Creación de Usuarios

El script `crear_usuario.ps1` automáticamente:

1. **Busca la empresa** por nombre (si no se proporciona ID) desde 000001 hasta 000100
2. **Obtiene todos los nodos** de la compañía
3. **Crea el usuario** con acceso a todos esos nodos

## 📋 Datos Requeridos

- **Nombre de la empresa** (obligatorio)
- **ID de la empresa** (opcional - se busca automáticamente si no se proporciona)
- **Nombre** de la persona
- **Apellido** de la persona
- **Correo electrónico**

**Nota**: Los nodos se obtienen automáticamente, no necesitas proporcionarlos.

## 💡 Ejemplo de Uso con el Asistente

```
Crea un usuario en WES API con los siguientes datos:

- Nombre de la empresa: Lo valledor
- Nombre: Jose Luis
- Apellido: Otarola
- Correo: usuario@ejemplo.com

El usuario debe tener acceso a TODOS los nodos de la empresa especificada.
```

## ❓ ¿Necesitas Ayuda?

Consulta la [Guía Completa](GUIA_CREACION_USUARIOS.md) que incluye:
- Explicación detallada del proceso
- Ejemplos prácticos
- Solución de problemas comunes
- Checklist para crear usuarios

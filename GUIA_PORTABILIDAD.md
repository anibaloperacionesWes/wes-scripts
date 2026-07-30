# Guía de Portabilidad - Proyecto WES Scripts

Esta guía explica cómo guardar todos los scripts, prompts y documentación en una carpeta portátil para trasladar a otro computador.

## 📁 Estructura Recomendada

```
D:\WES_PROYECTO\                    (o la ruta que prefieras)
├── scripts\                       # Scripts principales
│   ├── generar_reporte_word.py    # Script principal
│   ├── actualizar_empresas_nodos.py
│   └── [otros scripts .py]
├── documentacion\                  # Documentación y prompts
│   ├── README.md
│   ├── GUIA_CREACION_USUARIOS.md
│   ├── PROMPT_CREAR_USUARIO.md
│   └── [otros .md]
├── config\                         # Archivos de configuración
│   └── api-docs.json
├── scripts_auxiliares\             # Scripts temporales o auxiliares
│   ├── generar_y_enviar_reportes_*.py
│   └── [otros scripts auxiliares]
├── requirements.txt                # Dependencias Python
└── SETUP_INSTRUCCIONES.md          # Instrucciones de instalación
```

## 📋 Pasos para Crear la Carpeta Portátil

### 1. Crear la Carpeta Principal

Abre PowerShell o el Explorador de Archivos y crea una carpeta en una ubicación fija, por ejemplo:

```
D:\WES_PROYECTO
```

O en otra unidad:
```
E:\PROYECTOS\WES_SCRIPTS
```

### 2. Copiar Archivos Esenciales

**Scripts principales (obligatorios):**
- `generar_reporte_word.py` - Script principal
- `actualizar_empresas_nodos.py` - Para actualizar nombres

**Documentación:**
- `README.md`
- `GUIA_CREACION_USUARIOS.md`
- `PROMPT_CREAR_USUARIO.md`
- `GUIA_ENVIO_CORREO.md` (si existe)
- Cualquier otro archivo `.md` con documentación

**Configuración:**
- `api-docs.json` (si lo usas)

**Scripts auxiliares (opcionales, según necesidad):**
- `generar_y_enviar_reportes_*.py` - Scripts específicos por empresa
- Otros scripts que hayas creado

### 3. Crear requirements.txt

Crea un archivo `requirements.txt` con todas las dependencias:

```
requests>=2.31.0
python-docx>=1.1.0
matplotlib>=3.7.0
numpy>=1.24.0
Pillow>=10.0.0
```

### 4. Crear Instrucciones de Instalación

Crea `SETUP_INSTRUCCIONES.md` con los pasos para configurar en otro computador.

## 🚀 Script Automático para Copiar Archivos

He creado un script `copiar_proyecto_portatil.ps1` que automatiza este proceso.

## 📝 Notas Importantes

### Archivos que NO debes copiar:
- `__pycache__/` - Archivos compilados de Python (se regeneran)
- `reports/` - Reportes generados (son datos, no código)
- Archivos temporales de prueba

### Consideraciones para Portabilidad:

1. **Python**: Asegúrate de tener Python 3.10+ instalado en el otro computador
2. **Dependencias**: Instala con `pip install -r requirements.txt`
3. **Rutas absolutas**: El código usa rutas relativas, así que funciona en cualquier ubicación
4. **Variables de entorno**: Si usas credenciales, considera usar variables de entorno o archivos de configuración

## 🔧 Configuración en el Nuevo Computador

1. Copiar la carpeta completa a la nueva ubicación
2. Abrir Cursor en esa carpeta
3. Instalar dependencias: `pip install -r requirements.txt`
4. Verificar que Python esté en el PATH
5. Probar ejecutando un script simple

## 📍 Ubicación Recomendada

**Opción 1 (Disco D:):**
```
D:\WES_PROYECTO
```

**Opción 2 (Disco C:, carpeta Proyectos):**
```
C:\Proyectos\WES_SCRIPTS
```

**Opción 3 (USB o Disco Externo):**
```
E:\WES_PROYECTO
```

La ventaja de usar una ruta fija es que puedes crear accesos directos y scripts que siempre apunten a la misma ubicación.



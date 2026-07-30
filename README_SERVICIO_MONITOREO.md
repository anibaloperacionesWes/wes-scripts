# Servicio de Windows - Monitoreo Automático de Correos

Este servicio monitorea tu bandeja de correo automáticamente y genera reportes cuando detecta solicitudes.

## 🚀 Instalación

### Paso 1: Descargar NSSM

1. Ve a: https://nssm.cc/download
2. Descarga la versión más reciente (nssm-2.24.zip o superior)
3. Extrae el archivo ZIP
4. Copia `nssm.exe` de la carpeta `win64` a `C:\nssm\`
   - Si la carpeta no existe, créala

### Paso 2: Instalar el Servicio

**Opción A: Usar el script PowerShell (Recomendado)**

1. Haz clic derecho en `instalar_servicio_monitoreo.ps1`
2. Selecciona "Ejecutar con PowerShell" (como Administrador)
3. El script instalará y configurará el servicio automáticamente

**Opción B: Instalación Manual**

Abre PowerShell como Administrador y ejecuta:

```powershell
# Instalar servicio
C:\nssm\nssm.exe install WESMonitoreoCorreos "C:\Users\joseo\AppData\Local\Programs\Python\Python314\python.exe" "C:\Users\joseo\Desktop\wes-scripts\monitorear_correos_y_generar_reportes.py --continuo --intervalo 5"

# Configurar servicio
C:\nssm\nssm.exe set WESMonitoreoCorreos DisplayName "WES Monitoreo de Correos y Reportes"
C:\nssm\nssm.exe set WESMonitoreoCorreos Description "Monitorea correos automáticamente y genera reportes de consumo de agua"
C:\nssm\nssm.exe set WESMonitoreoCorreos Start SERVICE_AUTO_START

# Iniciar servicio
C:\nssm\nssm.exe start WESMonitoreoCorreos
```

## ⚙️ Configuración

### Intervalo de Revisión

Por defecto, el servicio revisa correos cada **5 minutos**. Para cambiar:

1. Detén el servicio: `net stop WESMonitoreoCorreos`
2. Edita el comando en NSSM o reinstala con el intervalo deseado
3. Inicia el servicio: `net start WESMonitoreoCorreos`

### Correos Autorizados

Edita `lista_contactos_reportes.py` y agrega correos a `CORREOS_AUTORIZADOS` con su configuración.

## 📋 Comandos Útiles

```powershell
# Ver estado del servicio
sc query WESMonitoreoCorreos

# Iniciar servicio
net start WESMonitoreoCorreos

# Detener servicio
net stop WESMonitoreoCorreos

# Reiniciar servicio
net stop WESMonitoreoCorreos
net start WESMonitoreoCorreos

# Ver logs
Get-Content C:\nssm\wes_monitoreo_stdout.log -Tail 50
Get-Content C:\nssm\wes_monitoreo_stderr.log -Tail 50
```

## 🔍 Verificar que Funciona

1. Abre el Visor de Eventos de Windows
2. Ve a: Servicios de Windows → WESMonitoreoCorreos
3. O revisa los logs en: `C:\nssm\wes_monitoreo_stdout.log`

## ❌ Desinstalar

```powershell
# Detener y eliminar servicio
C:\nssm\nssm.exe stop WESMonitoreoCorreos
C:\nssm\nssm.exe remove WESMonitoreoCorreos confirm
```

## ⚠️ Requisitos

- PC encendido (el servicio corre en segundo plano)
- Conexión a Internet
- Credenciales de Gmail configuradas
- Python y dependencias instaladas

## 📝 Notas

- El servicio se inicia automáticamente al encender el PC
- Corre en segundo plano, no necesitas tener Cursor o terminal abierta
- Revisa correos cada 5 minutos (configurable)
- Solo procesa correos autorizados en `lista_contactos_reportes.py`








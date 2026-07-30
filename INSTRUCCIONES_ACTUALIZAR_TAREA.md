# 🔧 Instrucciones para Actualizar la Tarea Programada de Windows

## ⚠️ Problema Detectado

La tarea programada `WESMonitoreoCorreos` está configurada para ejecutar Python **sin argumentos**, por lo que no ejecuta el script de monitoreo de correos.

## ✅ Solución: Actualizar la Tarea Manualmente

### Opción 1: Usar el Programador de Tareas (Recomendado)

1. **Abrir el Programador de Tareas:**
   - Presiona `Win + R`
   - Escribe: `taskschd.msc`
   - Presiona Enter

2. **Buscar la tarea:**
   - En el panel izquierdo, busca `WESMonitoreoCorreos`
   - Haz clic derecho → **Propiedades**

3. **Actualizar la acción:**
   - Ve a la pestaña **Acciones**
   - Selecciona la acción existente y haz clic en **Editar**
   - **Programa o script:** `C:\Users\joseo\AppData\Local\Programs\Python\Python314\python.exe`
   - **Agregar argumentos:** `"C:\Users\joseo\Desktop\wes-scripts\monitorear_correos_y_generar_reportes.py"`
   - **Iniciar en:** `C:\Users\joseo\Desktop\wes-scripts`
   - Haz clic en **Aceptar**

4. **Guardar:**
   - Haz clic en **Aceptar** en la ventana de propiedades

### Opción 2: Usar PowerShell como Administrador

1. **Abrir PowerShell como Administrador:**
   - Presiona `Win + X`
   - Selecciona "Windows PowerShell (Administrador)" o "Terminal (Administrador)"

2. **Ejecutar el siguiente comando:**

```powershell
$taskName = "WESMonitoreoCorreos"
$pythonPath = "C:\Users\joseo\AppData\Local\Programs\Python\Python314\python.exe"
$scriptPath = "C:\Users\joseo\Desktop\wes-scripts\monitorear_correos_y_generar_reportes.py"
$workingDir = "C:\Users\joseo\Desktop\wes-scripts"

# Obtener la tarea actual
$task = Get-ScheduledTask -TaskName $taskName

# Crear nueva acción
$action = New-ScheduledTaskAction -Execute $pythonPath -Argument "`"$scriptPath`"" -WorkingDirectory $workingDir

# Obtener configuración actual
$triggers = $task.Triggers
$settings = $task.Settings
$principal = $task.Principal

# Actualizar la tarea
Set-ScheduledTask -TaskName $taskName -Action $action -Trigger $triggers -Settings $settings -Principal $principal

Write-Host "Tarea actualizada exitosamente" -ForegroundColor Green
```

### Opción 3: Ejecutar el Script Directamente (Temporal)

Mientras se actualiza la tarea, puedes ejecutar el script directamente:

```powershell
.\ejecutar_script_directo.ps1
```

O manualmente:

```powershell
cd C:\Users\joseo\Desktop\wes-scripts
python monitorear_correos_y_generar_reportes.py
```

---

## 🔍 Verificar que Funciona

Después de actualizar la tarea:

1. **Verificar la configuración:**
```powershell
$task = Get-ScheduledTask -TaskName "WESMonitoreoCorreos"
$task.Actions[0] | Format-List *
```

Deberías ver:
- **Execute:** `C:\Users\joseo\AppData\Local\Programs\Python\Python314\python.exe`
- **Argument:** `"C:\Users\joseo\Desktop\wes-scripts\monitorear_correos_y_generar_reportes.py"`
- **WorkingDirectory:** `C:\Users\joseo\Desktop\wes-scripts`

2. **Ejecutar la tarea manualmente:**
```powershell
Start-ScheduledTask -TaskName "WESMonitoreoCorreos"
```

3. **Verificar el log:**
```powershell
.\ver_log.ps1 -Seguir
```

---

## 📝 Notas

- La tarea debe ejecutarse con permisos de administrador para actualizarse
- El script se ejecuta en un bucle continuo, revisando correos cada 5 minutos
- El log se guarda en `logs\monitoreo_correos.log`


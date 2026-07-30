# Guía Paso a Paso: Configurar Tarea Programada Manualmente

## 📋 Resumen
Esta guía te ayudará a configurar manualmente la tarea programada para el monitoreo automático de correos en el Programador de Tareas de Windows.

---

## 🎯 PASO 1: Crear la Tarea

1. En el **Programador de Tareas**, en el panel derecho (Acciones)
2. Haz clic en **"Crear tarea..."** (NO "Crear tarea básica")
3. Se abrirá una ventana con varias pestañas

---

## 📝 PASO 2: Pestaña "General"

Configura los siguientes campos:

- **Nombre:** `WESMonitoreoCorreos`
- **Descripción:** `Monitorea correos automáticamente y genera reportes de consumo de agua`
- **Seguridad:**
  - ✅ Marca: **"Ejecutar tanto si el usuario ha iniciado sesión como si no"**
  - ✅ Marca: **"Ejecutar con los privilegios más altos"**
- **Configurar para:** `Windows 10` (o tu versión de Windows)

---

## ⏰ PASO 3: Pestaña "Desencadenadores"

1. Haz clic en el botón **"Nuevo..."**
2. En **"Iniciar la tarea:"** selecciona: **"Al iniciar el equipo"**
3. ✅ Asegúrate de que **"Habilitado"** esté marcado
4. Haz clic en **"Aceptar"**

---

## 🚀 PASO 4: Pestaña "Acciones"

1. Haz clic en el botón **"Nuevo..."**
2. **Acción:** Debe estar en "Iniciar un programa" (por defecto)
3. **Programa o script:** 
   ```
   C:\Users\joseo\AppData\Local\Programs\Python\Python314\python.exe
   ```
4. **Agregar argumentos (opcional):**
   ```
   "C:\Users\joseo\Desktop\wes-scripts\monitorear_correos_y_generar_reportes.py" --continuo --intervalo 5
   ```
5. **Iniciar en (opcional):**
   ```
   C:\Users\joseo\Desktop\wes-scripts
   ```
6. Haz clic en **"Aceptar"**

---

## ⚙️ PASO 5: Pestaña "Condiciones"

- ✅ Marca: **"Iniciar la tarea solo si el equipo está conectado a la corriente"**
- ❌ Desmarca: **"Activar la tarea solo si el equipo está conectado a la corriente alterna"**
- Deja las demás opciones como están

---

## 🔧 PASO 6: Pestaña "Configuración"

- ✅ Marca: **"Permitir ejecutar la tarea a petición"**
- ✅ Marca: **"Ejecutar la tarea tan pronto como sea posible después de una programación perdida"**
- ✅ Marca: **"Si la tarea falla, reiniciar cada:"** → Selecciona **"1 minuto"**
- **Intento de reinicio hasta:** `3 veces`

---

## ✅ PASO 7: Finalizar

1. Haz clic en **"Aceptar"** en la ventana principal
2. Si te pide tu contraseña de Windows, ingrésala
3. La tarea se creará y debería iniciarse automáticamente

---

## 🔍 Verificar que Funciona

Después de crear la tarea:

1. En el Programador de Tareas, ve a **"Biblioteca del Programador de tareas"** (panel izquierdo)
2. Busca la tarea **"WESMonitoreoCorreos"**
3. Verifica que el **Estado** sea **"Lista"** o **"En ejecución"**
4. Si está en "Lista", haz clic derecho → **"Ejecutar"**

---

## 📊 Comandos Útiles (PowerShell como Administrador)

```powershell
# Ver estado
Get-ScheduledTask -TaskName "WESMonitoreoCorreos"

# Iniciar manualmente
Start-ScheduledTask -TaskName "WESMonitoreoCorreos"

# Detener
Stop-ScheduledTask -TaskName "WESMonitoreoCorreos"

# Ver información detallada
Get-ScheduledTaskInfo -TaskName "WESMonitoreoCorreos"
```

---

## ⚠️ Notas Importantes

- La tarea se iniciará automáticamente al encender el PC
- Corre en segundo plano, no necesitas tener Cursor abierto
- Solo requiere que el PC esté encendido
- Revisa correos cada 5 minutos automáticamente

---

¿Necesitas ayuda con algún paso específico? Avísame y te guío.








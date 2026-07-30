# 📋 Guía para Ver los Logs

## 🚀 Formas de Ver el Log

### 1. **Ver las últimas líneas (Recomendado)**
```powershell
.\ver_log.ps1
```
Muestra las últimas 50 líneas por defecto.

### 2. **Ver más líneas**
```powershell
.\ver_log.ps1 -Lineas 100
```
Muestra las últimas 100 líneas.

### 3. **Seguir el log en tiempo real**
```powershell
.\ver_log.ps1 -Seguir
```
Muestra las últimas 20 líneas y se actualiza automáticamente (como `tail -f` en Linux).
Presiona `Ctrl+C` para salir.

### 4. **Ver todo el log**
```powershell
.\ver_log.ps1 -Todo
```
⚠️ **Advertencia**: El log puede ser muy grande (más de 300,000 líneas).

### 5. **Buscar texto específico**
```powershell
.\ver_log.ps1 -Buscar "ERROR"
.\ver_log.ps1 -Buscar "Ciclo #85"
.\ver_log.ps1 -Buscar "diegocarrasco"
```

### 6. **Combinar opciones**
```powershell
# Ver últimas 200 líneas buscando "reporte"
.\ver_log.ps1 -Lineas 200 | Select-String "reporte"
```

---

## 📝 Comandos PowerShell Directos

### Ver últimas líneas
```powershell
Get-Content logs\monitoreo_correos.log -Tail 50
```

### Ver últimas 100 líneas
```powershell
Get-Content logs\monitoreo_correos.log -Tail 100
```

### Seguir el log en tiempo real
```powershell
Get-Content logs\monitoreo_correos.log -Wait -Tail 20
```

### Buscar texto
```powershell
Get-Content logs\monitoreo_correos.log | Select-String "ERROR"
Get-Content logs\monitoreo_correos.log | Select-String "reporte" -Context 5,5
```

### Ver solo ciclos recientes
```powershell
Get-Content logs\monitoreo_correos.log -Tail 200 | Select-String "Ciclo"
```

### Ver solo errores
```powershell
Get-Content logs\monitoreo_correos.log | Select-String "ERROR" -Context 3,3
```

### Ver correos procesados
```powershell
Get-Content logs\monitoreo_correos.log | Select-String "Procesando correo" -Context 2,5
```

### Ver resumen de cada ciclo
```powershell
Get-Content logs\monitoreo_correos.log | Select-String "PROCESO COMPLETADO|Ciclo|correos procesados" -Context 1,1
```

---

## 🔍 Ejemplos Útiles

### Ver actividad de la última hora
```powershell
$hora = (Get-Date).AddHours(-1).ToString("yyyy-MM-dd HH")
Get-Content logs\monitoreo_correos.log | Select-String $hora
```

### Ver solo los correos que fueron procesados como solicitudes
```powershell
Get-Content logs\monitoreo_correos.log | Select-String "solicitud de reporte|generando reporte" -Context 3,10
```

### Ver estadísticas de un ciclo específico
```powershell
Get-Content logs\monitoreo_correos.log | Select-String "Ciclo #85" -Context 0,30
```

### Contar cuántos correos se procesaron hoy
```powershell
$hoy = (Get-Date).ToString("yyyy-MM-dd")
(Get-Content logs\monitoreo_correos.log | Select-String $hoy | Select-String "Procesando correo").Count
```

---

## 📊 Información del Archivo de Log

- **Ubicación**: `logs\monitoreo_correos.log`
- **Tamaño**: Puede ser muy grande (varios MB)
- **Formato**: Texto plano con timestamps
- **Rotación**: No se rota automáticamente (puede crecer indefinidamente)

---

## 💡 Tips

1. **Usa `-Tail` para ver lo más reciente**: El log crece hacia abajo, así que las últimas líneas son las más relevantes.

2. **Usa `Select-String` para filtrar**: Es más eficiente que leer todo el archivo.

3. **Usa `-Context` para ver contexto**: Ver líneas antes y después de la coincidencia.

4. **Para logs grandes**: Siempre usa `-Tail` o `Select-String` en lugar de leer todo el archivo.


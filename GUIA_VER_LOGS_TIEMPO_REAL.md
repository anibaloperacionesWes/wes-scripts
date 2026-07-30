# Guía: Ver Logs en Tiempo Real

## 📋 Métodos para Ver Logs en Tiempo Real

### Método 1: Script Simple (Recomendado)
```powershell
.\ver_logs_tiempo_real.ps1
```
**Ventajas:**
- Muestra las últimas 20 líneas primero
- Luego espera nuevas entradas
- Formato claro y fácil de leer

---

### Método 2: Script con Opciones
```powershell
.\ver_logs_tarea.ps1 -TiempoReal
```
**Ventajas:**
- Parte del script completo con más opciones
- Puedes usar otros parámetros también

---

### Método 3: Comando Directo (Más Simple)
```powershell
Get-Content logs\monitoreo_correos.log -Wait -Tail 10
```
**Ventajas:**
- No requiere scripts
- Comando directo de PowerShell
- Muestra las últimas 10 líneas y espera nuevas

**Parámetros:**
- `-Wait`: Espera nuevas líneas (modo tiempo real)
- `-Tail 10`: Muestra las últimas 10 líneas antes de esperar
- Puedes cambiar el número (ej: `-Tail 20`, `-Tail 50`)

---

### Método 4: Ver Últimas Líneas + Tiempo Real
```powershell
Get-Content logs\monitoreo_correos.log -Tail 20; Get-Content logs\monitoreo_correos.log -Wait -Tail 1
```
**Ventajas:**
- Primero muestra las últimas 20 líneas
- Luego espera y muestra solo las nuevas líneas (1 a la vez)

---

## 🎯 Ejemplos de Uso

### Ver últimas 50 líneas y luego tiempo real:
```powershell
Get-Content logs\monitoreo_correos.log -Tail 50; Get-Content logs\monitoreo_correos.log -Wait -Tail 1
```

### Ver solo tiempo real (sin mostrar líneas previas):
```powershell
Get-Content logs\monitoreo_correos.log -Wait
```

### Ver tiempo real con colores (usando el script):
```powershell
.\ver_logs_tarea.ps1 -TiempoReal
```

---

## ⚠️ Notas Importantes

1. **Para Detener:** Presiona `Ctrl+C` en cualquier momento
2. **Ubicación:** Asegúrate de estar en la carpeta `wes-scripts` o ajusta la ruta
3. **Permisos:** Si hay problemas, usa:
   ```powershell
   powershell -ExecutionPolicy Bypass -File ver_logs_tiempo_real.ps1
   ```

---

## 🔍 Comandos Relacionados

### Ver últimas líneas (sin tiempo real):
```powershell
Get-Content logs\monitoreo_correos.log -Tail 50
```

### Buscar texto en los logs:
```powershell
Get-Content logs\monitoreo_correos.log | Select-String "ERROR"
```

### Contar líneas totales:
```powershell
(Get-Content logs\monitoreo_correos.log | Measure-Object -Line).Lines
```

### Ver tamaño del archivo:
```powershell
(Get-Item logs\monitoreo_correos.log).Length / 1MB
```

"""
Script para ver el proceso completo de la tarea: estado + logs en tiempo real.
"""

import sys
import subprocess
import json
import time
from pathlib import Path
from datetime import datetime

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

LOG_FILE = Path(__file__).parent / "logs" / "monitoreo_correos.log"


def obtener_estado_tarea():
    """Obtiene el estado de la tarea programada."""
    try:
        ps_command = """
        $tareas = Get-ScheduledTask | Where-Object {
            $taskName = $_.TaskName.ToLower()
            $taskName -like '*correo*' -or 
            $taskName -like '*monitoreo*' -or 
            $taskName -like '*wes*'
        }
        
        if ($tareas.Count -eq 0) {
            $allTasks = Get-ScheduledTask
            foreach ($task in $allTasks) {
                $action = $task.Actions.Execute
                if ($action -like '*python*' -and ($task.Actions.Arguments -like '*monitorear*')) {
                    $tareas = @($task)
                    break
                }
            }
        }
        
        $resultado = @()
        foreach ($tarea in $tareas) {
            $info = Get-ScheduledTaskInfo -TaskName $tarea.TaskName
            $resultado += @{
                TaskName = $tarea.TaskName
                State = $tarea.State.ToString()
                LastRunTime = if ($info.LastRunTime) { $info.LastRunTime.ToString('yyyy-MM-dd HH:mm:ss') } else { $null }
                NextRunTime = if ($info.NextRunTime) { $info.NextRunTime.ToString('yyyy-MM-dd HH:mm:ss') } else { $null }
                LastTaskResult = $info.LastTaskResult
            }
        }
        
        $resultado | ConvertTo-Json -Depth 10
        """
        
        result = subprocess.run(
            ["powershell", "-ExecutionPolicy", "Bypass", "-Command", ps_command],
            capture_output=True,
            text=True,
            encoding='utf-8'
        )
        
        if result.returncode == 0 and result.stdout.strip():
            data = json.loads(result.stdout)
            if isinstance(data, dict):
                return [data]
            return data
        return []
    except Exception as e:
        return []


def mostrar_estado():
    """Muestra el estado actual de la tarea."""
    print("=" * 70)
    print("  ESTADO DE LA TAREA")
    print("=" * 70)
    
    tareas = obtener_estado_tarea()
    if tareas:
        tarea = tareas[0]
        estado = tarea.get('State', 'Unknown')
        last_run = tarea.get('LastRunTime', 'Nunca ejecutada')
        resultado = tarea.get('LastTaskResult', 0)
        
        print(f"Tarea: {tarea.get('TaskName', 'N/A')}")
        print(f"Estado: {estado}")
        print(f"Última ejecución: {last_run}")
        if resultado == 0:
            print(f"Último resultado: ✅ Exitoso")
        else:
            print(f"Último resultado: ❌ Error (0x{resultado:X})")
    else:
        print("No se encontró la tarea")
    
    print("=" * 70)
    print()


def ver_logs_tiempo_real(ultimas_lineas=50):
    """Muestra los logs en tiempo real."""
    if not LOG_FILE.exists():
        print(f"[ERROR] No se encontró el archivo de log: {LOG_FILE}")
        return
    
    # Mostrar últimas líneas primero
    try:
        with open(LOG_FILE, 'r', encoding='utf-8') as f:
            lineas = f.readlines()
        
        total = len(lineas)
        inicio = max(0, total - ultimas_lineas)
        
        print("=" * 70)
        print("  LOGS DEL MONITOREO (Últimas líneas)")
        print("=" * 70)
        print(f"Archivo: {LOG_FILE}")
        print(f"Total de líneas: {total}")
        print(f"Mostrando últimas {min(ultimas_lineas, total)} líneas:")
        print("=" * 70)
        print()
        
        for linea in lineas[inicio:]:
            print(linea.rstrip())
        
        print()
        print("=" * 70)
        print("  MONITOREO EN TIEMPO REAL")
        print("=" * 70)
        print("Esperando nuevas líneas... (Presiona Ctrl+C para salir)")
        print("=" * 70)
        print()
        
        # Monitorear nuevas líneas
        ultimo_count = total
        try:
            while True:
                with open(LOG_FILE, 'r', encoding='utf-8') as f:
                    lineas_actuales = f.readlines()
                
                count_actual = len(lineas_actuales)
                
                if count_actual > ultimo_count:
                    # Mostrar nuevas líneas
                    nuevas = lineas_actuales[ultimo_count:]
                    for linea in nuevas:
                        print(linea.rstrip())
                    ultimo_count = count_actual
                
                time.sleep(1)
        except KeyboardInterrupt:
            print("\n\nMonitoreo detenido.")
    
    except Exception as e:
        print(f"[ERROR] Error al leer logs: {e}")


def main():
    """Función principal."""
    import argparse
    
    parser = argparse.ArgumentParser(description="Ver proceso completo de la tarea")
    parser.add_argument(
        "--solo-estado",
        action="store_true",
        help="Mostrar solo el estado de la tarea (sin logs)"
    )
    parser.add_argument(
        "--solo-logs",
        action="store_true",
        help="Mostrar solo los logs (sin estado)"
    )
    parser.add_argument(
        "--lineas",
        type=int,
        default=50,
        help="Número de líneas iniciales a mostrar (default: 50)"
    )
    
    args = parser.parse_args()
    
    if not args.solo_logs:
        mostrar_estado()
    
    if not args.solo_estado:
        ver_logs_tiempo_real(args.lineas)


if __name__ == "__main__":
    main()

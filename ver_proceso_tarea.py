"""
Script para mostrar el estado y proceso de la tarea programada de Windows en Python.
"""

import sys
import subprocess
import json
from datetime import datetime
from pathlib import Path

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')


def ejecutar_powershell(comando: str) -> dict:
    """Ejecuta un comando de PowerShell y retorna el resultado como diccionario."""
    try:
        ps_command = f"""
        $tareas = Get-ScheduledTask | Where-Object {{
            $taskName = $_.TaskName.ToLower()
            $taskName -like '*correo*' -or 
            $taskName -like '*reporte*' -or 
            $taskName -like '*monitoreo*' -or 
            $taskName -like '*wes*' -or
            $taskName -like '*monitorearcorreos*'
        }}
        
        if ($tareas.Count -eq 0) {{
            $allTasks = Get-ScheduledTask
            foreach ($task in $allTasks) {{
                $action = $task.Actions.Execute
                if ($action -like '*python*' -and ($task.Actions.Arguments -like '*monitorear*' -or $task.Actions.Arguments -like '*correo*')) {{
                    $tareas = @($task)
                    break
                }}
            }}
        }}
        
        $resultado = @()
        foreach ($tarea in $tareas) {{
            $info = Get-ScheduledTaskInfo -TaskName $tarea.TaskName
            $resultado += @{{
                TaskName = $tarea.TaskName
                State = $tarea.State.ToString()
                LastRunTime = if ($info.LastRunTime) {{ $info.LastRunTime.ToString('yyyy-MM-dd HH:mm:ss') }} else {{ $null }}
                NextRunTime = if ($info.NextRunTime) {{ $info.NextRunTime.ToString('yyyy-MM-dd HH:mm:ss') }} else {{ $null }}
                LastTaskResult = $info.LastTaskResult
                Execute = $tarea.Actions.Execute
                Arguments = $tarea.Actions.Arguments
            }}
        }}
        
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
            # Si es un solo diccionario, convertir a lista
            if isinstance(data, dict):
                return [data]
            # Si es una lista, retornarla
            elif isinstance(data, list):
                return data
            return []
        return []
    except Exception as e:
        print(f"[ERROR] Error al ejecutar PowerShell: {e}")
        return []


def obtener_procesos_python() -> list:
    """Obtiene los procesos de Python relacionados con monitoreo."""
    try:
        ps_command = """
        Get-Process python -ErrorAction SilentlyContinue | Where-Object {
            $_.Path -like '*wes-scripts*' -or 
            $_.CommandLine -like '*monitorear*' -or
            $_.CommandLine -like '*correo*'
        } | Select-Object Id, ProcessName, Path, StartTime | ConvertTo-Json -Depth 10
        """
        
        result = subprocess.run(
            ["powershell", "-ExecutionPolicy", "Bypass", "-Command", ps_command],
            capture_output=True,
            text=True,
            encoding='utf-8'
        )
        
        if result.returncode == 0 and result.stdout.strip():
            procesos = json.loads(result.stdout)
            # Si es un solo proceso, convertir a lista
            if isinstance(procesos, dict):
                return [procesos]
            return procesos
        return []
    except Exception as e:
        print(f"[ERROR] Error al obtener procesos: {e}")
        return []


def formatear_estado(estado: str) -> str:
    """Formatea el estado de la tarea con colores."""
    estados = {
        "Running": f"🟢 {estado} (En ejecución)",
        "Ready": f"🟡 {estado} (Lista para ejecutarse)",
        "Disabled": f"🔴 {estado} (Deshabilitada)",
        "Queued": f"🟠 {estado} (En cola)",
    }
    return estados.get(estado, f"⚪ {estado}")


def formatear_resultado(resultado: int) -> str:
    """Formatea el resultado de la última ejecución."""
    if resultado == 0:
        return "✅ Exitoso (0x0)"
    elif resultado:
        return f"❌ Error (0x{resultado:X})"
    else:
        return "⚪ N/A"


def main():
    """Función principal."""
    print("=" * 70)
    print("  ESTADO Y PROCESO DE LA TAREA DE WINDOWS")
    print("=" * 70)
    print()
    
    # Obtener información de las tareas
    print("[1/2] Obteniendo información de tareas programadas...")
    tareas = ejecutar_powershell("")
    
    if not tareas:
        print("[ADVERTENCIA] No se encontraron tareas relacionadas.")
        print()
        print("Buscando todas las tareas programadas...")
        try:
            ps_all = "Get-ScheduledTask | Select-Object TaskName, State | ConvertTo-Json -Depth 10"
            result = subprocess.run(
                ["powershell", "-ExecutionPolicy", "Bypass", "-Command", ps_all],
                capture_output=True,
                text=True,
                encoding='utf-8'
            )
            if result.returncode == 0:
                todas = json.loads(result.stdout)
                if isinstance(todas, dict):
                    todas = [todas]
                print(f"Total de tareas encontradas: {len(todas)}")
                for tarea in todas[:10]:  # Mostrar solo las primeras 10
                    print(f"  - {tarea.get('TaskName', 'N/A')}: {tarea.get('State', 'N/A')}")
        except:
            pass
        return 1
    
    print(f"[OK] Se encontraron {len(tareas)} tarea(s)")
    print()
    
    # Mostrar información de cada tarea
    for i, tarea in enumerate(tareas, 1):
        print("=" * 70)
        print(f"  TAREA {i}: {tarea.get('TaskName', 'N/A')}")
        print("=" * 70)
        print()
        
        # Estado
        estado = tarea.get('State', 'Unknown')
        print(f"Estado: {formatear_estado(estado)}")
        
        # Última ejecución
        last_run = tarea.get('LastRunTime')
        if last_run:
            print(f"Última ejecución: {last_run}")
        else:
            print("Última ejecución: Nunca ejecutada")
        
        # Próxima ejecución
        next_run = tarea.get('NextRunTime')
        if next_run:
            print(f"Próxima ejecución: {next_run}")
        else:
            print("Próxima ejecución: No programada")
        
        # Resultado de última ejecución
        resultado = tarea.get('LastTaskResult')
        print(f"Último resultado: {formatear_resultado(resultado)}")
        
        # Acción
        execute = tarea.get('Execute', 'N/A')
        arguments = tarea.get('Arguments', '')
        print(f"Acción: {execute}")
        if arguments:
            print(f"Argumentos: {arguments}")
        
        print()
    
    # Obtener procesos de Python
    print("=" * 70)
    print("  PROCESOS DE PYTHON EN EJECUCIÓN")
    print("=" * 70)
    print()
    
    print("[2/2] Buscando procesos de Python relacionados...")
    procesos = obtener_procesos_python()
    
    if procesos:
        print(f"[OK] Se encontraron {len(procesos)} proceso(s)")
        print()
        for proc in procesos:
            pid = proc.get('Id', 'N/A')
            nombre = proc.get('ProcessName', 'N/A')
            ruta = proc.get('Path', 'N/A')
            inicio = proc.get('StartTime', 'N/A')
            print(f"  PID: {pid}")
            print(f"  Nombre: {nombre}")
            print(f"  Ruta: {ruta}")
            if inicio and inicio != 'N/A':
                print(f"  Inicio: {inicio}")
            print()
    else:
        print("[INFO] No se encontraron procesos de Python relacionados.")
        print()
    
    print("=" * 70)
    print("  VERIFICACIÓN COMPLETADA")
    print("=" * 70)
    print()
    print(f"Fecha de verificación: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    return 0


if __name__ == "__main__":
    sys.exit(main())

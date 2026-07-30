"""Script para limpiar reportes duplicados de Fundo Zapallar, dejando solo los más recientes."""

import sys
from pathlib import Path
from datetime import datetime

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

def limpiar_duplicados_individuales():
    """Limpia reportes individuales duplicados."""
    reportes_dir = Path("reports/Fundo_Zapallar/REPORTE")
    if not reportes_dir.exists():
        print("No existe el directorio de reportes individuales.")
        return 0
    
    # Agrupar carpetas por nombre base (sin timestamp)
    grupos = {}
    for carpeta in reportes_dir.iterdir():
        if carpeta.is_dir():
            # Extraer nombre base: todo hasta el último _ seguido de 8 dígitos (fecha)
            # Formato esperado: Nombre_20251209_HHMM
            import re
            match = re.match(r'^(.+)_\d{8}_\d{4}$', carpeta.name)
            if match:
                nombre_base = match.group(1)
            else:
                # Si no coincide el patrón, usar todo excepto los últimos dos segmentos
                partes = carpeta.name.rsplit('_', 2)
                if len(partes) >= 3:
                    nombre_base = '_'.join(partes[:-2])
                else:
                    nombre_base = carpeta.name
            
            if nombre_base not in grupos:
                grupos[nombre_base] = []
            grupos[nombre_base].append(carpeta)
    
    eliminados = 0
    for nombre_base, carpetas in grupos.items():
        if len(carpetas) > 1:
            # Ordenar por fecha de modificación (más reciente primero)
            carpetas.sort(key=lambda x: x.stat().st_mtime, reverse=True)
            # Eliminar todas excepto la más reciente
            for carpeta_vieja in carpetas[1:]:
                try:
                    import shutil
                    shutil.rmtree(carpeta_vieja)
                    print(f"Eliminado: {carpeta_vieja.name}")
                    eliminados += 1
                except Exception as e:
                    print(f"Error al eliminar {carpeta_vieja.name}: {e}")
    
    return eliminados

def limpiar_duplicados_agregados():
    """Limpia reportes agregados duplicados."""
    agregado_dir = Path("reports/Fundo_Zapallar/ABREGADO")
    if not agregado_dir.exists():
        print("No existe el directorio de reportes agregados.")
        return 0
    
    # Buscar todas las carpetas AGREGADO
    carpetas = [c for c in agregado_dir.iterdir() if c.is_dir() and c.name.startswith("AGREGADO_")]
    
    if len(carpetas) <= 1:
        return 0
    
    # Ordenar por fecha de modificación (más reciente primero)
    carpetas.sort(key=lambda x: x.stat().st_mtime, reverse=True)
    
    eliminados = 0
    # Eliminar todas excepto la más reciente
    for carpeta_vieja in carpetas[1:]:
        try:
            import shutil
            shutil.rmtree(carpeta_vieja)
            print(f"Eliminado: {carpeta_vieja.name}")
            eliminados += 1
        except Exception as e:
            print(f"Error al eliminar {carpeta_vieja.name}: {e}")
    
    return eliminados

def main():
    print("=" * 70)
    print("LIMPIANDO REPORTES DUPLICADOS DE FUNDO ZAPALLAR")
    print("=" * 70)
    print()
    
    print("Limpiando reportes individuales...")
    eliminados_individuales = limpiar_duplicados_individuales()
    print(f"  Eliminados: {eliminados_individuales}")
    print()
    
    print("Limpiando reportes agregados...")
    eliminados_agregados = limpiar_duplicados_agregados()
    print(f"  Eliminados: {eliminados_agregados}")
    print()
    
    print("=" * 70)
    print("LIMPIEZA COMPLETADA")
    print("=" * 70)
    print(f"Total eliminados: {eliminados_individuales + eliminados_agregados}")
    print()
    print("NOTA: El código ahora elimina automáticamente duplicados")
    print("      antes de generar nuevos reportes.")

if __name__ == "__main__":
    main()


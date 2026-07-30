"""Script para limpiar reportes duplicados de Fundo Zapallar, dejando solo los más recientes."""

import sys
from pathlib import Path
from datetime import datetime

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

def limpiar_duplicados():
    """Elimina reportes duplicados, dejando solo el más reciente de cada nodo."""
    reportes_dir = Path("reports/Fundo_Zapallar/REPORTE")
    agregado_dir = Path("reports/Fundo_Zapallar/ABREGADO")
    
    if not reportes_dir.exists():
        print("No existe la carpeta de reportes individuales.")
        return
    
    print("=" * 70)
    print("LIMPIANDO REPORTES DUPLICADOS DE FUNDO ZAPALLAR")
    print("=" * 70)
    print()
    
    # Limpiar reportes individuales
    print("Limpiando reportes individuales...")
    carpetas = list(reportes_dir.iterdir())
    
    # Agrupar por nombre base (sin timestamp)
    grupos = {}
    for carpeta in carpetas:
        if carpeta.is_dir():
            # Extraer nombre base (ej: "Matriz_ESVAL" de "Matriz_ESVAL_20251209_2204")
            partes = carpeta.name.rsplit('_', 2)  # Separar por los últimos 2 guiones bajos
            if len(partes) >= 3:
                nombre_base = '_'.join(partes[:-2])  # Todo excepto los últimos 2 (fecha y hora)
                if nombre_base not in grupos:
                    grupos[nombre_base] = []
                grupos[nombre_base].append(carpeta)
    
    total_eliminados = 0
    for nombre_base, carpetas_grupo in grupos.items():
        if len(carpetas_grupo) > 1:
            # Ordenar por fecha de modificación (más reciente primero)
            carpetas_grupo.sort(key=lambda x: x.stat().st_mtime, reverse=True)
            mas_reciente = carpetas_grupo[0]
            duplicados = carpetas_grupo[1:]
            
            print(f"  {nombre_base}:")
            print(f"    Manteniendo: {mas_reciente.name}")
            for dup in duplicados:
                print(f"    Eliminando: {dup.name}")
                try:
                    import shutil
                    shutil.rmtree(dup)
                    total_eliminados += 1
                except Exception as e:
                    print(f"      ERROR: No se pudo eliminar {dup.name}: {e}")
    
    print()
    print(f"Total de carpetas duplicadas eliminadas: {total_eliminados}")
    print()
    
    # Limpiar reportes agregados
    if agregado_dir.exists():
        print("Limpiando reportes agregados...")
        carpetas_agregado = list(agregado_dir.iterdir())
        if len(carpetas_agregado) > 1:
            carpetas_agregado.sort(key=lambda x: x.stat().st_mtime, reverse=True)
            mas_reciente = carpetas_agregado[0]
            duplicados = carpetas_agregado[1:]
            
            print(f"  Manteniendo: {mas_reciente.name}")
            for dup in duplicados:
                print(f"  Eliminando: {dup.name}")
                try:
                    import shutil
                    shutil.rmtree(dup)
                    total_eliminados += 1
                except Exception as e:
                    print(f"    ERROR: No se pudo eliminar {dup.name}: {e}")
            print()
    
    print("=" * 70)
    print("LIMPIEZA COMPLETADA")
    print("=" * 70)
    print(f"Total eliminado: {total_eliminados} carpetas duplicadas")

if __name__ == "__main__":
    limpiar_duplicados()


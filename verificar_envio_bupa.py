"""Script para verificar el estado de los reportes y envío de correos de BUPA."""

import sys
from pathlib import Path
from generar_reporte_word import get_company_name, get_node_name

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

NODOS_BUPA = [
    "000029-01",  # Llenado de Estanques
    "000029-02",  # Torre A
    "000029-03",  # Torre B1
    "000029-04",  # Torre B2
    "000029-05",  # Torre C
    "000029-06",  # Central Térmica
]

COMPANY_ID = "000029"
START_DATE = "03/10/2025"
END_DATE = "09/12/2025"

def main():
    print("=" * 60)
    print("VERIFICACIÓN DE REPORTES BUPA")
    print("=" * 60)
    print(f"Periodo: {START_DATE} - {END_DATE}")
    print()
    
    company_name = get_company_name(COMPANY_ID)
    company_name_clean = company_name.replace(" ", "_")
    report_dir = Path("reports") / company_name_clean / "REPORTE"
    agregado_dir = Path("reports") / company_name_clean / "ABREGADO"
    
    print("REPORTES INDIVIDUALES:")
    print("-" * 60)
    
    reportes_encontrados = []
    for node_id in NODOS_BUPA:
        node_name = get_node_name(node_id)
        node_name_clean = node_name.replace(" ", "_")
        
        if report_dir.exists():
            matching_dirs = [d for d in report_dir.iterdir() if d.is_dir() and node_name_clean in d.name]
            if matching_dirs:
                latest_dir = max(matching_dirs, key=lambda d: d.stat().st_mtime)
                report_files = list(latest_dir.glob("Reporte_*.docx"))
                if report_files:
                    reporte = report_files[0]
                    size_kb = reporte.stat().st_size / 1024
                    fecha_mod = reporte.stat().st_mtime
                    from datetime import datetime
                    fecha_str = datetime.fromtimestamp(fecha_mod).strftime("%Y-%m-%d %H:%M:%S")
                    print(f"✓ {node_name} ({node_id})")
                    print(f"  Archivo: {reporte.name}")
                    print(f"  Tamaño: {size_kb:.1f} KB")
                    print(f"  Fecha: {fecha_str}")
                    reportes_encontrados.append(reporte)
                else:
                    print(f"✗ {node_name} ({node_id}) - No se encontró archivo")
            else:
                print(f"✗ {node_name} ({node_id}) - No se encontró carpeta")
        else:
            print(f"✗ {node_name} ({node_id}) - Directorio no existe")
        print()
    
    print("REPORTE AGREGADO:")
    print("-" * 60)
    
    if agregado_dir.exists():
        matching_dirs = [d for d in agregado_dir.iterdir() if d.is_dir() and "AGREGADO" in d.name]
        if matching_dirs:
            latest_dir = max(matching_dirs, key=lambda d: d.stat().st_mtime)
            report_files = list(latest_dir.glob("Reporte_*.docx"))
            if report_files:
                reporte = report_files[0]
                size_kb = reporte.stat().st_size / 1024
                fecha_mod = reporte.stat().st_mtime
                from datetime import datetime
                fecha_str = datetime.fromtimestamp(fecha_mod).strftime("%Y-%m-%d %H:%M:%S")
                print(f"✓ Reporte Agregado")
                print(f"  Archivo: {reporte.name}")
                print(f"  Tamaño: {size_kb:.1f} KB")
                print(f"  Fecha: {fecha_str}")
                reportes_encontrados.append(reporte)
            else:
                print("✗ No se encontró archivo agregado")
        else:
            print("✗ No se encontró carpeta de agregado")
    else:
        print("✗ Directorio de agregado no existe")
    
    print()
    print("=" * 60)
    print("RESUMEN")
    print("=" * 60)
    print(f"Reportes individuales encontrados: {len([r for r in reportes_encontrados if 'Agregado' not in r.name])}/6")
    print(f"Reporte agregado encontrado: {'Sí' if any('Agregado' in r.name for r in reportes_encontrados) else 'No'}")
    print(f"Total de reportes: {len(reportes_encontrados)}")
    
    if len(reportes_encontrados) == 7:
        print()
        print("✓ TODOS LOS REPORTES ESTÁN GENERADOS")
        print()
        print("NOTA: Para verificar si los correos se enviaron correctamente,")
        print("revisa las bandejas de entrada de los destinatarios:")
        print("  - agente.ia@wes.cl")
        print("  - benjamingumucio@wes.cl")
        print("  - diegocarrasco@wes.cl")
        print("  - juanlopez@wes.cl")
    else:
        print()
        print("⚠ ALGUNOS REPORTES FALTAN")
        print("Ejecuta el script generar_y_enviar_todos_reportes_bupa.py para completar")

if __name__ == "__main__":
    main()
















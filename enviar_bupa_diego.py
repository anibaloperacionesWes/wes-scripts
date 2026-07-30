"""Script para enviar los reportes de BUPA ya generados solo a diegocarrasco@wes.cl."""

import sys
from pathlib import Path
from generar_y_enviar_todos_reportes_bupa import enviar_reportes_por_correo, get_company_name
from generar_reporte_word import get_node_name

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

NODOS_BUPA = [
    "000029-01",
    "000029-02",
    "000029-03",
    "000029-04",
    "000029-05",
    "000029-06",
]

COMPANY_ID = "000029"
START_DATE = "03/10/2025"
END_DATE = "09/12/2025"

DESTINATARIO = "diegocarrasco@wes.cl"

def main():
    print("=" * 60)
    print("ENVIANDO REPORTES DE BUPA A DIEGO CARRASCO")
    print("=" * 60)
    print(f"Destinatario: {DESTINATARIO}")
    print(f"Periodo: {START_DATE} - {END_DATE}")
    print()
    
    # Buscar todos los reportes ya generados
    company_name = get_company_name(COMPANY_ID)
    company_name_clean = company_name.replace(" ", "_")
    report_dir = Path("reports") / company_name_clean / "REPORTE"
    agregado_dir = Path("reports") / company_name_clean / "ABREGADO"
    
    reportes_individuales = []
    
    print("Buscando reportes individuales...")
    for node_id in NODOS_BUPA:
        node_name = get_node_name(node_id)
        node_name_clean = node_name.replace(" ", "_")
        
        if report_dir.exists():
            matching_dirs = [d for d in report_dir.iterdir() if d.is_dir() and node_name_clean in d.name]
            if matching_dirs:
                latest_dir = max(matching_dirs, key=lambda d: d.stat().st_mtime)
                report_files = list(latest_dir.glob("Reporte_*.docx"))
                if report_files:
                    reportes_individuales.append(report_files[0])
                    print(f"  ✓ {node_name}")
                else:
                    print(f"  ✗ {node_name} - No se encontró archivo")
            else:
                print(f"  ✗ {node_name} - No se encontró carpeta")
        else:
            print(f"  ✗ {node_name} - Directorio no existe")
    
    print()
    print("Buscando reporte agregado...")
    reporte_agregado = None
    if agregado_dir.exists():
        matching_dirs = [d for d in agregado_dir.iterdir() if d.is_dir() and "AGREGADO" in d.name]
        if matching_dirs:
            latest_dir = max(matching_dirs, key=lambda d: d.stat().st_mtime)
            report_files = list(latest_dir.glob("Reporte_*.docx"))
            if report_files:
                reporte_agregado = report_files[0]
                print(f"  ✓ Reporte Agregado")
            else:
                print(f"  ✗ No se encontró archivo agregado")
        else:
            print(f"  ✗ No se encontró carpeta de agregado")
    else:
        print(f"  ✗ Directorio de agregado no existe")
    
    if not reportes_individuales:
        print()
        print("[ERROR] No se encontraron reportes individuales para enviar")
        return
    
    todos_reportes = reportes_individuales.copy()
    if reporte_agregado:
        todos_reportes.append(reporte_agregado)
    
    print()
    print("=" * 60)
    print(f"ENVIANDO {len(todos_reportes)} REPORTES A {DESTINATARIO}")
    print("=" * 60)
    print()
    
    exito = enviar_reportes_por_correo(
        reportes_paths=todos_reportes,
        destinatario=DESTINATARIO,
        company_name=company_name,
        start_date=START_DATE,
        end_date=END_DATE,
        total_nodos=len(reportes_individuales),
    )
    
    if exito:
        print()
        print("=" * 60)
        print("[OK] CORREO ENVIADO EXITOSAMENTE")
        print("=" * 60)
        print(f"Destinatario: {DESTINATARIO}")
        print(f"Total de reportes enviados: {len(todos_reportes)}")
        print(f"  - Reportes individuales: {len(reportes_individuales)}")
        print(f"  - Reporte agregado: {'Sí' if reporte_agregado else 'No'}")
    else:
        print()
        print("=" * 60)
        print("[ERROR] FALLO EL ENVÍO DEL CORREO")
        print("=" * 60)

if __name__ == "__main__":
    main()















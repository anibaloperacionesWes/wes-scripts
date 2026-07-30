"""Script para generar reportes de todos los nodos de Parque Arauco y enviar por correo."""

import argparse
import subprocess
import sys
from pathlib import Path
from generar_reporte_word import generate_aggregated_report, enviar_reporte_por_correo, get_company_name, get_node_name

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

# Puntos de Parque Arauco Kennedy (empresa 000025)
NODOS_PARQUE_ARAUCO = [
    "000025-20",  # PAK Impulsión Ander3-4 Matriz Principal
    "000025-21",  # PAK Impulsión Ander3-4 Locales Gast.
    "000025-22",  # PAK Impulsión Sandia Baños 2-3-6-7 Fredo
    "000025-23",  # PAK Llenado Pileta
    "000025-24",  # PAK Llenado Pileta Cascada
    "000025-27",  # PAK Distrito de lujo DL
    "000025-28",  # PAK Impulsión Mall 1 Piso-4
    "000025-29",  # PAK Impulsión Anden 3-4 Restaurante
    "000025-35",  # PAK BAZAR GOURMET (reemplazo 000025-25)
    "000025-36",  # PAK DL KENNEDY (reemplazo 000025-26)
]

COMPANY_ID = "000025"
START_DATE = "07122025"  # 07 de diciembre 2025
END_DATE = "14122025"    # 14 de diciembre 2025

# Configuración de correo
SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_PASSWORD = "vxbynfpoehbweelj"  # Contraseña de aplicación
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587
DESTINATARIO = "agente.ia@wes.cl"

PYTHON_EXE = r"C:\Users\joseo\AppData\Local\Programs\Python\Python314\python.exe"
SCRIPT_PATH = "generar_reporte_word.py"

def main():
    # Parsear argumentos de línea de comandos
    parser = argparse.ArgumentParser(description="Genera reportes de todos los nodos de Parque Arauco")
    parser.add_argument(
        "--enviar-correo",
        action="store_true",
        default=False,  # Por defecto NO enviar correo
        help="Enviar los reportes por correo electrónico"
    )
    parser.add_argument(
        "--destinatario",
        default=DESTINATARIO,
        help=f"Correo electrónico del destinatario (default: {DESTINATARIO})"
    )
    args = parser.parse_args()
    
    enviar_correo = args.enviar_correo
    destinatario = args.destinatario
    
    print("=" * 60)
    print("GENERANDO REPORTES INDIVIDUALES PARA PARQUE ARAUCO KENNEDY")
    print("=" * 60)
    print(f"Empresa: Parque Arauco ({COMPANY_ID})")
    print(f"Periodo: {START_DATE} - {END_DATE}")
    print(f"Total de nodos: {len(NODOS_PARQUE_ARAUCO)}")
    print("=" * 60)
    print()
    
    nodos_exitosos = []
    nodos_fallidos = []
    reportes_individuales = []
    
    for i, node_id in enumerate(NODOS_PARQUE_ARAUCO, 1):
        print(f"[{i}/{len(NODOS_PARQUE_ARAUCO)}] Generando reporte para nodo {node_id}...")
        
        cmd = [
            PYTHON_EXE,
            SCRIPT_PATH,
            "--company-id", COMPANY_ID,
            "--node-id", node_id,
            "--start-date", START_DATE,
            "--end-date", END_DATE,
        ]
        
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
            if result.returncode == 0:
                print(f"  [OK] Reporte generado exitosamente")
                nodos_exitosos.append(node_id)
                # Buscar el archivo generado
                # El formato es: reports/{COMPANY_NAME}/REPORTE/{NODE_NAME}_{TIMESTAMP}/Reporte_*.docx
                # Necesitamos encontrar el más reciente
                company_name = get_company_name(COMPANY_ID).replace(" ", "_")
                node_name = get_node_name(node_id).replace(" ", "_")
                report_dir = Path("reports") / company_name / "REPORTE"
                if report_dir.exists():
                    # Buscar la carpeta más reciente que contenga el nombre del nodo
                    matching_dirs = [d for d in report_dir.iterdir() if d.is_dir() and node_name in d.name]
                    if matching_dirs:
                        latest_dir = max(matching_dirs, key=lambda d: d.stat().st_mtime)
                        report_files = list(latest_dir.glob("Reporte_*.docx"))
                        if report_files:
                            reportes_individuales.append(report_files[0])
            else:
                error_msg = result.stderr[:200] if result.stderr else result.stdout[:200]
                print(f"  [ERROR] {error_msg}")
                nodos_fallidos.append(node_id)
        except subprocess.TimeoutExpired:
            print(f"  [ERROR] Timeout al generar reporte")
            nodos_fallidos.append(node_id)
        except Exception as e:
            print(f"  [ERROR] {e}")
            nodos_fallidos.append(node_id)
        
        print()
    
    print("=" * 60)
    print("RESUMEN DE REPORTES INDIVIDUALES")
    print("=" * 60)
    print(f"Exitosos: {len(nodos_exitosos)}")
    print(f"Fallidos: {len(nodos_fallidos)}")
    
    if nodos_fallidos:
        print(f"\nNodos con errores: {', '.join(nodos_fallidos)}")
    
    print()
    print("=" * 60)
    print("GENERANDO REPORTE AGREGADO")
    print("=" * 60)
    
    reporte_agregado_path = None
    
    # Generar reporte agregado solo con los nodos exitosos
    if nodos_exitosos:
        print(f"Generando reporte agregado con {len(nodos_exitosos)} nodos...")
        
        try:
            reporte_agregado_path = generate_aggregated_report(
                COMPANY_ID,
                nodos_exitosos,
                START_DATE,
                END_DATE
            )
            print(f"[OK] Reporte agregado generado exitosamente:")
            print(f"  {reporte_agregado_path}")
        except Exception as e:
            print(f"[ERROR] Error al generar reporte agregado: {e}")
            import traceback
            traceback.print_exc()
    else:
        print("No se pueden generar reportes agregados porque no hay nodos exitosos.")
    
    # Enviar correos si se solicita
    if enviar_correo:
        company_name = get_company_name(COMPANY_ID)
        
        # Enviar reportes individuales
        if reportes_individuales:
            print()
            print("=" * 60)
            print(f"ENVIANDO REPORTES INDIVIDUALES A {destinatario}")
            print("=" * 60)
            
            for i, reporte_path in enumerate(reportes_individuales, 1):
                if reporte_path.exists():
                    node_id = nodos_exitosos[i-1] if i-1 < len(nodos_exitosos) else "unknown"
                    node_name = get_node_name(node_id)
                    print(f"[{i}/{len(reportes_individuales)}] Enviando reporte de {node_name}...")
                    
                    exito = enviar_reporte_por_correo(
                        reporte_path=reporte_path,
                        destinatario=destinatario,
                        smtp_servidor=SMTP_SERVIDOR,
                        smtp_puerto=SMTP_PUERTO,
                        smtp_usuario=SMTP_USUARIO,
                        smtp_password=SMTP_PASSWORD,
                        company_name=company_name,
                        node_name=node_name,
                        start_date="04-12-25",
                        end_date="06-12-25",
                    )
                    
                    if exito:
                        print(f"  [OK] Reporte de {node_name} enviado exitosamente")
                    else:
                        print(f"  [ERROR] Fallo al enviar reporte de {node_name}")
        
        # Enviar reporte agregado
        if reporte_agregado_path and reporte_agregado_path.exists():
            print()
            print("=" * 60)
            print(f"ENVIANDO REPORTE AGREGADO A {destinatario}")
            print("=" * 60)
            print(f"Enviando reporte agregado...")
            
            exito = enviar_reporte_por_correo(
                reporte_path=reporte_agregado_path,
                destinatario=destinatario,
                smtp_servidor=SMTP_SERVIDOR,
                smtp_puerto=SMTP_PUERTO,
                smtp_usuario=SMTP_USUARIO,
                smtp_password=SMTP_PASSWORD,
                company_name=company_name,
                node_name=None,  # Es reporte agregado, no tiene un nodo específico
                start_date="04-12-25",
                end_date="06-12-25",
            )
            
            if exito:
                print()
                print("=" * 60)
                print("[OK] REPORTE AGREGADO ENVIADO EXITOSAMENTE")
                print("=" * 60)
                print(f"Destinatario: {destinatario}")
                print(f"Reporte: {reporte_agregado_path.name}")
            else:
                print()
                print("=" * 60)
                print("[ERROR] FALLO EL ENVÍO DEL REPORTE AGREGADO")
                print("=" * 60)
    
    print()
    print("=" * 60)
    print("PROCESO COMPLETADO")
    print("=" * 60)
    print()
    print("Resumen:")
    print(f"  - Reportes individuales generados: {len(nodos_exitosos)}/{len(NODOS_PARQUE_ARAUCO)}")
    print(f"  - Reporte agregado: {'Sí' if reporte_agregado_path else 'No'}")
    if enviar_correo:
        print(f"  - Reportes individuales enviados: {len(reportes_individuales)}")
        print(f"  - Reporte agregado enviado: {'Sí' if reporte_agregado_path and reporte_agregado_path.exists() else 'No'}")
    else:
        print(f"  - Correo enviado: No (use --enviar-correo para habilitar)")

if __name__ == "__main__":
    main()


"""Script para generar reportes de todos los nodos de Puente Alto (diciembre) y enviar por correo."""

import argparse
import subprocess
import sys
from pathlib import Path
from generar_reporte_word import generate_aggregated_report, enviar_reporte_por_correo, get_company_name, get_node_name
import requests

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

ENTITY_BASE_URL = "http://104.248.53.141:7001/wes/api/acl-entities/v1"

# ID de la empresa Puente Alto
COMPANY_ID = "000010"
COMPANY_NAME = "Corporación Puente Alto"

# Periodo: Diciembre 2025
START_DATE = "01122025"  # 01 de diciembre 2025
END_DATE = "31122025"    # 31 de diciembre 2025

# Configuración de correo
SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_PASSWORD = "vxbynfpoehbweelj"  # Contraseña de aplicación
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

# Destinatarios
DESTINATARIOS = [
    "anibal.aoperaciones@wes.cl",
    "joseotarola@wes.cl"
]

PYTHON_EXE = r"C:\Users\joseo\AppData\Local\Programs\Python\Python314\python.exe"
SCRIPT_PATH = "generar_reporte_word.py"


def obtener_nodos_empresa(company_id: str) -> list[str]:
    """Obtiene todos los nodos de una empresa."""
    try:
        url = f"{ENTITY_BASE_URL}/companies/{company_id}"
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            data = response.json()
            nodes = data.get('nodes', [])
            node_ids = [node.get('nodeId') for node in nodes if node.get('nodeId')]
            return node_ids
        else:
            print(f"[ERROR] No se pudo obtener información de la empresa: {response.status_code}")
            return []
    except Exception as e:
        print(f"[ERROR] Error al obtener nodos de la empresa: {e}")
        return []


def main():
    # Parsear argumentos de línea de comandos
    parser = argparse.ArgumentParser(description="Genera reportes de todos los nodos de Puente Alto (diciembre)")
    parser.add_argument(
        "--enviar-correo",
        action="store_true",
        default=True,  # Por defecto SÍ enviar correo
        help="Enviar los reportes por correo electrónico"
    )
    args = parser.parse_args()
    
    enviar_correo = args.enviar_correo
    
    print("=" * 60)
    print("GENERANDO REPORTES PARA PUENTE ALTO - DICIEMBRE 2025")
    print("=" * 60)
    print(f"Empresa: {COMPANY_NAME} ({COMPANY_ID})")
    print(f"Periodo: {START_DATE} - {END_DATE}")
    print("=" * 60)
    print()
    
    # Obtener todos los nodos de la empresa
    print("[INFO] Obteniendo nodos de la empresa...")
    nodos_empresa = obtener_nodos_empresa(COMPANY_ID)
    
    if not nodos_empresa:
        print("[ERROR] No se encontraron nodos para la empresa. Abortando.")
        sys.exit(1)
    
    print(f"[OK] Se encontraron {len(nodos_empresa)} nodos:")
    for node_id in nodos_empresa:
        node_name = get_node_name(node_id)
        print(f"  - {node_id}: {node_name}")
    print()
    
    nodos_exitosos = []
    nodos_fallidos = []
    reportes_individuales = []
    
    # Generar reportes individuales
    print("=" * 60)
    print("GENERANDO REPORTES INDIVIDUALES")
    print("=" * 60)
    
    for i, node_id in enumerate(nodos_empresa, 1):
        node_name = get_node_name(node_id)
        print(f"[{i}/{len(nodos_empresa)}] Generando reporte para nodo {node_id} ({node_name})...")
        
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
                company_name = get_company_name(COMPANY_ID).replace(" ", "_")
                node_name_clean = node_name.replace(" ", "_")
                report_dir = Path("reports") / company_name / "REPORTE"
                if report_dir.exists():
                    # Buscar la carpeta más reciente que contenga el nombre del nodo
                    matching_dirs = [d for d in report_dir.iterdir() if d.is_dir() and node_name_clean in d.name]
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
    
    # Generar reporte agregado solo si hay más de un nodo exitoso
    if len(nodos_exitosos) > 1:
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
        print(f"[INFO] Solo hay {len(nodos_exitosos)} nodo(s) exitoso(s). No se generará reporte agregado.")
    
    # Enviar correos si se solicita
    if enviar_correo:
        company_name = get_company_name(COMPANY_ID)
        
        # Enviar reportes individuales a cada destinatario
        if reportes_individuales:
            print()
            print("=" * 60)
            print(f"ENVIANDO REPORTES INDIVIDUALES")
            print("=" * 60)
            
            for destinatario in DESTINATARIOS:
                print(f"\n[INFO] Enviando reportes individuales a {destinatario}...")
                
                for i, reporte_path in enumerate(reportes_individuales, 1):
                    if reporte_path.exists():
                        node_id = nodos_exitosos[i-1] if i-1 < len(nodos_exitosos) else "unknown"
                        node_name = get_node_name(node_id)
                        print(f"  [{i}/{len(reportes_individuales)}] Enviando reporte de {node_name}...")
                        
                        exito = enviar_reporte_por_correo(
                            reporte_path=reporte_path,
                            destinatario=destinatario,
                            smtp_servidor=SMTP_SERVIDOR,
                            smtp_puerto=SMTP_PUERTO,
                            smtp_usuario=SMTP_USUARIO,
                            smtp_password=SMTP_PASSWORD,
                            company_name=company_name,
                            node_name=node_name,
                            start_date="01-12-25",
                            end_date="31-12-25",
                        )
                        
                        if exito:
                            print(f"    [OK] Reporte de {node_name} enviado exitosamente")
                        else:
                            print(f"    [ERROR] Fallo al enviar reporte de {node_name}")
        
        # Enviar reporte agregado a cada destinatario
        if reporte_agregado_path and reporte_agregado_path.exists():
            print()
            print("=" * 60)
            print(f"ENVIANDO REPORTE AGREGADO")
            print("=" * 60)
            
            for destinatario in DESTINATARIOS:
                print(f"\n[INFO] Enviando reporte agregado a {destinatario}...")
                
                exito = enviar_reporte_por_correo(
                    reporte_path=reporte_agregado_path,
                    destinatario=destinatario,
                    smtp_servidor=SMTP_SERVIDOR,
                    smtp_puerto=SMTP_PUERTO,
                    smtp_usuario=SMTP_USUARIO,
                    smtp_password=SMTP_PASSWORD,
                    company_name=company_name,
                    node_name=None,  # Es reporte agregado, no tiene un nodo específico
                    start_date="01-12-25",
                    end_date="31-12-25",
                )
                
                if exito:
                    print(f"  [OK] Reporte agregado enviado exitosamente a {destinatario}")
                else:
                    print(f"  [ERROR] Fallo al enviar reporte agregado a {destinatario}")
    
    print()
    print("=" * 60)
    print("PROCESO COMPLETADO")
    print("=" * 60)
    print()
    print("Resumen:")
    print(f"  - Reportes individuales generados: {len(nodos_exitosos)}/{len(nodos_empresa)}")
    print(f"  - Reporte agregado: {'Sí' if reporte_agregado_path else 'No'}")
    if enviar_correo:
        print(f"  - Reportes individuales enviados: {len(reportes_individuales) * len(DESTINATARIOS)}")
        print(f"  - Reporte agregado enviado: {'Sí' if reporte_agregado_path and reporte_agregado_path.exists() else 'No'}")
        print(f"  - Destinatarios: {', '.join(DESTINATARIOS)}")
    else:
        print(f"  - Correo enviado: No (use --enviar-correo para habilitar)")


if __name__ == "__main__":
    main()

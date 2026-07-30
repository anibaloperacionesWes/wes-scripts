"""Script para generar reportes de todos los nodos de Fundo Zapallar y opcionalmente enviar el agregado por correo."""

import argparse
import subprocess
import sys
from pathlib import Path
from generar_reporte_word import generate_aggregated_report, enviar_reporte_por_correo, get_company_name

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

from exclusiones_reportes import FUNDO_ZAPALLAR_NODE_IDS

NODOS_FUNDO_ZAPALLAR = list(FUNDO_ZAPALLAR_NODE_IDS)

COMPANY_ID = "000027"
START_DATE = "01122025"  # 01 de diciembre 2025
END_DATE = "31122025"    # 31 de diciembre 2025

# Configuración de correo
SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_PASSWORD = "vxbynfpoehbweelj"  # Contraseña de aplicación
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587
DESTINATARIO = "silvanaaraya.rojas@gmail.com"

PYTHON_EXE = r"C:\Users\joseo\AppData\Local\Programs\Python\Python314\python.exe"
SCRIPT_PATH = "generar_reporte_word.py"

def main():
    # Parsear argumentos de línea de comandos
    parser = argparse.ArgumentParser(description="Genera reportes de todos los nodos de Fundo Zapallar")
    parser.add_argument(
        "--enviar-correo",
        action="store_true",
        help="Enviar el reporte agregado por correo electrónico"
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
    print("GENERANDO REPORTES INDIVIDUALES PARA FUNDO ZAPALLAR")
    print("=" * 60)
    print(f"Empresa: Fundo Zapallar ({COMPANY_ID})")
    print(f"Periodo: {START_DATE} - {END_DATE}")
    print(f"Total de nodos: {len(NODOS_FUNDO_ZAPALLAR)}")
    print("=" * 60)
    print()
    
    nodos_exitosos = []
    nodos_fallidos = []
    
    for i, node_id in enumerate(NODOS_FUNDO_ZAPALLAR, 1):
        print(f"[{i}/{len(NODOS_FUNDO_ZAPALLAR)}] Generando reporte para nodo {node_id}...")
        
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
    
    # Enviar correo solo si se solicita
    if enviar_correo:
        print()
        print("=" * 60)
        print("ENVIANDO CORREO CON REPORTE AGREGADO")
        print("=" * 60)
        
        if reporte_agregado_path and reporte_agregado_path.exists():
            print(f"Enviando reporte agregado a {destinatario}...")
            
            company_name = get_company_name(COMPANY_ID)
            
            # Convertir fechas de ddMMyyyy a dd-mm-yy para el correo
            from datetime import datetime
            start_dt = datetime.strptime(START_DATE, "%d%m%Y")
            end_dt = datetime.strptime(END_DATE, "%d%m%Y")
            start_date_str = start_dt.strftime("%d-%m-%y")
            end_date_str = end_dt.strftime("%d-%m-%y")
            
            exito = enviar_reporte_por_correo(
                reporte_path=reporte_agregado_path,
                destinatario=destinatario,
                smtp_servidor=SMTP_SERVIDOR,
                smtp_puerto=SMTP_PUERTO,
                smtp_usuario=SMTP_USUARIO,
                smtp_password=SMTP_PASSWORD,
                company_name=company_name,
                node_name=None,  # Es reporte agregado, no tiene un nodo específico
                start_date=start_date_str,
                end_date=end_date_str,
            )
            
            if exito:
                print()
                print("=" * 60)
                print("[OK] CORREO ENVIADO EXITOSAMENTE")
                print("=" * 60)
                print(f"Destinatario: {destinatario}")
                print(f"Reporte: {reporte_agregado_path.name}")
            else:
                print()
                print("=" * 60)
                print("[ERROR] FALLO EL ENVÍO DEL CORREO")
                print("=" * 60)
        else:
            print("[ERROR] No se puede enviar el correo porque el reporte agregado no existe.")
    
    print()
    print("=" * 60)
    print("PROCESO COMPLETADO")
    print("=" * 60)
    print()
    print("Resumen:")
    print(f"  - Reportes individuales generados: {len(nodos_exitosos)}/{len(NODOS_FUNDO_ZAPALLAR)}")
    print(f"  - Reporte agregado: {'Sí' if reporte_agregado_path else 'No'}")
    if reporte_agregado_path and reporte_agregado_path.exists():
        print(f"  - Ubicación reporte agregado: {reporte_agregado_path}")
    if enviar_correo:
        print(f"  - Correo enviado: {'Sí' if reporte_agregado_path and reporte_agregado_path.exists() else 'No'}")
    else:
        print(f"  - Correo enviado: No (use --enviar-correo para habilitar)")

if __name__ == "__main__":
    main()


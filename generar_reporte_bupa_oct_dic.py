"""Script para generar reporte agregado de BUPA desde octubre hasta hoy y enviar por correo."""

import subprocess
import sys
from datetime import datetime
from pathlib import Path
from generar_reporte_word import generate_aggregated_report, enviar_reporte_por_correo, get_company_name

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

# Todos los nodos de BUPA (empresa 000029)
NODOS_BUPA = [
    "000029-01",  # Llenado de Estanques
    "000029-02",  # Torre A
    "000029-03",  # Torre B1
    "000029-04",  # Torre B2
    "000029-05",  # Torre C
    "000029-06",  # Central Térmica
]

COMPANY_ID = "000029"
START_DATE = "03/10/2025"  # 03 de octubre 2025
END_DATE = datetime.now().strftime("%d/%m/%Y")  # Fecha actual

# Configuración de correo
SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_PASSWORD = "vxbynfpoehbweelj"  # Contraseña de aplicación
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587
DESTINATARIOS = [
    "agente.ia@wes.cl",
    "anibal.aoperaciones@wes.cl",
]

PYTHON_EXE = r"C:\Users\joseo\AppData\Local\Programs\Python\Python314\python.exe"
SCRIPT_PATH = "generar_reporte_word.py"

def main():
    print("=" * 60)
    print("GENERANDO REPORTES INDIVIDUALES PARA BUPA")
    print("=" * 60)
    print(f"Empresa: BUPA ({COMPANY_ID})")
    print(f"Periodo: {START_DATE} - {END_DATE}")
    print(f"Total de nodos: {len(NODOS_BUPA)}")
    print("=" * 60)
    print()
    
    nodos_exitosos = []
    nodos_fallidos = []
    
    for i, node_id in enumerate(NODOS_BUPA, 1):
        print(f"[{i}/{len(NODOS_BUPA)}] Generando reporte para nodo {node_id}...")
        
        cmd = [
            PYTHON_EXE,
            SCRIPT_PATH,
            "--company-id", COMPANY_ID,
            "--node-id", node_id,
            "--start-date", START_DATE,
            "--end-date", END_DATE,
        ]
        
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
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
    
    print()
    print("=" * 60)
    print("ENVIANDO CORREOS CON REPORTE AGREGADO")
    print("=" * 60)
    
    if reporte_agregado_path and reporte_agregado_path.exists():
        company_name = get_company_name(COMPANY_ID)
        
        # Formatear fechas para el correo
        start_dt = datetime.strptime(START_DATE, "%d/%m/%Y")
        end_dt = datetime.strptime(END_DATE, "%d/%m/%Y")
        start_date_str = start_dt.strftime("%d-%m-%y")
        end_date_str = end_dt.strftime("%d-%m-%y")
        
        correos_enviados = 0
        correos_fallidos = 0
        
        for destinatario in DESTINATARIOS:
            print(f"Enviando reporte agregado a {destinatario}...")
            
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
                print(f"  [OK] Correo enviado exitosamente a {destinatario}")
                correos_enviados += 1
            else:
                print(f"  [ERROR] Fallo al enviar correo a {destinatario}")
                correos_fallidos += 1
            
            print()
        
        print("=" * 60)
        print("RESUMEN DE ENVÍO DE CORREOS")
        print("=" * 60)
        print(f"Correos enviados exitosamente: {correos_enviados}/{len(DESTINATARIOS)}")
        print(f"Correos fallidos: {correos_fallidos}")
        
        if correos_enviados > 0:
            print()
            print("=" * 60)
            print("[OK] PROCESO COMPLETADO")
            print("=" * 60)
    else:
        print("[ERROR] No se puede enviar el correo porque el reporte agregado no existe.")
    
    print()
    print("=" * 60)
    print("RESUMEN FINAL")
    print("=" * 60)
    print(f"  - Reportes individuales generados: {len(nodos_exitosos)}/{len(NODOS_BUPA)}")
    print(f"  - Reporte agregado: {'Sí' if reporte_agregado_path else 'No'}")
    if reporte_agregado_path and reporte_agregado_path.exists():
        print(f"  - Correos enviados: {correos_enviados}/{len(DESTINATARIOS)}")
        print(f"  - Destinatarios: {', '.join(DESTINATARIOS)}")

if __name__ == "__main__":
    main()
















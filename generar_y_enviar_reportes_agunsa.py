"""Script para generar reportes de todos los nodos de AGUNSA y opcionalmente enviar el agregado por correo.

Orden: primero un Word por punto; después el agregado (nunca al revés).
"""

import argparse
import sys
from pathlib import Path
from generar_reporte_word import (
    generate_aggregated_report,
    generate_report,
    enviar_reporte_por_correo,
    get_company_name,
)

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

# Todos los nodos de AGUNSA (empresa 000020)
NODOS_AGUNSA = [
    "000020-01",  # Deposito
    "000020-02",  # Modulo D
    "000020-03",  # Módulo ABC
    "000020-04",  # Módulo E
    "000020-05",  # Intermodal-San Antonio
]

COMPANY_ID = "000020"
DEFAULT_START_DATE = "01/11/2025"
DEFAULT_END_DATE = "07/12/2025"

# Configuración de correo
SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_PASSWORD = "vxbynfpoehbweelj"  # Contraseña de aplicación
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587
DESTINATARIO = "silvanaaraya.rojas@gmail.com"

def main():
    # Parsear argumentos de línea de comandos
    parser = argparse.ArgumentParser(description="Genera reportes de todos los nodos de AGUNSA")
    parser.add_argument("--start-date", default=DEFAULT_START_DATE, help="Inicio (dd/mm/aaaa)")
    parser.add_argument("--end-date", default=DEFAULT_END_DATE, help="Término (dd/mm/aaaa)")
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
    start_date = args.start_date
    end_date = args.end_date
    root = Path(__file__).resolve().parent
    reports_dir = root / "reports"

    print("=" * 60)
    print("GENERANDO REPORTES INDIVIDUALES PARA AGUNSA")
    print("=" * 60)
    print(f"Empresa: AGUNSA ({COMPANY_ID})")
    print(f"Periodo: {start_date} - {end_date}")
    print(f"Total de nodos: {len(NODOS_AGUNSA)}")
    print("=" * 60)
    print()
    
    nodos_exitosos = []
    nodos_fallidos = []
    
    for i, node_id in enumerate(NODOS_AGUNSA, 1):
        print(f"[{i}/{len(NODOS_AGUNSA)}] Generando reporte para nodo {node_id}...")
        
        gen_args = argparse.Namespace(
            company_id=COMPANY_ID,
            node_id=node_id,
            start_date=start_date,
            end_date=end_date,
            output_dir=str(reports_dir),
            enviar_correo=False,
        )
        try:
            out = generate_report(gen_args)
            if out:
                print(f"  [OK] {out}")
                nodos_exitosos.append(node_id)
            else:
                print(f"  [ERROR] generate_report no devolvió ruta")
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
                start_date,
                end_date,
                output_dir=str(reports_dir),
                generate_ppt=False,
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
            
            exito = enviar_reporte_por_correo(
                reporte_path=reporte_agregado_path,
                destinatario=destinatario,
                smtp_servidor=SMTP_SERVIDOR,
                smtp_puerto=SMTP_PUERTO,
                smtp_usuario=SMTP_USUARIO,
                smtp_password=SMTP_PASSWORD,
                company_name=company_name,
                node_name=None,  # Es reporte agregado, no tiene un nodo específico
                start_date=start_date.replace("/", "-"),
                end_date=end_date.replace("/", "-"),
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
    print(f"  - Reportes individuales generados: {len(nodos_exitosos)}/{len(NODOS_AGUNSA)}")
    print(f"  - Reporte agregado: {'Sí' if reporte_agregado_path else 'No'}")
    if reporte_agregado_path and reporte_agregado_path.exists():
        print(f"  - Ubicación reporte agregado: {reporte_agregado_path}")
    if enviar_correo:
        print(f"  - Correo enviado: {'Sí' if reporte_agregado_path and reporte_agregado_path.exists() else 'No'}")
    else:
        print(f"  - Correo enviado: No (use --enviar-correo para habilitar)")

if __name__ == "__main__":
    main()


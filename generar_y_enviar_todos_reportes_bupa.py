"""Script para generar todos los reportes de BUPA y enviarlos por correo con explicación."""

import subprocess
import sys
import smtplib
from datetime import datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from generar_reporte_word import generate_aggregated_report, get_company_name, get_node_name, convertir_word_a_pdf

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
SMTP_PASSWORD = "vxbynfpoehbweelj"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587
DESTINATARIOS = [
    "agente.ia@wes.cl",
    "benjamingumucio@wes.cl",
    "diegocarrasco@wes.cl",
    "juanlopez@wes.cl",
]

PYTHON_EXE = r"C:\Users\joseo\AppData\Local\Programs\Python\Python314\python.exe"
SCRIPT_PATH = "generar_reporte_word.py"

def enviar_reportes_por_correo(
    reportes_paths: list[Path],
    destinatario: str,
    company_name: str,
    start_date: str,
    end_date: str,
    total_nodos: int,
) -> bool:
    """Envía múltiples reportes por correo con explicación."""
    try:
        msg = MIMEMultipart()
        msg["From"] = SMTP_USUARIO
        msg["To"] = destinatario
        
        # Crear asunto
        start_dt = datetime.strptime(start_date, "%d/%m/%Y")
        end_dt = datetime.strptime(end_date, "%d/%m/%Y")
        start_date_str = start_dt.strftime("%d-%m-%y")
        end_date_str = end_dt.strftime("%d-%m-%y")
        
        msg["Subject"] = f"Reportes de Consumo BUPA - {start_date_str} a {end_date_str}"
        
        # Crear cuerpo del mensaje con explicación
        cuerpo = f"""
Estimado/a,

Se adjuntan los reportes de consumo y análisis de agua para BUPA correspondientes al periodo del {start_date_str} al {end_date_str}.

CONTENIDO DEL ENVÍO:
- Reportes individuales por punto de monitoreo ({total_nodos} reportes)
- Reporte agregado consolidado de todos los puntos

METODOLOGÍA DE CONFECCIÓN DEL INFORME:

Los reportes fueron generados mediante un proceso automatizado que incluye:

1. RECOPILACIÓN DE DATOS:
   - Se obtuvieron las mediciones de consumo de agua de cada punto de monitoreo
   - Los datos se recopilaron con frecuencia horaria para el periodo completo
   - Se validaron y procesaron los datos para asegurar su integridad

2. ANÁLISIS Y CÁLCULOS:
   - Consumo diario y total del periodo para cada punto
   - Identificación de patrones de consumo horario
   - Detección de alertas de fuga mediante algoritmos de análisis
   - Cálculo de métricas estadísticas (promedios, máximos, mínimos)

3. GENERACIÓN DE VISUALIZACIONES:
   - Gráficos de consumo diario y horario
   - Comparaciones entre periodos cuando aplica
   - Representaciones visuales de alertas y anomalías detectadas

4. CONSOLIDACIÓN:
   - Los reportes individuales muestran el detalle por cada punto de monitoreo
   - El reporte agregado consolida la información de todos los puntos
   - Se incluyen tablas resumen con métricas clave

5. FORMATO Y PRESENTACIÓN:
   - Los reportes se generan en formato Word con gráficos y tablas
   - Se incluyen secciones de resumen ejecutivo, análisis detallado y conclusiones
   - Los gráficos facilitan la interpretación visual de los datos

PUNTOS DE MONITOREO INCLUIDOS:
"""
        for node_id in NODOS_BUPA:
            node_name = get_node_name(node_id)
            cuerpo += f"- {node_name} ({node_id})\n"
        
        cuerpo += f"""
Periodo analizado: {start_date_str} a {end_date_str}

Los reportes contienen información detallada sobre consumo, patrones de uso, detección de fugas y recomendaciones para optimización del uso del agua.

Saludos cordiales,
Sistema WES - Water Efficiency System
"""
        
        msg.attach(MIMEText(cuerpo, "plain", "utf-8"))
        
        # Adjuntar todos los reportes
        pdf_paths = []
        for reporte_path in reportes_paths:
            if not reporte_path.exists():
                print(f"  [ADVERTENCIA] Reporte no encontrado: {reporte_path}")
                continue
            
            # Intentar convertir a PDF
            try:
                pdf_path = convertir_word_a_pdf(reporte_path)
                if pdf_path and pdf_path.exists():
                    with open(pdf_path, "rb") as f:
                        adjunto = MIMEApplication(f.read(), _subtype="pdf")
                        adjunto.add_header(
                            "Content-Disposition",
                            "attachment",
                            filename=reporte_path.stem + ".pdf"
                        )
                        msg.attach(adjunto)
                    pdf_paths.append(pdf_path)
                else:
                    # Si falla, adjuntar Word
                    with open(reporte_path, "rb") as f:
                        adjunto = MIMEApplication(f.read(), _subtype="docx")
                        adjunto.add_header(
                            "Content-Disposition",
                            "attachment",
                            filename=reporte_path.name
                        )
                        msg.attach(adjunto)
            except Exception as e:
                print(f"  [ADVERTENCIA] Error al procesar {reporte_path.name}: {e}")
                # Adjuntar Word original
                with open(reporte_path, "rb") as f:
                    adjunto = MIMEApplication(f.read(), _subtype="docx")
                    adjunto.add_header(
                        "Content-Disposition",
                        "attachment",
                        filename=reporte_path.name
                    )
                    msg.attach(adjunto)
        
        # Enviar correo
        with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
            server.starttls()
            server.login(SMTP_USUARIO, SMTP_PASSWORD)
            server.send_message(msg)
        
        # Limpiar PDFs temporales
        for pdf_path in pdf_paths:
            try:
                if pdf_path.exists():
                    pdf_path.unlink()
            except:
                pass
        
        return True
        
    except Exception as e:
        print(f"  [ERROR] {e}")
        import traceback
        traceback.print_exc()
        return False

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
    reportes_individuales = []
    
    for i, node_id in enumerate(NODOS_BUPA, 1):
        node_name = get_node_name(node_id)
        print(f"[{i}/{len(NODOS_BUPA)}] Generando reporte para {node_name} ({node_id})...")
        
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
                
                # Buscar el archivo generado
                company_name_clean = get_company_name(COMPANY_ID).replace(" ", "_")
                node_name_clean = node_name.replace(" ", "_")
                report_dir = Path("reports") / company_name_clean / "REPORTE"
                
                if report_dir.exists():
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
    print("ENVIANDO CORREOS CON TODOS LOS REPORTES")
    print("=" * 60)
    
    # Preparar lista de todos los reportes
    todos_reportes = reportes_individuales.copy()
    if reporte_agregado_path and reporte_agregado_path.exists():
        todos_reportes.append(reporte_agregado_path)
    
    if not todos_reportes:
        print("[ERROR] No hay reportes para enviar")
        return
    
    company_name = get_company_name(COMPANY_ID)
    correos_enviados = 0
    correos_fallidos = 0
    
    for destinatario in DESTINATARIOS:
        print(f"Enviando {len(todos_reportes)} reportes a {destinatario}...")
        
        exito = enviar_reportes_por_correo(
            reportes_paths=todos_reportes,
            destinatario=destinatario,
            company_name=company_name,
            start_date=START_DATE,
            end_date=END_DATE,
            total_nodos=len(nodos_exitosos),
        )
        
        if exito:
            print(f"  [OK] Correo enviado exitosamente")
            correos_enviados += 1
        else:
            print(f"  [ERROR] Fallo al enviar correo")
            correos_fallidos += 1
        
        print()
    
    print("=" * 60)
    print("RESUMEN FINAL")
    print("=" * 60)
    print(f"  - Reportes individuales generados: {len(nodos_exitosos)}/{len(NODOS_BUPA)}")
    print(f"  - Reporte agregado: {'Sí' if reporte_agregado_path else 'No'}")
    print(f"  - Total de reportes enviados: {len(todos_reportes)}")
    print(f"  - Correos enviados exitosamente: {correos_enviados}/{len(DESTINATARIOS)}")
    print(f"  - Correos fallidos: {correos_fallidos}")
    print(f"  - Destinatarios: {', '.join(DESTINATARIOS)}")

if __name__ == "__main__":
    main()
















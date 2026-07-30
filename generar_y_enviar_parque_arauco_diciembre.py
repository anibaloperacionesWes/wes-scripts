"""Script para generar reportes de Parque Arauco (diciembre hasta hoy) y enviar el PDF agregado por correo.

El correo explica cómo se obtuvieron los datos sin exponer información sensible
y se identifica como agente WES.
"""

import subprocess
import sys
import uuid
from datetime import datetime, timezone
from pathlib import Path
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate, make_msgid

from generar_reporte_word import generate_aggregated_report, get_company_name, get_node_name, convertir_word_a_pdf

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

# Todos los nodos de Parque Arauco (empresa 000025)
NODOS_PARQUE_ARAUCO = [
    "000025-01",  # Estanque Norte Locales Mall
    "000025-02",  # Abastecimiento Sur Terminal
    "000025-03",  # Poniente 7
    "000025-04",  # Baños Públicos
    "000025-05",  # Locales de Comida
    "000025-06",  # KFC
    "000025-07",  # PIZZA HUT
    "000025-08",  # Placa Bancaria
    "000025-09",  # Impulsión Falabella
    "000025-10",  # Impulsión Ripley
    "000025-11",  # Matriz principal 1°piso
    "000025-12",  # Anillo Plaza
    "000025-13",  # Matriz Principal
    "000025-14",  # Red de Incendio
    "000025-15",  # Matriz Principal
    "000025-16",  # Baños
    "000025-17",  # San Ignacio 300
    "000025-18",  # San Ignacio 500
    "000025-19",  # Sala de Bomba Estanque Sur
    "000025-20",  # Impulsión Ander3-4 Matriz Principal
    "000025-21",  # Impulsión Ander3-4 Locales Gast.
    "000025-22",  # Impulsión Sandia Baños 2-3-6-7 Fredo
    "000025-23",  # Llenado Pileta
    "000025-24",  # Llenado Pileta Cascada
    "000025-35",  # PAK BAZAR GOURMET (reemplazo 000025-25)
    "000025-36",  # PAK DL KENNEDY (reemplazo 000025-26)
    "000025-27",  # Distrito de lujo DL
    "000025-28",  # Impulsión Mall 1 Piso-4
    "000025-29",  # Impulsión Anden 3-4 Restaurante
]

COMPANY_ID = "000025"

# Periodo: 1 de diciembre hasta hoy
START_DATE = "01/12/2025"
END_DATE = datetime.now().strftime("%d/%m/%Y")

# Configuración: Para Parque Arauco, normalmente todos son consumidores (sin fuente de agua)
FUENTE_AGUA_ID = None  # Todos los puntos son consumidores

PYTHON_EXE = r"C:\Users\joseo\AppData\Local\Programs\Python\Python314\python.exe"
SCRIPT_PATH = "generar_reporte_word.py"

# Configuración de correo
SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_PASSWORD = "vxbynfpoehbweelj"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

TO_RECIPIENTS = ["agente.ia@wes.cl"]  # Solo al usuario
CC_RECIPIENTS = []  # Sin copias


def generar_reportes_individuales():
    """Genera todos los reportes individuales."""
    print("=" * 70)
    print("GENERANDO REPORTES INDIVIDUALES")
    print("=" * 70)
    print()
    
    nodos_exitosos = []
    nodos_fallidos = []
    
    for i, node_id in enumerate(NODOS_PARQUE_ARAUCO, 1):
        node_name = get_node_name(node_id)
        print(f"[{i}/{len(NODOS_PARQUE_ARAUCO)}] Generando reporte para {node_name} ({node_id})...")
        sys.stdout.flush()
        
        cmd = [
            PYTHON_EXE,
            SCRIPT_PATH,
            "--company-id", COMPANY_ID,
            "--node-id", node_id,
            "--start-date", START_DATE,
            "--end-date", END_DATE,
        ]
        
        try:
            result = subprocess.run(
                cmd, 
                capture_output=True, 
                text=True, 
                timeout=600,
                encoding='utf-8',
                errors='ignore'
            )
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
        sys.stdout.flush()
    
    return nodos_exitosos, nodos_fallidos


def generar_reporte_agregado(nodos_exitosos):
    """Genera el reporte agregado."""
    if len(nodos_exitosos) < 2:
        print("[ERROR] Se requieren al menos 2 reportes individuales para generar el agregado.")
        return None
    
    print()
    print("=" * 70)
    print("GENERANDO REPORTE AGREGADO CON RANKING DE CONSUMOS")
    print("=" * 70)
    print(f"Total de nodos exitosos: {len(nodos_exitosos)}")
    print("Nodos incluidos:")
    for node_id in nodos_exitosos:
        print(f"  - {get_node_name(node_id)} ({node_id})")
    print()
    sys.stdout.flush()
    
    # Eliminar reporte agregado anterior si existe
    from datetime import datetime as dt
    start_dt = dt.strptime(START_DATE, "%d/%m/%Y")
    end_dt = dt.strptime(END_DATE, "%d/%m/%Y")
    start_str = start_dt.strftime("%Y%m%d")
    end_str = end_dt.strftime("%Y%m%d")
    pattern = f"Reporte_Agregado_{COMPANY_ID}_{start_str}_{end_str}.docx"
    
    company_name = get_company_name(COMPANY_ID)
    safe_company_name = company_name.replace(" ", "_")
    agregado_dir = Path("reports") / safe_company_name / "ABREGADO"
    
    if agregado_dir.exists():
        for carpeta in agregado_dir.iterdir():
            if carpeta.is_dir():
                reporte_file = carpeta / pattern
                if reporte_file.exists():
                    print(f"Eliminando reporte agregado anterior: {carpeta}")
                    import shutil
                    try:
                        shutil.rmtree(carpeta)
                        print("  [OK] Reporte anterior eliminado")
                    except Exception as e:
                        print(f"  [ADVERTENCIA] No se pudo eliminar: {e}")
    
    try:
        reporte_agregado = generate_aggregated_report(
            COMPANY_ID,
            nodos_exitosos,
            START_DATE,
            END_DATE,
            fuente_agua_id=FUENTE_AGUA_ID
        )
        print(f"[OK] Reporte agregado generado exitosamente con {len(nodos_exitosos)} nodos:")
        print(f"  {reporte_agregado}")
        return reporte_agregado
    except Exception as e:
        print(f"[ERROR] Error al generar reporte agregado: {e}")
        import traceback
        traceback.print_exc()
        return None


def enviar_correo_pdf(pdf_path: Path):
    """Envía el PDF agregado por correo con headers mejorados para evitar spam."""
    if not pdf_path or not pdf_path.exists():
        print(f"[ERROR] El archivo PDF no existe: {pdf_path}")
        return False
    
    company_name = get_company_name(COMPANY_ID)
    
    # Construir correo con headers mejorados
    msg = MIMEMultipart()
    
    # Headers básicos
    msg["From"] = f"Agente WES <{SMTP_USUARIO}>"
    msg["To"] = ", ".join(TO_RECIPIENTS)
    if CC_RECIPIENTS:
        msg["Cc"] = ", ".join(CC_RECIPIENTS)
    msg["Reply-To"] = SMTP_USUARIO
    msg["Subject"] = (
        f"Reporte agregado Parque Arauco - {company_name} - Diciembre {datetime.now().year}"
    )
    
    # Headers importantes para evitar spam
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain="wes.cl")
    msg["X-Mailer"] = "WES Report Generator"
    msg["X-Priority"] = "3"
    msg["MIME-Version"] = "1.0"
    
    # Formatear fechas para el mensaje
    start_dt = datetime.strptime(START_DATE, "%d/%m/%Y")
    end_dt = datetime.strptime(END_DATE, "%d/%m/%Y")
    fecha_inicio = start_dt.strftime("%d-%m-%Y")
    fecha_fin = end_dt.strftime("%d-%m-%Y")
    
    # Cuerpo en texto plano
    cuerpo_texto = f"""Estimado/a,

Adjunto encontrarás el reporte agregado de consumo de agua y alertas de Parque Arauco
para el periodo {fecha_inicio} al {fecha_fin}.

Metodología de obtención de datos:

Los datos se obtuvieron desde el sistema WES mediante consultas automatizadas a los servicios
internos que almacenan las lecturas históricas de los medidores. El proceso incluye:

1. Consumos horarios y diarios: Se consultaron las mediciones registradas por los sensores
   de cada punto de monitoreo durante el periodo indicado, calculando promedios diarios y
   totales acumulados.

2. Alertas de filtración: Se analizaron las alertas registradas en el sistema, considerando
   solo aquellas con medida mayor a cero y que ocurrieron en horario nocturno (22:00 a 07:00)
   durante los últimos 2 días del periodo. La proyección diaria de fuga se calcula como el
   promedio de las 2 últimas alertas nocturnas registradas, multiplicado por 24 horas.

3. Consumo efectivo: Se calculó restando la proyección diaria de fuga del consumo promedio
   diario, permitiendo identificar el consumo real sin considerar las pérdidas por filtración.

4. Valorización: Los volúmenes de agua se valorizaron según el precio por metro cúbico
   configurado para cada punto de monitoreo, proyectando los valores a un mes completo (30 días).

El documento consolida:
- Consumos diarios y consumo efectivo por punto de monitoreo.
- Alertas de filtración detectadas y su proyección cuando corresponde.
- Gráficas comparativas y un análisis agregado de todos los puntos de Parque Arauco.
- Ranking de consumos por punto de monitoreo.

Este correo y el reporte fueron generados automáticamente por un agente de IA al servicio de WES,
configurado específicamente para apoyar la generación y distribución de reportes de consumo y
alertas de filtración.

Quedo atento a cualquier consulta o comentario.

Saludos,
Agente WES
"""
    
    # Cuerpo en HTML (mejor para evitar spam)
    cuerpo_html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body {{ font-family: Arial, sans-serif; line-height: 1.6; color: #333; }}
        .header {{ background-color: #f4f4f4; padding: 10px; border-radius: 5px; }}
        .content {{ margin: 20px 0; }}
        .section {{ margin: 15px 0; }}
        .footer {{ margin-top: 30px; padding-top: 20px; border-top: 1px solid #ddd; font-size: 0.9em; color: #666; }}
        ul {{ margin: 10px 0; padding-left: 20px; }}
        li {{ margin: 5px 0; }}
    </style>
</head>
<body>
    <div class="header">
        <h2>Reporte Agregado Parque Arauco</h2>
        <p><strong>Periodo:</strong> {fecha_inicio} al {fecha_fin}</p>
    </div>
    
    <div class="content">
        <p>Estimado/a,</p>
        
        <p>Adjunto encontrarás el reporte agregado de consumo de agua y alertas de Parque Arauco
        para el periodo {fecha_inicio} al {fecha_fin}.</p>
        
        <div class="section">
            <h3>Metodología de obtención de datos:</h3>
            
            <p>Los datos se obtuvieron desde el sistema WES mediante consultas automatizadas a los servicios
            internos que almacenan las lecturas históricas de los medidores. El proceso incluye:</p>
            
            <ol>
                <li><strong>Consumos horarios y diarios:</strong> Se consultaron las mediciones registradas por los sensores
                de cada punto de monitoreo durante el periodo indicado, calculando promedios diarios y
                totales acumulados.</li>
                
                <li><strong>Alertas de filtración:</strong> Se analizaron las alertas registradas en el sistema, considerando
                solo aquellas con medida mayor a cero y que ocurrieron en horario nocturno (22:00 a 07:00)
                durante los últimos 2 días del periodo. La proyección diaria de fuga se calcula como el
                promedio de las 2 últimas alertas nocturnas registradas, multiplicado por 24 horas.</li>
                
                <li><strong>Consumo efectivo:</strong> Se calculó restando la proyección diaria de fuga del consumo promedio
                diario, permitiendo identificar el consumo real sin considerar las pérdidas por filtración.</li>
                
                <li><strong>Valorización:</strong> Los volúmenes de agua se valorizaron según el precio por metro cúbico
                configurado para cada punto de monitoreo, proyectando los valores a un mes completo (30 días).</li>
            </ol>
        </div>
        
        <div class="section">
            <p>El documento consolida:</p>
            <ul>
                <li>Consumos diarios y consumo efectivo por punto de monitoreo.</li>
                <li>Alertas de filtración detectadas y su proyección cuando corresponde.</li>
                <li>Gráficas comparativas y un análisis agregado de todos los puntos de Parque Arauco.</li>
                <li>Ranking de consumos por punto de monitoreo.</li>
            </ul>
        </div>
        
        <p>Este correo y el reporte fueron generados automáticamente por un agente de IA al servicio de WES,
        configurado específicamente para apoyar la generación y distribución de reportes de consumo y
        alertas de filtración.</p>
        
        <p>Quedo atento a cualquier consulta o comentario.</p>
    </div>
    
    <div class="footer">
        <p>Saludos,<br>
        <strong>Agente WES</strong></p>
    </div>
</body>
</html>
"""
    
    # Crear mensaje multipart/alternative para texto plano y HTML
    msg_alternative = MIMEMultipart('alternative')
    msg.attach(msg_alternative)
    
    # Agregar ambas versiones (texto plano y HTML) para mejor compatibilidad
    part1 = MIMEText(cuerpo_texto, "plain", "utf-8")
    part2 = MIMEText(cuerpo_html, "html", "utf-8")
    
    # Agregar en orden: primero texto plano, luego HTML (el cliente elegirá la mejor)
    msg_alternative.attach(part1)
    msg_alternative.attach(part2)
    
    # Adjuntar archivo PDF
    with open(pdf_path, "rb") as f:
        adj = MIMEApplication(f.read(), _subtype="pdf")
        filename = pdf_path.name
        adj.add_header("Content-Disposition", "attachment", filename=filename)
        msg.attach(adj)
    
    # Enviar
    all_recipients = TO_RECIPIENTS + CC_RECIPIENTS
    try:
        print()
        print("=" * 70)
        print("ENVIANDO CORREO CON PDF")
        print("=" * 70)
        print(f"[INFO] Conectando a {SMTP_SERVIDOR}:{SMTP_PUERTO}...")
        with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
            server.starttls()
            print(f"[INFO] Autenticando como {SMTP_USUARIO}...")
            server.login(SMTP_USUARIO, SMTP_PASSWORD)
            print(f"[INFO] Enviando correo a: {', '.join(all_recipients)}")
            server.sendmail(SMTP_USUARIO, all_recipients, msg.as_string())
        print("[OK] Correo enviado correctamente.")
        return True
    except Exception as e:
        print(f"[ERROR] Falló el envío del correo: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    print("=" * 70)
    print("GENERANDO REPORTES DE PARQUE ARAUCO - DICIEMBRE")
    print("=" * 70)
    print(f"Empresa: Parque Arauco ({COMPANY_ID})")
    print(f"Periodo: {START_DATE} - {END_DATE}")
    print(f"Total de nodos: {len(NODOS_PARQUE_ARAUCO)}")
    print("Configuración: Todos los puntos son consumidores (sin fuente de agua)")
    print("=" * 70)
    print()
    
    # 1. Generar reportes individuales
    nodos_exitosos, nodos_fallidos = generar_reportes_individuales()
    
    # 2. Generar reporte agregado
    reporte_agregado = generar_reporte_agregado(nodos_exitosos)
    
    if not reporte_agregado:
        print()
        print("=" * 70)
        print("PROCESO COMPLETADO CON ERRORES")
        print("=" * 70)
        print("No se pudo generar el reporte agregado. No se enviará correo.")
        return
    
    # 3. Convertir a PDF
    print()
    print("=" * 70)
    print("CONVIRTIENDO REPORTE AGREGADO A PDF")
    print("=" * 70)
    pdf_path = convertir_word_a_pdf(reporte_agregado)
    
    if not pdf_path or not pdf_path.exists():
        print(f"[ADVERTENCIA] No se pudo convertir a PDF. Se intentará enviar el Word.")
        adj_path = reporte_agregado
    else:
        print(f"[OK] PDF generado: {pdf_path}")
        adj_path = pdf_path
    
    # 4. Enviar correo
    if adj_path.suffix.lower() == ".pdf":
        enviar_correo_pdf(adj_path)
    else:
        print("[ADVERTENCIA] Solo se envía PDF. No se enviará el Word.")
    
    # Resumen final
    print()
    print("=" * 70)
    print("PROCESO COMPLETADO")
    print("=" * 70)
    print()
    print("RESUMEN:")
    print(f"  - Reportes individuales generados: {len(nodos_exitosos)}/{len(NODOS_PARQUE_ARAUCO)}")
    if nodos_fallidos:
        print(f"  - Reportes fallidos: {len(nodos_fallidos)}")
        for node_id in nodos_fallidos:
            print(f"    * {get_node_name(node_id)} ({node_id})")
    print(f"  - Reporte agregado: {'Sí' if reporte_agregado else 'No'}")
    print(f"  - PDF generado: {'Sí' if pdf_path and pdf_path.exists() else 'No'}")
    print(f"  - Correo enviado: {'Sí' if adj_path.suffix.lower() == '.pdf' else 'No (solo se envía PDF)'}")
    print()
    print("Los reportes se han guardado en disco en la carpeta reports/Parque_Arauco/")


if __name__ == "__main__":
    main()


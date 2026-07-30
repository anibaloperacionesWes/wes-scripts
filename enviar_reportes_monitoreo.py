"""
Script para enviar los reportes de monitoreo (reporte jose y reporte puntos en cero) por correo.
"""

import sys
import smtplib
from pathlib import Path
from datetime import datetime, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.utils import formatdate, make_msgid

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

# Configuración de correo
SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_PASSWORD = "vxbynfpoehbweelj"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

# Destinatarios
TO_RECIPIENTS = ["anibal@wes.cl"]  # Cambiar por el correo real de Aníbal
CC_RECIPIENTS = [
    "diego@wes.cl",      # Cambiar por el correo real de Diego
    "juan@wes.cl",       # Cambiar por el correo real de Juan
    "benjamin@wes.cl"    # Cambiar por el correo real de Benjamín
]


def buscar_archivos_reporte(carpeta: Path):
    """
    Busca los archivos de reporte en la carpeta.
    """
    reporte_jose = None
    reporte_cero = None
    
    # Buscar reporte de José
    for archivo in carpeta.glob("*jose*.docx"):
        if not archivo.name.startswith("~$"):  # Ignorar archivos temporales
            reporte_jose = archivo
            break
    
    # Buscar reporte de puntos en cero más reciente
    reportes_cero = sorted(
        [f for f in carpeta.glob("*Puntos_En_Cero*.docx") if not f.name.startswith("~$")],
        key=lambda x: x.stat().st_mtime,
        reverse=True
    )
    
    if reportes_cero:
        reporte_cero = reportes_cero[0]
    
    return reporte_jose, reporte_cero


def enviar_correo_reporte(reporte_jose: Path, reporte_cero: Path):
    """
    Envía un correo con los dos reportes adjuntos.
    """
    if not reporte_jose or not reporte_jose.exists():
        print(f"[ERROR] No se encontró el archivo de reporte de José: {reporte_jose}")
        return False
    
    if not reporte_cero or not reporte_cero.exists():
        print(f"[ERROR] No se encontró el archivo de reporte de puntos en cero: {reporte_cero}")
        return False
    
    # Construir correo con headers mejorados para evitar spam
    msg = MIMEMultipart()
    
    # Headers básicos
    msg["From"] = f"Agente WES <{SMTP_USUARIO}>"
    msg["To"] = ", ".join(TO_RECIPIENTS)
    msg["Cc"] = ", ".join(CC_RECIPIENTS)
    msg["Reply-To"] = SMTP_USUARIO
    msg["Subject"] = "Reportes de Monitoreo WES - Análisis del Sistema"
    
    # Headers importantes para evitar spam
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain="wes.cl")
    msg["X-Mailer"] = "WES Report Generator"
    msg["X-Priority"] = "3"
    msg["MIME-Version"] = "1.0"
    
    fecha_generacion = datetime.now().strftime("%d de %B de %Y")
    
    # Cuerpo en texto plano
    cuerpo_texto = f"""Estimado Aníbal,

Te envío los reportes de monitoreo del sistema WES generados mediante acceso vía API a toda la información de la aplicación WES.

REPORTES ADJUNTOS:

1. Reporte de Monitoreo General: Contiene información adicional del sistema de monitoreo, incluyendo gráficos, métricas y análisis complementarios.

2. Reporte de Puntos en Cero y Sin Datos: Análisis detallado de los puntos de monitoreo que están registrando consumo cero o no tienen datos disponibles en los últimos días.

METODOLOGÍA:

Los datos fueron obtenidos mediante consultas automatizadas a la API de WES, accediendo a:
- Información de todos los puntos de monitoreo del sistema
- Datos de consumo horario y diario
- Estado de conectividad y disponibilidad de datos
- Análisis de puntos con consumo cero o sin datos disponibles

El proceso incluyó la verificación de {len(list(Path("reporte en cero").glob("*Puntos_En_Cero*.docx")))} puntos de monitoreo distribuidos en todo el sistema, analizando su estado durante los últimos 3 días.

Este correo y los reportes fueron generados automáticamente por un agente de IA al servicio de WES, configurado específicamente para apoyar la generación y distribución de reportes de monitoreo.

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
        .header {{ background-color: #f4f4f4; padding: 15px; border-radius: 5px; margin-bottom: 20px; }}
        .content {{ margin: 20px 0; }}
        .section {{ margin: 15px 0; padding: 10px; background-color: #f9f9f9; border-left: 4px solid #4472C4; }}
        .footer {{ margin-top: 30px; padding-top: 20px; border-top: 1px solid #ddd; font-size: 0.9em; color: #666; }}
        ul {{ margin: 10px 0; padding-left: 20px; }}
        li {{ margin: 5px 0; }}
        h3 {{ color: #4472C4; }}
    </style>
</head>
<body>
    <div class="header">
        <h2>Reportes de Monitoreo WES</h2>
        <p><strong>Fecha de generación:</strong> {fecha_generacion}</p>
    </div>
    
    <div class="content">
        <p>Estimado Aníbal,</p>
        
        <p>Te envío los reportes de monitoreo del sistema WES generados mediante acceso vía API a toda la información de la aplicación WES.</p>
        
        <div class="section">
            <h3>Reportes Adjuntos:</h3>
            <ol>
                <li><strong>Reporte de Monitoreo General:</strong> Contiene información adicional del sistema de monitoreo, incluyendo gráficos, métricas y análisis complementarios.</li>
                <li><strong>Reporte de Puntos en Cero y Sin Datos:</strong> Análisis detallado de los puntos de monitoreo que están registrando consumo cero o no tienen datos disponibles en los últimos días.</li>
            </ol>
        </div>
        
        <div class="section">
            <h3>Metodología:</h3>
            <p>Los datos fueron obtenidos mediante consultas automatizadas a la API de WES, accediendo a:</p>
            <ul>
                <li>Información de todos los puntos de monitoreo del sistema</li>
                <li>Datos de consumo horario y diario</li>
                <li>Estado de conectividad y disponibilidad de datos</li>
                <li>Análisis de puntos con consumo cero o sin datos disponibles</li>
            </ul>
            <p>El proceso incluyó la verificación de múltiples puntos de monitoreo distribuidos en todo el sistema, analizando su estado durante los últimos 3 días.</p>
        </div>
        
        <p>Este correo y los reportes fueron generados automáticamente por un <strong>agente de IA al servicio de WES</strong>, configurado específicamente para apoyar la generación y distribución de reportes de monitoreo.</p>
        
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
    
    # Adjuntar archivos
    archivos_adjuntos = [
        (reporte_jose, "Reporte_Monitoreo_General.docx"),
        (reporte_cero, "Reporte_Puntos_En_Cero.docx")
    ]
    
    for archivo_path, nombre_adjunto in archivos_adjuntos:
        try:
            with open(archivo_path, "rb") as f:
                adj = MIMEApplication(f.read(), _subtype="docx")
                adj.add_header("Content-Disposition", "attachment", filename=nombre_adjunto)
                msg.attach(adj)
            print(f"[OK] Archivo adjunto: {nombre_adjunto}")
        except Exception as e:
            print(f"[ERROR] No se pudo adjuntar {archivo_path}: {e}")
            return False
    
    # Enviar
    all_recipients = TO_RECIPIENTS + CC_RECIPIENTS
    try:
        print()
        print("=" * 70)
        print("ENVIANDO CORREO CON REPORTES")
        print("=" * 70)
        print(f"[INFO] Conectando a {SMTP_SERVIDOR}:{SMTP_PUERTO}...")
        with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
            server.starttls()
            print(f"[INFO] Autenticando como {SMTP_USUARIO}...")
            server.login(SMTP_USUARIO, SMTP_PASSWORD)
            print(f"[INFO] Enviando correo a: {', '.join(TO_RECIPIENTS)}")
            print(f"[INFO] Con copia a: {', '.join(CC_RECIPIENTS)}")
            server.sendmail(SMTP_USUARIO, all_recipients, msg.as_string())
        print("[OK] Correo enviado correctamente.")
        return True
    except Exception as e:
        print(f"[ERROR] Falló el envío del correo: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """Función principal."""
    print("=" * 70)
    print("ENVÍO DE REPORTES DE MONITOREO")
    print("=" * 70)
    print()
    
    carpeta = Path("reporte en cero")
    
    if not carpeta.exists():
        print(f"[ERROR] La carpeta '{carpeta}' no existe.")
        return
    
    # Buscar archivos
    print("Buscando archivos de reporte...")
    reporte_jose, reporte_cero = buscar_archivos_reporte(carpeta)
    
    if not reporte_jose:
        print("[ERROR] No se encontró el archivo 'reporte jose' en la carpeta.")
        return
    
    if not reporte_cero:
        print("[ERROR] No se encontró el archivo de reporte de puntos en cero.")
        return
    
    print(f"[OK] Archivo encontrado: {reporte_jose.name}")
    print(f"[OK] Archivo encontrado: {reporte_cero.name}")
    print()
    
    # Enviar correo
    if enviar_correo_reporte(reporte_jose, reporte_cero):
        print()
        print("=" * 70)
        print("PROCESO COMPLETADO")
        print("=" * 70)
        print("Correo enviado exitosamente con los reportes adjuntos.")
    else:
        print()
        print("=" * 70)
        print("PROCESO FALLIDO")
        print("=" * 70)
        print("No se pudo enviar el correo.")


if __name__ == "__main__":
    main()









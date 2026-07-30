"""
Script para enviar los reportes de Fundo Zapallar por correo.
Busca los archivos más recientes y los envía.
"""

import sys
from pathlib import Path
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

# Configuración SMTP
SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_PASSWORD = "vxbynfpoehbweelj"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

# Destinatarios
DESTINATARIOS = [
    "joseotarola@wes.cl",      # José
    "diegocarrasco@wes.cl",    # Diego
    "anibal.aoperaciones@wes.cl"  # Aníbal
]

START_DATE = "01/01/2026"
END_DATE = "04/02/2026"


def encontrar_archivos_recientes():
    """Encuentra los archivos más recientes generados."""
    base_dir = Path("reports") / "Fundo_Zapallar"
    
    archivos = {
        "reporte_agregado": None,
        "ppt": None,
        "pdf": None,
        "reportes_individuales": []
    }
    
    # Buscar reporte agregado más reciente
    agregado_dir = base_dir / "ABREGADO"
    if agregado_dir.exists():
        docx_files = list(agregado_dir.rglob("*.docx"))
        if docx_files:
            archivos["reporte_agregado"] = max(docx_files, key=lambda p: p.stat().st_mtime)
        
        pptx_files = list(agregado_dir.rglob("*.pptx"))
        if pptx_files:
            archivos["ppt"] = max(pptx_files, key=lambda p: p.stat().st_mtime)
        
        pdf_files = list(agregado_dir.rglob("*.pdf"))
        if pdf_files:
            archivos["pdf"] = max(pdf_files, key=lambda p: p.stat().st_mtime)
    
    # Buscar reportes individuales (solo DOCX, no PDFs por ahora)
    reporte_dir = base_dir / "REPORTE"
    if reporte_dir.exists():
        docx_files = list(reporte_dir.rglob("Reporte_*.docx"))
        if docx_files:
            # Ordenar por fecha de modificación y tomar los más recientes
            archivos["reportes_individuales"] = sorted(
                docx_files, 
                key=lambda p: p.stat().st_mtime, 
                reverse=True
            )
    
    return archivos


def enviar_correo(archivos):
    """Envía el correo con todos los archivos adjuntos."""
    try:
        # Crear mensaje
        msg = MIMEMultipart()
        msg['From'] = SMTP_USUARIO
        msg['To'] = ", ".join(DESTINATARIOS)
        msg['Subject'] = f"Reportes Fundo Zapallar - {START_DATE} al {END_DATE}"
        
        # Cuerpo del correo
        cuerpo = f"""
Estimados José, Diego y Aníbal,

Se adjuntan los reportes completos de Fundo Zapallar para el período {START_DATE} al {END_DATE}:

- Reportes individuales de todos los puntos de monitoreo
- Reporte agregado consolidado
- Presentación en PowerPoint (PPT)
- Presentación en PDF (si está disponible)

Los reportes fueron generados automáticamente desde el sistema WES.

Saludos cordiales,
Sistema WES
"""
        msg.attach(MIMEText(cuerpo, 'plain', 'utf-8'))
        
        archivos_adjuntados = 0
        
        # Adjuntar reporte agregado
        if archivos["reporte_agregado"] and archivos["reporte_agregado"].exists():
            print(f"[INFO] Adjuntando reporte agregado: {archivos['reporte_agregado'].name}")
            try:
                with open(archivos["reporte_agregado"], 'rb') as f:
                    adjunto = MIMEApplication(f.read())
                    adjunto.add_header(
                        'Content-Disposition',
                        'attachment',
                        filename=archivos["reporte_agregado"].name
                    )
                    msg.attach(adjunto)
                archivos_adjuntados += 1
            except Exception as e:
                print(f"  [ERROR] No se pudo adjuntar reporte agregado: {e}")
        
        # Adjuntar presentación PPT
        if archivos["ppt"] and archivos["ppt"].exists():
            print(f"[INFO] Adjuntando presentación PPT: {archivos['ppt'].name}")
            try:
                with open(archivos["ppt"], 'rb') as f:
                    adjunto = MIMEApplication(f.read())
                    adjunto.add_header(
                        'Content-Disposition',
                        'attachment',
                        filename=archivos["ppt"].name
                    )
                    msg.attach(adjunto)
                archivos_adjuntados += 1
            except Exception as e:
                print(f"  [ERROR] No se pudo adjuntar PPT: {e}")
        
        # Adjuntar presentación PDF
        if archivos["pdf"] and archivos["pdf"].exists():
            print(f"[INFO] Adjuntando presentación PDF: {archivos['pdf'].name}")
            try:
                with open(archivos["pdf"], 'rb') as f:
                    adjunto = MIMEApplication(f.read())
                    adjunto.add_header(
                        'Content-Disposition',
                        'attachment',
                        filename=archivos["pdf"].name
                    )
                    msg.attach(adjunto)
                archivos_adjuntados += 1
            except Exception as e:
                print(f"  [ERROR] No se pudo adjuntar PDF: {e}")
        
        # Adjuntar reportes individuales (máximo 10 para no sobrecargar el correo)
        print(f"[INFO] Adjuntando reportes individuales...")
        reportes_adjuntados = 0
        for report_path in archivos["reportes_individuales"][:10]:  # Máximo 10 reportes
            if report_path.exists():
                try:
                    with open(report_path, 'rb') as f:
                        adjunto = MIMEApplication(f.read())
                        adjunto.add_header(
                            'Content-Disposition',
                            'attachment',
                            filename=report_path.name
                        )
                        msg.attach(adjunto)
                    reportes_adjuntados += 1
                except Exception as e:
                    print(f"  [ADVERTENCIA] No se pudo adjuntar {report_path.name}: {e}")
        
        print(f"  [OK] {reportes_adjuntados} reporte(s) individual(es) adjuntado(s)")
        archivos_adjuntados += reportes_adjuntados
        
        if archivos_adjuntados == 0:
            print("[ERROR] No se encontraron archivos para adjuntar")
            return False
        
        # Enviar correo
        print(f"[INFO] Enviando correo a {len(DESTINATARIOS)} destinatario(s)...")
        with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
            server.starttls()
            server.login(SMTP_USUARIO, SMTP_PASSWORD)
            server.send_message(msg)
        
        print(f"[OK] Correo enviado exitosamente a: {', '.join(DESTINATARIOS)}")
        print(f"[OK] Total de archivos adjuntados: {archivos_adjuntados}")
        return True
        
    except Exception as e:
        print(f"[ERROR] Error al enviar correo: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    print("=" * 70)
    print("  ENVÍO DE REPORTES FUNDO ZAPALLAR POR CORREO")
    print("=" * 70)
    print()
    
    print("[1/2] Buscando archivos generados...")
    archivos = encontrar_archivos_recientes()
    
    print(f"[OK] Archivos encontrados:")
    if archivos["reporte_agregado"]:
        print(f"  - Reporte agregado: {archivos['reporte_agregado']}")
    if archivos["ppt"]:
        print(f"  - Presentación PPT: {archivos['ppt']}")
    if archivos["pdf"]:
        print(f"  - Presentación PDF: {archivos['pdf']}")
    print(f"  - Reportes individuales: {len(archivos['reportes_individuales'])}")
    print()
    
    print("[2/2] Enviando correo...")
    exito = enviar_correo(archivos)
    
    print()
    if exito:
        print("=" * 70)
        print("  CORREO ENVIADO EXITOSAMENTE")
        print("=" * 70)
    else:
        print("=" * 70)
        print("  ERROR AL ENVIAR CORREO")
        print("=" * 70)


if __name__ == "__main__":
    main()

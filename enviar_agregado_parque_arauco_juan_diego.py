"""Enviar el reporte agregado de Parque Arauco Kennedy a Juan y Diego, con copia al usuario.

El correo explica brevemente cómo se obtuvieron los datos sin exponer accesos a las APIs
y menciona que fue generado por un agente de IA al servicio de WES.
"""

from pathlib import Path
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from generar_reporte_word import get_company_name, convertir_word_a_pdf


SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_PASSWORD = "vxbynfpoehbweelj"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

TO_RECIPIENTS = ["juanlopez@wes.cl", "diegocarrasco@wes.cl"]
CC_RECIPIENTS = ["agente.ia@wes.cl"]

COMPANY_ID = "000025"


def main() -> None:
    base_dir = Path("reports") / "Parque_Arauco" / "ABREGADO"

    if not base_dir.exists():
        print("[ERROR] No se encontró la carpeta de reportes agregados de Parque Arauco.")
        return

    candidatos = sorted(base_dir.rglob("Reporte_Agregado_*.docx"))
    if not candidatos:
        print("[ERROR] No se encontró ningún Reporte_Agregado_*.docx.")
        return

    reporte_path = candidatos[-1]
    company_name = get_company_name(COMPANY_ID)

    # Convertir a PDF (si es posible)
    pdf_path = convertir_word_a_pdf(reporte_path)
    adj_path: Path = pdf_path if pdf_path and pdf_path.exists() else reporte_path
    es_pdf = adj_path.suffix.lower() == ".pdf"

    print(f"Enviando reporte agregado desde: {adj_path}")

    # Construir correo
    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(TO_RECIPIENTS)
    msg["Cc"] = ", ".join(CC_RECIPIENTS)
    msg["Subject"] = (
        f"Reporte agregado Parque Arauco Kennedy - {company_name} - 07-12-25 a 14-12-25"
    )

    cuerpo = f"""Estimados Juan y Diego,

Adjunto encontrarán el reporte agregado de consumo de agua y alertas de Parque Arauco Kennedy
para el periodo 07-12-2025 al 14-12-2025.

Los datos se obtuvieron desde el sistema WES, utilizando los servicios internos que consultan
las lecturas históricas de los medidores (consumos horarios y diarios) y las alertas configuradas
para cada punto. El agente no tiene acceso directo a credenciales ni a las cuentas de usuario:
solo ejecuta scripts autorizados dentro del entorno WES para leer la información necesaria y
procesarla en este formato de reporte.

El documento consolida:
- Consumos diarios y consumo efectivo por punto.
- Alertas de filtración detectadas y su proyección cuando corresponde.
- Gráficas comparativas y un análisis agregado de todos los puntos de Parque Arauco Kennedy.

Este correo y el reporte fueron generados automáticamente por un agente de IA al servicio de WES,
configurado específicamente para apoyar la generación y distribución de reportes a clientes internos.

Quedo atento a cualquier comentario.

Saludos,
Agente IA WES
"""

    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

    # Adjuntar archivo
    with open(adj_path, "rb") as f:
        subtype = "pdf" if es_pdf else "docx"
        adj = MIMEApplication(f.read(), _subtype=subtype)
        filename = adj_path.name
        adj.add_header("Content-Disposition", "attachment", filename=filename)
        msg.attach(adj)

    # Enviar
    all_recipients = TO_RECIPIENTS + CC_RECIPIENTS
    try:
        print(f"[INFO] Conectando a {SMTP_SERVIDOR}:{SMTP_PUERTO}...")
        with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
            server.starttls()
            print(f"[INFO] Autenticando como {SMTP_USUARIO}...")
            server.login(SMTP_USUARIO, SMTP_PASSWORD)
            print(f"[INFO] Enviando correo a: {', '.join(all_recipients)}")
            server.sendmail(SMTP_USUARIO, all_recipients, msg.as_string())
        print("[OK] Correo enviado correctamente a Juan, Diego y con copia a ti.")
    except Exception as e:
        print(f"[ERROR] Falló el envío del correo: {e}")


if __name__ == "__main__":
    main()















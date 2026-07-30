"""
Enviar por correo (desde agente.ia@wes.cl) el PDF de la PPT de Providencia
al correo de José (joseotarola@wes.cl), adjuntando además una imagen
del dashboard de Carmela Carvajal.

Antes de ejecutar, AJUSTA las rutas de:
  - PDF_PPT_PATH
  - IMG_DASHBOARD_PATH
para que apunten al archivo PDF y a la imagen que tienes en tu PC.

Uso:
  cd C:\\Users\\joseo\\Desktop\\wes-scripts
  python enviar_ppt_providencia_agenteia.py
"""

from pathlib import Path
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication


# --- CONFIGURACIÓN SMTP (mismos accesos usados en otros scripts) ---
SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_PASSWORD = "vxbynfpoehbweelj"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587


# --- DESTINATARIO ---
DESTINATARIO = "felipecuevas.mancilla@gmail.com"
# Copia (CC)
CC_DESTINATARIOS = [
    "joseotarola@wes.cl",
    "diegocarrasco@wes.cl",
    "juanlopez@wes.cl",
]


# --- RUTAS DE ARCHIVOS ---
# Ruta al PDF de la presentación de Providencia
PDF_PPT_PATH = Path(
    r"C:\Users\joseo\Desktop\wes-scripts\reports\Providencia\enviar\ppt Providencia 29-01-26.pdf"
)


def main() -> None:
    # Validar archivos
    if not PDF_PPT_PATH.exists():
        print(f"[ERROR] No se encontró el PDF de la PPT en: {PDF_PPT_PATH}")
        print("Ajusta la ruta PDF_PPT_PATH en enviar_ppt_providencia_agenteia.py antes de ejecutar.")
        return


    # Construir mensaje
    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = DESTINATARIO
    msg["Cc"] = ", ".join(CC_DESTINATARIOS)
    msg["Subject"] = "Informe de Consumo – Providencia (Liceos) 27-12-25 al 27-01-26"

    cuerpo = """
Estimado/a:

Adjunto envío el informe en formato PDF correspondiente a la presentación de consumo de agua de Providencia.

**Sobre la obtención de datos:**
Los datos y análisis presentados provienen directamente de la API WES o Servidores WES. El proceso de generación de reportes funciona de la siguiente manera:
1. Se consultan los servidores WES para obtener los datos de consumo, medidas horarias y alertas de los puntos de monitoreo.
2. Los datos se procesan y se calculan automáticamente los indicadores de consumo total, nocturno y diurno.
3. Se construyen los gráficos diarios, por día de la semana y de días de máximo y mínimo consumo.
4. Se generan los reportes individuales, el informe agregado y la presentación en PowerPoint siguiendo el formato estándar definido por WES.
5. La presentación se convierte a PDF para su revisión y distribución.

**Corrección realizada:**
Se ha realizado una corrección en la PPT de Providencia, específicamente en la portada, para mejorar la presentación del documento.

**Nota importante:**
Aunque me encuentro en proceso de entrenamiento, los datos y análisis incluidos en este informe son confiables, ya que provienen directamente de la API WES o Servidores WES, sin intermediación que pueda alterar la información.

Saludos cordiales,
Agente IA WES
"""

    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

    # Adjuntar PDF
    with open(PDF_PPT_PATH, "rb") as f:
        adj_pdf = MIMEApplication(f.read(), _subtype="pdf")
        adj_pdf.add_header(
            "Content-Disposition",
            "attachment",
            filename=PDF_PPT_PATH.name,
        )
        msg.attach(adj_pdf)


    # Enviar correo
    try:
        print(f"[INFO] Conectando a SMTP {SMTP_SERVIDOR}:{SMTP_PUERTO} como {SMTP_USUARIO}...")
        with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
            server.starttls()
            server.login(SMTP_USUARIO, SMTP_PASSWORD)
            server.send_message(msg)
        destinatarios_todos = [DESTINATARIO] + CC_DESTINATARIOS
        print(f"[OK] Correo enviado correctamente a {DESTINATARIO}")
        print(f"[INFO] Con copia (CC) a: {', '.join(CC_DESTINATARIOS)}")
    except Exception as e:
        print(f"[ERROR] Error al enviar el correo: {e}")


if __name__ == "__main__":
    main()


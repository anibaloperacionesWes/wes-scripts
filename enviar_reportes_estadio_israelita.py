"""
Enviar por correo (desde agente.ia@wes.cl) los reportes PDF de Estadio Israelita Maccabi
a fcuevas@eragroup.com con copia a juan, diego y jose.

Uso:
  cd C:\\Users\\joseo\\Desktop\\wes-scripts
  python enviar_reportes_estadio_israelita.py
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


# --- DESTINATARIOS ---
DESTINATARIO = "fcuevas@eragroup.com"
# Copia (CC)
CC_DESTINATARIOS = [
    "juanlopez@wes.cl",
    "diegocarrasco@wes.cl",
    "joseotarola@wes.cl",
]


# --- RUTAS DE ARCHIVOS PDF ---
PDF_REPORTES = [
    Path(r"C:\Users\joseo\Desktop\wes-scripts\reports\Estadio_Israelita_Maccabi\REPORTE\enviar\Matriz principal Av. Las condes. 29-01-26.pdf"),
    Path(r"C:\Users\joseo\Desktop\wes-scripts\reports\Estadio_Israelita_Maccabi\REPORTE\enviar\Matriz Chesterton 29-01-26.pdf"),
]


def main() -> None:
    # Validar archivos
    pdfs_validos = []
    for pdf_path in PDF_REPORTES:
        if pdf_path.exists():
            pdfs_validos.append(pdf_path)
            print(f"[OK] PDF encontrado: {pdf_path.name}")
        else:
            print(f"[ADVERTENCIA] No se encontró: {pdf_path}")
    
    if not pdfs_validos:
        print("[ERROR] No se encontró ningún PDF para adjuntar.")
        return
    
    # Construir mensaje
    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = DESTINATARIO
    msg["Cc"] = ", ".join(CC_DESTINATARIOS)
    msg["Subject"] = "Reportes de Consumo - Estadio Israelita Maccabi"

    cuerpo = """
Estimado/a:

Adjunto envío los reportes de consumo de agua correspondientes a Estadio Israelita Maccabi.

**Archivos adjuntos:**
- Matriz principal Av. Las condes. 29-01-26.pdf
- Matriz Chesterton 29-01-26.pdf

**Sobre la obtención de datos:**
Los datos y análisis presentados en estos reportes provienen directamente de la API WES o Servidores WES. El proceso de generación de reportes funciona de la siguiente manera:
1. Se consultan los servidores WES para obtener los datos de consumo, medidas horarias y alertas de los puntos de monitoreo.
2. Los datos se procesan y se calculan automáticamente los indicadores de consumo total, nocturno y diurno.
3. Se construyen los gráficos diarios, por día de la semana y de días de máximo y mínimo consumo.
4. Se generan los reportes individuales siguiendo el formato estándar definido por WES.
5. Los reportes se convierten a PDF para su revisión y distribución.

**Nota importante:**
Aunque me encuentro en proceso de entrenamiento, los datos y análisis incluidos en estos reportes son confiables, ya que provienen directamente de la API WES o Servidores WES, sin intermediación que pueda alterar la información.

Saludos cordiales,
Agente IA WES
"""

    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

    # Adjuntar todos los PDFs
    for pdf_path in pdfs_validos:
        with open(pdf_path, "rb") as f:
            adj_pdf = MIMEApplication(f.read(), _subtype="pdf")
            adj_pdf.add_header(
                "Content-Disposition",
                "attachment",
                filename=pdf_path.name,
            )
            msg.attach(adj_pdf)
        print(f"[OK] PDF adjuntado: {pdf_path.name}")

    # Enviar correo
    try:
        print(f"\n[INFO] Conectando a SMTP {SMTP_SERVIDOR}:{SMTP_PUERTO} como {SMTP_USUARIO}...")
        with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
            server.starttls()
            server.login(SMTP_USUARIO, SMTP_PASSWORD)
            server.send_message(msg)
        print(f"[OK] Correo enviado correctamente a {DESTINATARIO}")
        print(f"[INFO] Con copia (CC) a: {', '.join(CC_DESTINATARIOS)}")
    except Exception as e:
        print(f"[ERROR] Error al enviar el correo: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()

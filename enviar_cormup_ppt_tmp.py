import smtplib
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_PASSWORD = "vxbynfpoehbweelj"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

DESTINATARIOS = [
    "anibal.aoperaciones@wes.cl",
    "joseotarola@wes.cl",
]

PPT_PATH = Path(r"C:\Users\joseo\Desktop\wes-scripts\reports\CORMUP\ABREGADO\AGREGADO_20260303_1706\Agregado PPT.pptx")
if not PPT_PATH.exists():
    raise SystemExit(f"No se encontro la PPT: {PPT_PATH}")

msg = MIMEMultipart()
msg["From"] = SMTP_USUARIO
msg["To"] = ", ".join(DESTINATARIOS)
msg["Subject"] = "Presentación PPT CORMUP - 01/02/2026 al 28/02/2026"

cuerpo = (
    "Equipo,\n\n"
    "Adjunto la presentación PPT agregada de CORMUP/Peñalolén para el periodo 01/02/2026 al 28/02/2026.\n\n"
    "Saludos,\n"
    "Sistema WES\n"
)
msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

with open(PPT_PATH, "rb") as f:
    adjunto = MIMEApplication(f.read(), _subtype="vnd.openxmlformats-officedocument.presentationml.presentation")
    adjunto.add_header("Content-Disposition", "attachment", filename=PPT_PATH.name)
    msg.attach(adjunto)

with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
    server.starttls()
    server.login(SMTP_USUARIO, SMTP_PASSWORD)
    server.send_message(msg, to_addrs=DESTINATARIOS)

print("OK")

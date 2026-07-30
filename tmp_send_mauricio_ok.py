import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from pathlib import Path

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_PASSWORD = "vxbynfpoehbweelj"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

to_email = "mauricioorellana@wes.cl"
pdf_path = Path(r"c:\Users\aniba\OneDrive\Desktop\alerta control nocturno.pdf")

if not pdf_path.exists():
    raise SystemExit(f"No existe adjunto: {pdf_path}")

msg = MIMEMultipart()
msg["From"] = SMTP_USUARIO
msg["To"] = to_email
msg["Subject"] = "Alerta Control Nocturno (cortes programados) - 23-03-2026"
msg.attach(MIMEText("Se adjunta reporte de control nocturno (madrugada 23-03-2026).", "plain", "utf-8"))

with pdf_path.open("rb") as f:
    part = MIMEApplication(f.read(), _subtype="pdf")
    part.add_header("Content-Disposition", "attachment", filename=pdf_path.name)
    msg.attach(part)

with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
    server.starttls()
    server.login(SMTP_USUARIO, SMTP_PASSWORD)
    server.send_message(msg, to_addrs=[to_email])

print(f"OK enviado a {to_email}")

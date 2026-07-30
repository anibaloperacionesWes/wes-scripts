import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from pathlib import Path

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_PASSWORD = "vxbynfpoehbweelj"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

to_addrs = [
    "diegocarrasco@wes.cl",
    "juanlopez@wes.cl",
    "joseotarola@wes.cl",
    "maurcioorellana@wes.cl",
]

adjuntos = [
    Path(r"c:\Users\aniba\OneDrive\Desktop\alerta control nocturno.pdf"),
    Path(r"c:\Users\aniba\OneDrive\Desktop\Reporte Puntos en Cero.pdf"),
]

faltantes = [str(p) for p in adjuntos if not p.exists()]
if faltantes:
    raise SystemExit("Faltan adjuntos: " + " | ".join(faltantes))

msg = MIMEMultipart()
msg["From"] = SMTP_USUARIO
msg["To"] = ", ".join(to_addrs)
msg["Subject"] = "Reportes WES: Alerta Control Nocturno + Puntos en Cero"

cuerpo = """Estimados,

Se adjuntan los reportes solicitados:
- Alerta Control Nocturno
- Reporte Puntos en Cero

Saludos,
Sistema WES
"""
msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

for p in adjuntos:
    with p.open("rb") as f:
        part = MIMEApplication(f.read(), _subtype="pdf")
        part.add_header("Content-Disposition", "attachment", filename=p.name)
        msg.attach(part)

with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
    server.starttls()
    server.login(SMTP_USUARIO, SMTP_PASSWORD)
    server.send_message(msg, to_addrs=to_addrs)

print("OK enviado a:")
for e in to_addrs:
    print(e)

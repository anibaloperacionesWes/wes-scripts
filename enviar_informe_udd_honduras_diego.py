"""
Envía a Diego (y equipo) el Informe Ejecutivo Nº2 UDD — Red Impulsión Honduras.

Generar antes:
  python generar_reporte_udd_sectorizacion_honduras.py

Adjunta Word + PDF desde:
  reports/udd_sectorizacion_honduras/entrega_diego/
"""

from __future__ import annotations

import os
import smtplib
import sys
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

ROOT = Path(__file__).resolve().parent
ENTREGA = ROOT / "reports" / "udd_sectorizacion_honduras" / "entrega_diego"
DOCX_PATH = ENTREGA / "Informe_Ejecutivo_Pruebas_Hidricas_UDD_Honduras_N2.docx"
PDF_PATH = ENTREGA / "Informe_Ejecutivo_Pruebas_Hidricas_UDD_Honduras_N2.pdf"

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

TO_RECIPIENTS = ["diegocarrasco@wes.cl"]
CC_RECIPIENTS = ["juanlopez@wes.cl"]


def _smtp_password() -> str:
    p = (
        os.environ.get("WES_GMAIL_APP_PASSWORD", "").strip()
        or os.environ.get("WES_SMTP_PASSWORD", "").strip()
    )
    if p:
        return p.replace(" ", "").strip()
    f = ROOT / "gmail_oauth" / "app_password.txt"
    if f.is_file():
        for line in f.read_text(encoding="utf-8", errors="replace").splitlines():
            line = line.strip()
            if line and not line.startswith("#"):
                return line.replace(" ", "").strip()
    return ""


def main() -> int:
    pw = _smtp_password()
    if not pw:
        print("[ERROR] Falta contraseña SMTP.", file=sys.stderr)
        return 1
    if not DOCX_PATH.is_file():
        print(f"[ERROR] No existe: {DOCX_PATH}", file=sys.stderr)
        print("Ejecute: python generar_reporte_udd_sectorizacion_honduras.py")
        return 1
    if not PDF_PATH.is_file():
        print(f"[ERROR] No existe: {PDF_PATH}", file=sys.stderr)
        return 1

    cuerpo = """Estimado Diego,

Juan en copia.

Adjunto el Informe Ejecutivo Nº2 de pruebas hídricas — Red Impulsión Honduras (UDD), en PDF y Word. Fue generado con el agente de reportes WES a partir de tu informe de referencia del 01.06.2026 y los datos horarios de la API (punto 000026-01, Sala impulsión Honduras).

Resumen de lo realizado
-----------------------
• Estructura alineada a tu Informe Nº2: secciones I. Antecedentes, II. Metodología, III. Análisis y IV. Conclusión, con pie de página WES.
• Misma ventana de maniobra: cierre 30/05/2026 20:00 — reapertura 01/06/2026 07:00 (35 horas).
• Referencia histórica: promedio de la misma ventana en los 3 fines de semana anteriores (23–25/05, 16–18/05 y 09–11/05), no días hábiles de semana.
• Tabla de indicadores con los mismos valores que tu informe (consumo 7,37 vs 76,81 m³; reducción ~90,4 %; caudales medio/máx/mín; hito madrugada domingo 31/05: 0,75 vs 15,11 m³).
• Cuatro perfiles horarios en el documento, en el mismo orden que tu versión:
  - Sábado 30/05 (24 h) y Lunes 01/06 (0–9 h) al cierre de Metodología.
  - Domingo 24/05 (referencia) y Domingo 31/05 (maniobra) al cierre del Análisis.
• Paginación: texto y gráficos en la misma página (sin hojas solo con imágenes), dos gráficos por fila.

Por qué gráficos generados y no capturas de la app
--------------------------------------------------
• Automatización: el flujo puede correr sin intervención manual ni pantallazos por cada informe.
• Misma fuente de datos: usan la serie horaria de la API WES (equivalente a lo que muestra la app), hora Chile.
• Reproducibilidad: mismos días, mismos títulos y escala en todos los clientes que usemos esta plantilla.
• Calidad en Word/PDF: resolución uniforme al insertar en el documento.

Los gráficos imitan el estilo de la app (barras 24 h, madrugada 00–06 en rojo/sombreado, azul WES en el resto), pero no son captura de pantalla. Si para ustedes lo corporativo exige screenshot exacto de la app, lo podemos incorporar como opción.

Consultas para seguir mejorando (maquetas corporativas)
-------------------------------------------------------
1. ¿Qué ajustaría usted para que este informe quede idéntico al estándar WES (tipografía, márgenes, redacción, orden de secciones)?
2. ¿Le parecen correctos los colores usados (azul #4A8CB8 / rojo madrugada #c41e1e) o debemos calzar con una paleta oficial?
3. ¿Dispone de los 4 informes de referencia que mencionó por correo (enlace Drive) para entrenar texto y maqueta en todos los reportes agregados?
4. ¿Prefiere capturas de app, gráficos generados, o ambos (app en informe cliente / generados en anexo técnico)?

Quedamos atentos a su feedback para dejar una plantilla única y que los próximos informes salgan todos iguales y corporativos.

Saludos cordiales,
Sistema WES / Agente IA WES
"""

    msg = MIMEMultipart()
    msg["From"] = SMTP_USUARIO
    msg["To"] = ", ".join(TO_RECIPIENTS)
    if CC_RECIPIENTS:
        msg["Cc"] = ", ".join(CC_RECIPIENTS)
    msg["Subject"] = (
        "UDD — Informe Nº2 Pruebas Hídricas — Red Impulsión Honduras (01.06.2026) — revisión maqueta"
    )
    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

    for path, subtype in ((PDF_PATH, "pdf"), (DOCX_PATH, "vnd.openxmlformats-officedocument.wordprocessingml.document")):
        with open(path, "rb") as f:
            part = MIMEApplication(f.read(), _subtype=subtype)
            part.add_header("Content-Disposition", "attachment", filename=path.name)
            msg.attach(part)

    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
        server.starttls()
        server.login(SMTP_USUARIO, pw)
        server.send_message(msg)

    print(f"[OK] Para: {', '.join(TO_RECIPIENTS)}")
    if CC_RECIPIENTS:
        print(f"     Cc:  {', '.join(CC_RECIPIENTS)}")
    print(f"     PDF:  {PDF_PATH}")
    print(f"     DOCX: {DOCX_PATH}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

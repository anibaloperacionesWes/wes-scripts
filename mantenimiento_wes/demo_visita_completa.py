# -*- coding: utf-8 -*-
"""Demo: genera PDF + registra en Excel (sin enviar correo real)."""

from __future__ import annotations

import base64
import io
import shutil
import sys
from datetime import datetime
from pathlib import Path

from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT.parent))

from enviar_acta_cliente_pdf import enviar_acta_pdf_cliente  # noqa: E402
from generar_pdf_acta_visita import generar_pdf_acta  # noqa: E402
from registrar_visita_excel import registrar_visita_en_excel  # noqa: E402


def _firma_demo_png() -> str:
    img = Image.new("RGB", (800, 280), "white")
    draw = ImageDraw.Draw(img)
    draw.line((40, 180, 220, 120, 360, 200, 520, 110, 700, 170), fill=(18, 38, 58), width=4)
    draw.text((40, 30), "Firma demo — Juan Pérez", fill=(31, 78, 121))
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    b64 = base64.b64encode(buf.getvalue()).decode("ascii")
    return f"data:image/png;base64,{b64}"


def main() -> int:
    data = {
        "cliente": "CORMUP",
        "maquina": "TOBALABA",
        "comuna": "Providencia",
        "fecha": datetime.now().strftime("%Y-%m-%d"),
        "hora": "11:15",
        "tecnico": "Anibal Aranda",
        "ot": "OT-DEMO-FORM-001",
        "motivos": ["Mantenimiento"],
        "tecnologias": ["CPA y CIR"],
        "tipo_mtto": "Mtto Preventivo",
        "tipo_falla": "Auditoría",
        "falla_especifica": "Validación Data",
        "solucion": (
            "Visita preventiva. Se valida conectividad App, enlace CIR y "
            "funcionamiento de válvulas CPA. Sin anomalías críticas."
        ),
        "observaciones": "Demo del formulario vistoso con PDF + Excel.",
        "estado_visita": "cerrada",
        "lectura_medidor": "12845.2",
        "recibido_por": "Juan Pérez",
        "cargo": "Encargado mantención",
        "email_cliente": "demo.cliente@wes.cl",
        "email_cc": "anibal.aoperaciones@wes.cl",
        "firma_png": _firma_demo_png(),
        "checklist_cir": [
            {"elemento": "Conectividad App", "estado": "OK", "obs": "Online"},
            {"elemento": "Enlace CIR", "estado": "OK", "obs": ""},
            {"elemento": "Voltaje 12V DC CIR", "estado": "OK", "obs": "12.1 V"},
        ],
        "checklist_cpa": [
            {"elemento": "Func. presión nocturna", "estado": "OK", "obs": ""},
            {"elemento": "Func. medidor", "estado": "OK", "obs": "Pulso OK"},
        ],
        "checklist_sab": [
            {"elemento": "Temporización", "estado": "N/A", "obs": ""},
        ],
    }

    xlsx, row = registrar_visita_en_excel(data)
    pdf = generar_pdf_acta(data)
    reports = ROOT.parent / "reports" / "Mantenimientos" / "formulario_visita"
    reports.mkdir(parents=True, exist_ok=True)
    dest = reports / pdf.name
    shutil.copy2(pdf, dest)

    email = enviar_acta_pdf_cliente(pdf, data, dry_run=True)

    print("OK demo visita")
    print(f"  Excel: {xlsx} fila {row}")
    print(f"  PDF:   {pdf}")
    print(f"  Copia: {dest}")
    print(f"  Mail dry-run → {email}")

    # subir a Drive si hay credenciales
    try:
        from wes_google_drive import credenciales_configuradas, subir_a_mantenimiento_wes

        if credenciales_configuradas():
            cli = str(data.get("cliente") or "SIN_CLIENTE").strip() or "SIN_CLIENTE"
            info = subir_a_mantenimiento_wes(
                dest,
                subcarpeta=f"Tecnicos_WES_Formulario/Actas_visita_PDF/{cli}",
            )
            print(f"  Drive: {info.get('web_view_link')}")
            (reports / "ULTIMO_LINK_DRIVE.txt").write_text(
                info.get("web_view_link", "") + "\n", encoding="utf-8"
            )
        else:
            print("  Drive: secretos no configurados")
    except Exception as exc:  # noqa: BLE001
        print(f"  Drive error: {exc}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

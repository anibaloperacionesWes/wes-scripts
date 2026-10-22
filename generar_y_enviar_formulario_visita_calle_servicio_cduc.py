"""
Formulario oficial de visita técnica WES — CDUC / CALLE DE SERVICO.

Genera el PDF con el mismo formato del formulario web
(ACTA DE VISITA TÉCNICA WES), sube a Drive y envía con acuse de recibo
a los destinatarios del hilo del 22/10/2026.

Uso:
  python3 generar_y_enviar_formulario_visita_calle_servicio_cduc.py
  python3 generar_y_enviar_formulario_visita_calle_servicio_cduc.py --solo-generar
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "mantenimiento_wes"))
sys.path.insert(0, str(ROOT))

from enviar_acta_cliente_pdf import enviar_acta_pdf_cliente  # noqa: E402
from generar_pdf_acta_visita import generar_pdf_acta  # noqa: E402
from wes_google_drive import subir_a_drive  # noqa: E402

FOLIO = 2341  # siguiente al último del Registro de fallas WES (2340)


def payload() -> dict:
    return {
        "cliente": "CDUC",
        "maquina": "CALLE DE SERVICO",
        "comuna": "Las Condes",
        "fecha": "2026-10-22",
        "hora": "15:30",
        "tecnico": "Anibal Aranda",
        "ot": str(FOLIO),
        "folio": FOLIO,
        "estado_visita": "cerrada",
        "motivos": ["Visita física"],
        "tecnologias": ["CIR"],
        "tipo_mtto": "Mtto Correctivo",
        "tipo_falla": "Ultrasonido",
        "falla_especifica": "Equipo Dañado",
        "lectura_medidor": "—",
        "solucion": (
            "Visita presencial por caída abrupta de caudal medido (comunicación OK; "
            "datos horarios llegaban). Hallazgo: sensores ultrasónicos con sulfato en "
            "su interior, producto del agua que cayó/ingresó en el lugar; cables también "
            "afectados por sulfato en su interior. Además, los parámetros del ultrasonido "
            "se encontraban desconfigurados.\n"
            "Acciones: se limpió/retiró el sulfato de los sensores, se reacondicionaron y "
            "reconectaron correctamente los cables (cortando el tramo deteriorado hasta el "
            "sobrante disponible) y se reconfiguraron los parámetros del ultrasonido. "
            "Quedó midiendo de forma normal/correcta."
        ),
        "observaciones": (
            "Se deja en seguimiento para mantener el correcto funcionamiento. "
            "El cable quedó al límite del sobrante: se cortó hasta donde más se podía y "
            "ya no hay margen para otro corte. Si vuelve a fallar, la próxima intervención "
            "será más larga e incluirá cambio de cables (no se puede cortar más). "
            "Acta sin firma presencial; envío por correo para acuse de recibo "
            "(hilo Calle de Servicio 22/10/2026)."
        ),
        "recibido_por": "Felipe Cuevas (sin firma presencial)",
        "cargo": "Contacto CDUC / hilo de correo",
        "email_cliente": "felipecuevas.mancilla@gmail.com",
        "email_cc": (
            "anibal.aoperaciones@wes.cl, diegocarrasco@wes.cl, "
            "broman@cduc.cl, eleal@cduc.cl"
        ),
        "enviar_correo_cliente": True,
        "checklist_cir": [
            {
                "elemento": "Conectividad App",
                "estado": "OK",
                "obs": "Datos horarios OK; no era falla de comunicación",
            },
            {"elemento": "Enlace CIR", "estado": "OK", "obs": ""},
            {
                "elemento": "Estado exterior caja CIR",
                "estado": "Observación",
                "obs": "Ingreso de agua en el lugar",
            },
            {
                "elemento": "Estado interior caja CIR",
                "estado": "Observación",
                "obs": "Sulfato por agua; limpio en visita",
            },
            {
                "elemento": "Equipo ultrasónico",
                "estado": "OK",
                "obs": "Sulfato interior removido; midiendo normal",
            },
            {
                "elemento": "Transductores",
                "estado": "OK",
                "obs": "Limpieza de sulfato realizada",
            },
            {"elemento": "Sensor de pulso", "estado": "N/A", "obs": ""},
        ],
        "checklist_cpa": [],
        "checklist_sab": [],
        "firma_png": "",
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--solo-generar", action="store_true")
    parser.add_argument("--sin-correo", action="store_true")
    args = parser.parse_args()

    data = payload()
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = (
        ROOT
        / "reports"
        / "CDUC"
        / "Visitas_Tecnicas"
        / f"Calle_de_Servicio_formulario_{stamp[:8]}"
    )
    out_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = out_dir / f"Acta_visita_WES_CDUC_CALLE_DE_SERVICO_{stamp}.pdf"
    (out_dir / "payload_formulario.json").write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    path = generar_pdf_acta(data, out_path=pdf_path)
    print(f"[OK] PDF formulario: {path}")

    if args.solo_generar:
        return 0

    info = subir_a_drive(path, subcarpeta="CDUC/Visitas_Tecnicas")
    print(f"[OK] Drive: {info.get('web_view_link')}")

    if args.sin_correo:
        return 0

    result = enviar_acta_pdf_cliente(path, data)
    print(f"[OK] Correo: {result}")
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())

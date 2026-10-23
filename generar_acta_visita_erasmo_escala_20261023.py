"""
Acta oficial de visita técnica WES — CORMUP / ERASMO ESCALA.

Soporte remoto del viernes 23/10/2026, 09:00 (hora Chile).
Folio siguiente al 2341 del Registro de fallas.
"""

from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "mantenimiento_wes"))

from generar_pdf_acta_visita import generar_pdf_acta  # noqa: E402

FOLIO = 2342
CHILE = ZoneInfo("America/Santiago")
STAMP = datetime.now(CHILE).strftime("%Y%m%d_%H%M%S")
NOMBRE = f"folio_{FOLIO}_2026-10-23_CORMUP_ERASMO_ESCALA_{STAMP}.pdf"
OUT = (
    ROOT
    / "reports"
    / "CORMUP"
    / "Visitas_Tecnicas"
    / NOMBRE
)


def payload() -> dict:
    return {
        "cliente": "CORMUP",
        "maquina": "ERASMO ESCALA",
        "comuna": "Peñalolén",
        "fecha": "2026-10-23",
        "hora": "09:00",
        "tecnico": "Anibal Aranda",
        "ot": str(FOLIO),
        "folio": FOLIO,
        "estado_visita": "cerrada",
        "motivos": ["Soporte técnico a distancia"],
        "tecnologias": ["CPA y CIR"],
        "tipo_mtto": "Soporte remoto",
        "tipo_falla": "Soporte remoto",
        "falla_especifica": "Programación de auditoría",
        "lectura_medidor": "—",
        "solucion": (
            "Soporte técnico a distancia tras el llamado telefónico del viernes "
            "23/10/2026 a las 09:00. Se revisó la programación y había un error "
            "en el día viernes: el corte estaba programado en la mañana. "
            "Se normalizó esa programación.\n"
            "Se conversó por teléfono con la directora Elizabeth González para "
            "confirmar la reparación.\n"
            "El equipo había activado el protocolo de emergencia bypass. "
            "Se normalizó el equipo, que quedó fuera de bypass y en operación regular."
        ),
        "observaciones": (
            "Sin desplazamiento a terreno. Punto WES 000008-07. "
            "Horario de corte de referencia del control nocturno: 01:00 a 05:59. "
            "Acta sin firma presencial."
        ),
        "recibido_por": "Elizabeth González (confirmación telefónica, sin firma presencial)",
        "cargo": "Directora",
        "email_cliente": "dcampos.henriquez@cormup.cl",
        "email_cc": "",
        "enviar_correo_cliente": True,
        "checklist_cir": [
            {
                "elemento": "Conectividad App",
                "estado": "OK",
                "obs": "Revisión remota de la programación del viernes",
            },
        ],
        "checklist_cpa": [],
        "checklist_sab": [],
        "firma_png": "",
    }


def main() -> int:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    path = generar_pdf_acta(payload(), out_path=OUT)
    print(path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

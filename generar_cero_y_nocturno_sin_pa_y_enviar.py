"""
Genera:
- Reporte de puntos en cero (DOCX) usando reporte_puntos_en_cero.py
- Reporte control nocturno (CSV/DOCX/PDF) con HORARIOS CONTROL NOCTURNO.xlsx completo
  (incluye Parque Arauco / 000025), igual que enviar_reportes_diarios.py

Opcional --sin-pa: control nocturno excluyendo nodos PA (comportamiento antiguo).

Y envía los adjuntos por correo a Mauricio y Aníbal.

Uso:
  python generar_cero_y_nocturno_sin_pa_y_enviar.py
  python generar_cero_y_nocturno_sin_pa_y_enviar.py --no-email
  python generar_cero_y_nocturno_sin_pa_y_enviar.py --solo-control-nocturno --no-email
  python generar_cero_y_nocturno_sin_pa_y_enviar.py --sin-pa --no-email   # control sin PA
  python generar_cero_y_nocturno_sin_pa_y_enviar.py --solo-control-nocturno --sin-pa --fecha 06/04/2026 --no-email
"""

from __future__ import annotations

import argparse
import os
import sys
from datetime import date, datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

# Mismos correos que enviar_reportes_diarios.py (control nocturno / operaciones)
DESTINATARIOS = [
    "mauricioorellana@wes.cl",
    "anibal.aoperaciones@wes.cl",
]


def _is_parque_arauco_target(node_id: str, meta: Dict) -> bool:
    # 1) Excluir por companyId (Parque Arauco = 000025)
    if str(node_id).startswith("000025-"):
        return True
    # 2) Excluir por texto en CLIENTE
    cliente = str((meta or {}).get("cliente", "")).upper().strip()
    if "PARQUE ARAUCO" in cliente:
        return True
    if cliente in {"PA", "P.A"}:
        return True
    return False


def generar_reporte_puntos_en_cero() -> Path:
    from reporte_puntos_en_cero import (
        obtener_todos_los_nodos,
        verificar_consumo_cero,
        construir_resumen_alertas,
        crear_reporte_word,
    )
    from datetime import timedelta, timezone

    output_dir = Path("reporte en cero")
    output_dir.mkdir(exist_ok=True)

    todos_nodos = obtener_todos_los_nodos()
    if not todos_nodos:
        raise RuntimeError("No se encontraron nodos para reporte en cero.")

    puntos_en_cero = []
    puntos_sin_datos = []

    for nodo in todos_nodos:
        node_id = nodo["nodeId"]
        esta_en_cero, error = verificar_consumo_cero(node_id)
        if esta_en_cero:
            puntos_en_cero.append(nodo)
        elif error and "Sin datos" in error:
            puntos_sin_datos.append(nodo)

    fecha_alertas = datetime.now(timezone.utc) - timedelta(days=1)
    alertas_resumen = construir_resumen_alertas(todos_nodos, fecha_alertas)

    reporte_path = crear_reporte_word(
        puntos_en_cero,
        puntos_sin_datos,
        len(todos_nodos),
        output_dir,
        alertas_resumen=alertas_resumen,
        alertas_fecha=fecha_alertas.strftime("%d-%m-%Y"),
    )
    if not reporte_path or not Path(reporte_path).exists():
        raise RuntimeError("No se pudo generar el reporte de puntos en cero.")
    return Path(reporte_path)


def generar_control_nocturno_sin_pa(dia: Optional[date] = None) -> Tuple[Path, Path, Path]:
    from control_nocturno import (
        default_excel_path,
        cargar_targets_desde_excel,
        analizar_control_nocturno,
        guardar_csv,
        crear_reporte_word_control,
        convertir_docx_a_pdf,
        guardar_pdf_simple,
    )

    excel = default_excel_path()
    targets = cargar_targets_desde_excel(excel)
    targets_sin_pa = {
        nid: meta
        for nid, meta in targets.items()
        if not _is_parque_arauco_target(nid, meta)
    }

    hoy = dia if dia is not None else datetime.now().date()
    desde = datetime.combine(hoy, datetime.min.time())
    hasta = datetime.combine(hoy, datetime.min.time())

    rows = analizar_control_nocturno(targets_sin_pa, desde, hasta, umbral=0.0)

    base_dir = Path("reports") / "control_nocturno"
    base_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M")
    base_name = f"control_nocturno_sin_PA_{desde:%Y%m%d}_{hasta:%Y%m%d}_{ts}"
    out_csv = base_dir / f"{base_name}.csv"
    out_docx = base_dir / f"{base_name}.docx"
    out_pdf = base_dir / f"{base_name}.pdf"

    guardar_csv(rows, out_csv, targets_sin_pa)
    crear_reporte_word_control(rows, out_docx, desde, hasta, 0.0, targets_sin_pa)

    pdf_convertido = convertir_docx_a_pdf(out_docx)
    if pdf_convertido and pdf_convertido.exists():
        out_pdf = pdf_convertido
    else:
        guardar_pdf_simple(rows, out_pdf, desde, hasta, 0.0, targets_sin_pa)

    return out_csv, out_docx, out_pdf


def generar_control_nocturno_con_pa() -> Tuple[Path, Path, Path]:
    """Mismo criterio que enviar_reportes_diarios: Excel completo, incluye Parque Arauco."""
    from control_nocturno import generar_reporte_control_nocturno, default_excel_path

    hoy = datetime.now()
    desde = datetime(hoy.year, hoy.month, hoy.day)
    hasta = desde
    _rows, out_csv, out_docx, out_pdf = generar_reporte_control_nocturno(
        desde=desde,
        hasta=hasta,
        umbral=0.0,
        excel_path=default_excel_path(),
    )
    return out_csv, out_docx, out_pdf


def enviar_correo_adjuntos(
    archivos: List[Path],
    asunto: str,
    cuerpo: str,
    destinatarios: List[str],
) -> None:
    import smtplib
    from email.mime.application import MIMEApplication
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText

    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    smtp_user = "agente.ia@wes.cl"
    smtp_pass = (
        os.environ.get("WES_GMAIL_APP_PASSWORD")
        or os.environ.get("WES_SMTP_PASSWORD")
        or "vxbynfpoehbweelj"
    ).replace(" ", "").strip()

    to_header = ", ".join(destinatarios)
    msg = MIMEMultipart()
    msg["From"] = smtp_user
    msg["To"] = to_header
    msg["Subject"] = asunto
    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

    for fp in archivos:
        fp = Path(fp)
        if not fp.exists():
            continue
        with open(fp, "rb") as f:
            part = MIMEApplication(f.read())
            part.add_header("Content-Disposition", "attachment", filename=fp.name)
            msg.attach(part)

    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_pass)
        server.sendmail(smtp_user, destinatarios, msg.as_string())


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Puntos en cero + control nocturno (por defecto con PA); correo a Mauricio y Aníbal."
    )
    parser.add_argument(
        "--no-email",
        action="store_true",
        help="Generar archivos pero no enviar correo.",
    )
    parser.add_argument(
        "--solo-control-nocturno",
        action="store_true",
        help="Solo generar control nocturno (no ejecuta puntos en cero).",
    )
    parser.add_argument(
        "--sin-pa",
        action="store_true",
        help="Control nocturno excluyendo Parque Arauco (000025) y clientes PA en el Excel.",
    )
    parser.add_argument(
        "--fecha",
        metavar="DD/MM/YYYY",
        help="Día del control nocturno sin PA (requiere --sin-pa). Ej: 06/04/2026",
    )
    args = parser.parse_args()

    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass

    cero_docx: Path | None = None
    if not args.solo_control_nocturno:
        print("[INFO] Generando reporte de puntos en cero...")
        cero_docx = generar_reporte_puntos_en_cero()
        print(f"[OK] Puntos en cero: {cero_docx.resolve()}")

    dia_sin_pa: Optional[date] = None
    if args.fecha:
        if not args.sin_pa:
            print("[ERROR] --fecha solo aplica con --sin-pa (control nocturno sin PA).")
            return 1
        dia_sin_pa = datetime.strptime(args.fecha.strip(), "%d/%m/%Y").date()

    if args.sin_pa:
        print("[INFO] Generando control nocturno SIN Parque Arauco...")
        if dia_sin_pa:
            print(f"       Día: {dia_sin_pa:%d/%m/%Y}")
        cn_csv, cn_docx, cn_pdf = generar_control_nocturno_sin_pa(dia_sin_pa)
        print(f"[OK] Control nocturno sin PA: {cn_pdf.resolve()}")
    else:
        print("[INFO] Generando control nocturno (Excel completo, incluye Parque Arauco)...")
        cn_csv, cn_docx, cn_pdf = generar_control_nocturno_con_pa()
        print(f"[OK] Control nocturno (con PA): {cn_pdf.resolve()}")

    if args.no_email:
        print("[INFO] --no-email: no se envía correo.")
        return 0

    if args.solo_control_nocturno:
        if args.sin_pa:
            asunto = f"Reportes WES: Control nocturno sin PA — {datetime.now():%d-%m-%Y}"
            cuerpo = (
                "Estimados Mauricio y Aníbal,\n\n"
                "Adjuntos (control nocturno sin Parque Arauco):\n"
                f"- CSV: {cn_csv.name}\n"
                f"- Word: {cn_docx.name}\n"
                f"- PDF: {cn_pdf.name}\n\n"
                "Se excluyen nodos 000025-* y clientes PA según el Excel de horarios.\n\n"
                "Saludos,\n"
                "Agente IA WES"
            )
        else:
            asunto = f"Reportes WES: Control nocturno (con PA) — {datetime.now():%d-%m-%Y}"
            cuerpo = (
                "Estimados Mauricio y Aníbal,\n\n"
                "Adjuntos (control nocturno con Excel completo, incluye Parque Arauco):\n"
                f"- CSV: {cn_csv.name}\n"
                f"- Word: {cn_docx.name}\n"
                f"- PDF: {cn_pdf.name}\n\n"
                "Saludos,\n"
                "Agente IA WES"
            )
        adjuntos = [cn_csv, cn_docx, cn_pdf]
    else:
        assert cero_docx is not None
        if args.sin_pa:
            asunto = f"Reportes WES: Puntos en cero + Control nocturno sin PA — {datetime.now():%d-%m-%Y}"
            cuerpo = (
                "Estimados Mauricio y Aníbal,\n\n"
                "Adjuntos:\n"
                f"- Reporte de puntos en cero (DOCX): {cero_docx.name}\n"
                f"- Control nocturno sin Parque Arauco — CSV: {cn_csv.name}\n"
                f"- Control nocturno sin Parque Arauco — Word: {cn_docx.name}\n"
                f"- Control nocturno sin Parque Arauco — PDF: {cn_pdf.name}\n\n"
                "El control nocturno excluye nodos Parque Arauco (000025) y clientes PA en el Excel de horarios.\n"
                "El reporte de puntos en cero considera todos los puntos del sistema.\n\n"
                "Saludos,\n"
                "Agente IA WES"
            )
        else:
            asunto = f"Reportes WES: Puntos en cero + Control nocturno (con PA) — {datetime.now():%d-%m-%Y}"
            cuerpo = (
                "Estimados Mauricio y Aníbal,\n\n"
                "Adjuntos:\n"
                f"- Reporte de puntos en cero (DOCX): {cero_docx.name}\n"
                f"- Control nocturno (Excel completo, incluye Parque Arauco) — CSV: {cn_csv.name}\n"
                f"- Control nocturno — Word: {cn_docx.name}\n"
                f"- Control nocturno — PDF: {cn_pdf.name}\n\n"
                "El reporte de puntos en cero considera todos los puntos del sistema.\n\n"
                "Saludos,\n"
                "Agente IA WES"
            )
        adjuntos = [cero_docx, cn_csv, cn_docx, cn_pdf]

    print(f"[INFO] Enviando correo a: {', '.join(DESTINATARIOS)}")
    enviar_correo_adjuntos(
        adjuntos,
        asunto,
        cuerpo,
        DESTINATARIOS,
    )
    print("[OK] Correo enviado.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())


"""
Envía por correo a Aníbal el reporte agregado CORMUP y todos los reportes individuales
de un periodo (por defecto marzo 2026: 20260301-20260331).

Uso:
  python enviar_cormup_reportes_anibal.py
  python enviar_cormup_reportes_anibal.py --dry-run
  python enviar_cormup_reportes_anibal.py --inicio 20260301 --fin 20260331
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from pa_hallazgos_word_helpers import enviar_anibal_adjuntos

# Límite conservador para un solo correo (Gmail ~25 MB; dejamos margen base64).
_MAX_UN_CORREO_BYTES = 20 * 1024 * 1024


def _root() -> Path:
    return Path(__file__).resolve().parent


def _period_suffix(inicio: str, fin: str) -> str:
    return f"{inicio}_{fin}"


def _collect_cormup_docx(base: Path, sufijo_periodo: str) -> tuple[Path | None, list[Path]]:
    """
    Busca:
    - Agregado: .../ABREGADO/*/Reporte_Agregado_CORMUP_{inicio}_{fin}.docx (el más reciente si hay varios)
    - Individuales: .../REPORTE/**/Reporte_CORMUP_*_{inicio}_{fin}.docx
    """
    agregado_pattern = f"Reporte_Agregado_CORMUP_{sufijo_periodo}.docx"
    individual_glob = f"Reporte_CORMUP_*_{sufijo_periodo}.docx"

    agregado_dir = base / "CORMUP" / "ABREGADO"
    reporte_dir = base / "CORMUP" / "REPORTE"

    agregado: Path | None = None
    if agregado_dir.is_dir():
        candidatos = list(agregado_dir.rglob(agregado_pattern))
        if candidatos:
            agregado = max(candidatos, key=lambda p: p.stat().st_mtime)

    individuales: list[Path] = []
    if reporte_dir.is_dir():
        for p in reporte_dir.rglob(individual_glob):
            if p.is_file():
                individuales.append(p)

    individuales.sort(key=lambda p: p.name)

    return agregado, individuales


def _enviar_en_partes(
    agregado: Path | None,
    individuales: list[Path],
    sufijo: str,
    inicio: str,
    fin: str,
) -> None:
    """Si el peso total supera el límite, divide en varios correos (agregado aparte + lotes de individuales)."""
    todos: list[Path] = []
    if agregado:
        todos.append(agregado)
    todos.extend(individuales)
    total = sum(p.stat().st_size for p in todos)

    if total <= _MAX_UN_CORREO_BYTES or not todos:
        asunto = f"Reportes CORMUP (Peñalolén) — agregado e individuales — periodo {inicio} a {fin}"
        cuerpo = (
            "Aníbal,\n\n"
            f"Adjunto el reporte agregado y los reportes individuales de CORMUP/Peñalolén "
            f"para el periodo (sufijos de archivo {sufijo}).\n\n"
            f"Cantidad de adjuntos: {len(todos)}.\n\n"
            "Saludos,\n"
            "Sistema WES\n"
        )
        enviar_anibal_adjuntos(todos, asunto, cuerpo)
        return

    n = 0
    if agregado:
        n += 1
        asunto = f"[1/…] CORMUP agregado — {inicio}-{fin}"
        cuerpo = (
            "Aníbal,\n\n"
            f"Correo 1: solo reporte agregado CORMUP (periodo {sufijo}). "
            "Los individuales van en el siguiente correo.\n\n"
            "Saludos,\n"
            "Sistema WES\n"
        )
        enviar_anibal_adjuntos([agregado], asunto, cuerpo)

    if not individuales:
        return

    lotes: list[list[Path]] = []
    actual: list[Path] = []
    suma = 0
    for p in individuales:
        sz = p.stat().st_size
        if actual and suma + sz > _MAX_UN_CORREO_BYTES:
            lotes.append(actual)
            actual = []
            suma = 0
        actual.append(p)
        suma += sz
    if actual:
        lotes.append(actual)

    total_correos = (1 if agregado else 0) + len(lotes)
    idx = 2 if agregado else 1
    for lote in lotes:
        asunto = f"[{idx}/{total_correos}] CORMUP individuales — {inicio}-{fin} ({len(lote)} archivos)"
        cuerpo = (
            "Aníbal,\n\n"
            f"Adjunto lote de reportes individuales CORMUP/Peñalolén (periodo {sufijo}). "
            f"Archivos en este correo: {len(lote)}.\n\n"
            "Saludos,\n"
            "Sistema WES\n"
        )
        enviar_anibal_adjuntos(lote, asunto, cuerpo)
        idx += 1


def main() -> int:
    parser = argparse.ArgumentParser(description="Envía reportes CORMUP (agregado + individuales) a Aníbal.")
    parser.add_argument("--inicio", default="20260301", help="YYYYMMDD inicio (default marzo 2026)")
    parser.add_argument("--fin", default="20260331", help="YYYYMMDD fin (default marzo 2026)")
    parser.add_argument(
        "--reports-root",
        type=Path,
        default=None,
        help="Carpeta reports (default: <script>/reports)",
    )
    parser.add_argument("--dry-run", action="store_true", help="Solo lista archivos, no envía correo.")
    args = parser.parse_args()

    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
            sys.stderr.reconfigure(encoding="utf-8")
        except Exception:
            pass

    base = (args.reports_root or (_root() / "reports")).resolve()
    sufijo = _period_suffix(args.inicio, args.fin)

    agregado, individuales = _collect_cormup_docx(base, sufijo)

    print(f"[INFO] Buscando en: {base}")
    print(f"[INFO] Periodo archivo: *_{sufijo}.docx")

    if agregado:
        print(f"[OK] Agregado: {agregado}")
    else:
        print(f"[AVISO] No se encontró {base / 'CORMUP' / 'ABREGADO' / '...' / f'Reporte_Agregado_CORMUP_{sufijo}.docx'}")

    print(f"[OK] Individuales encontrados: {len(individuales)}")
    for p in individuales:
        print(f"       - {p.name}")

    archivos: list[Path] = []
    if agregado:
        archivos.append(agregado)
    archivos.extend(individuales)

    existentes = [p for p in archivos if p.exists()]
    if not existentes:
        print("[ERROR] No hay archivos .docx para enviar. Genera los reportes o revisa --inicio/--fin y --reports-root.")
        return 1

    total_bytes = sum(p.stat().st_size for p in existentes)
    print(f"[INFO] Total adjuntos: {len(existentes)} archivos, ~{total_bytes / (1024 * 1024):.2f} MB")

    if args.dry_run:
        print("[DRY-RUN] No se envió correo.")
        if total_bytes > _MAX_UN_CORREO_BYTES:
            print("[INFO] Con el tamaño actual se usarían varios correos (agregado + lotes).")
        return 0

    _enviar_en_partes(agregado, individuales, sufijo, args.inicio, args.fin)
    print("[OK] Envío completado a anibal.aoperaciones@wes.cl")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

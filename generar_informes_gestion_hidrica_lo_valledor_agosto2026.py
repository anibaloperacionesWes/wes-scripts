#!/usr/bin/env python3
"""Regenera gestión hídrica Lo Valledor (horario full 22:00–03:00)."""
from __future__ import annotations

import sys

from generar_informes_gestion_hidrica_lote_agosto2026 import CLIENTES, run_lote


def main() -> None:
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
            sys.stderr.reconfigure(encoding="utf-8", line_buffering=True)
        except Exception:
            pass
    clientes = [c for c in CLIENTES if c["key"] == "valledor"]
    run_lote(clientes, "GESTIÓN HÍDRICA — Lo Valledor — 01/08/2026 a 31/08/2026")


if __name__ == "__main__":
    main()

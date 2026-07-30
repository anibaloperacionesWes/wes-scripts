"""Rutas canónicas del proyecto WES (Google Drive)."""

from __future__ import annotations

import os
from pathlib import Path

WES_DRIVE_ROOT = Path(r"G:\Mi unidad\Agente WES\wes-scripts")


def wes_scripts_root() -> Path:
    env = os.environ.get("WES_SCRIPTS_ROOT", "").strip()
    if env:
        p = Path(env)
        if p.is_dir():
            return p.resolve()
    if WES_DRIVE_ROOT.is_dir():
        return WES_DRIVE_ROOT.resolve()
    return Path(__file__).resolve().parent


def reporte_cero_dir() -> Path:
    return wes_scripts_root() / "reporte en cero"

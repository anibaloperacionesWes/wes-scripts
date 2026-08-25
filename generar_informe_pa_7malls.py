# -*- coding: utf-8 -*-
"""
Regenera el Informe WES Parque Arauco 7 Malls con data 01/06/2026–16/08/2026.

Copia el deck 07.07.2026 a un archivo nuevo (no lo pisa) y corre
``editar_lamina_pa.py --todas``. Entrega en:

  reports/Parque_Arauco/TMP_7MALLS/entrega_diego_anibal/

Uso:
  python generar_informe_pa_7malls.py
"""

from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "reports" / "_tmp_pa_7malls_charts" / (
    "Informe WES __ Parque Arauco 7 Malls (07.07.2026).pptx"
)
WORK = ROOT / "reports" / "_tmp_pa_7malls_charts" / (
    "Informe WES __ Parque Arauco 7 Malls (16.08.2026).pptx"
)
ENTREGA = ROOT / "reports" / "Parque_Arauco" / "TMP_7MALLS" / "entrega_diego_anibal"
ENTREGA_TMP = ROOT / "reports" / "_tmp_pa_7malls_charts" / "entrega_diego_anibal"


def main() -> int:
    if not SRC.is_file():
        print(f"[ERROR] No está el deck original: {SRC}", flush=True)
        return 1
    WORK.parent.mkdir(parents=True, exist_ok=True)
    if not WORK.is_file():
        shutil.copy2(SRC, WORK)
        print(f"[OK] Copia de trabajo: {WORK}", flush=True)
    else:
        print(f"[OK] Usa copia existente: {WORK}", flush=True)

    cmd = [sys.executable, str(ROOT / "editar_lamina_pa.py"), "--todas"]
    print("[INFO] Regenerando láminas…", flush=True)
    rc = subprocess.call(cmd)
    if rc != 0:
        print(f"[AVISO] Regeneración terminó con código {rc}", flush=True)

    if not WORK.is_file():
        print("[ERROR] No se generó el PPT de trabajo", flush=True)
        return 1
    for dest_dir in (ENTREGA, ENTREGA_TMP):
        dest_dir.mkdir(parents=True, exist_ok=True)
        dest = dest_dir / WORK.name
        shutil.copy2(WORK, dest)
        print(f"[OK] Entrega: {dest}", flush=True)
    return 0 if rc == 0 else rc


if __name__ == "__main__":
    raise SystemExit(main())

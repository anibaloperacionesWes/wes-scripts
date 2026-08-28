# -*- coding: utf-8 -*-
"""
Puntos CPA (monitoreo con control): sitios con horario programado en el Excel
de Drive «Horarios de contron hidrico clientes wes».

El resto de puntos activos WES se consideran solo monitoreo.
"""

from __future__ import annotations

import io
import re
import unicodedata
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Set, Tuple

DRIVE_HORARIOS_ID = "1eM03xh4Pmqx5YTTKWiOku8_NT0115ausiWOKNaLc8Lk"
CACHE_XLSX = Path(__file__).resolve().parent / "reports" / "Horarios_control_hidrico_clientes_wes.xlsx"

# Nombres del Excel de horarios → nodeId WES.
CPA_ALIAS = {
    "171 antonio hermidas fabres": "000008-01",
    "antonio hermidas fabres": "000008-01",
    "antonio hermida fabres": "000008-01",
    "carlos fernandes pena": "000008-03",
    "carlos fernandez pena": "000008-03",
    "colegio tobalaba": "000008-04",
    "tobalaba": "000008-04",
    "colegio santa maria": "000008-05",
    "santa maria": "000008-05",
    "pae estanquer norte locales": "000025-01",
    "pae estanque norte locales": "000025-01",
    "estanque norte locales": "000025-01",
    "mae sala de bomba estanque sur": "000025-19",
    "sala de bomba estanque sur": "000025-19",
    "pizza hut": "000025-07",
    "liceo alto cordillera la florida": "000028-01",
    "alto cordillera": "000028-01",
    "lo valledor p1": "000002-01",
    "club hause uc": "000021-01",
    "club house uc": "000021-01",
    "club house cduc": "000021-01",
    "raymundo tupper": "000021-03",
    "raimundo tupper": "000021-03",
    "agunsa modulo d": "000020-02",
    "modulo d": "000020-02",
    "iccp": "000017-07",
    "icco": "000017-08",
    "eugenio maria de hostos": "000024-01",
}

SKIP_TITULOS = {
    "corporacion penalolen colegios",
    "corporacion penalolen",
}


def _slug(texto: str) -> str:
    t = unicodedata.normalize("NFKD", texto or "")
    t = "".join(ch for ch in t if not unicodedata.combining(ch))
    t = t.lower()
    t = re.sub(r"[^a-z0-9]+", " ", t)
    return t.strip()


def _es_celda_horario(val: object) -> bool:
    s = str(val or "").strip().lower()
    if not s:
        return True
    if s.startswith("#######") or s.startswith("ev."):
        return True
    if re.match(r"^\d+\s*=", s):
        return True
    return False


def _titulos_desde_xlsx(path: Path) -> List[str]:
    from openpyxl import load_workbook

    wb = load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    out: List[str] = []
    for row in ws.iter_rows(max_col=12, values_only=True):
        for val in row:
            if _es_celda_horario(val):
                continue
            t = str(val).strip()
            slug = _slug(t)
            if not slug or slug in SKIP_TITULOS:
                continue
            out.append(t)
    return out


def _descargar_drive(destino: Path) -> Path:
    from googleapiclient.http import MediaIoBaseDownload

    from wes_google_drive import obtener_servicio_drive

    svc = obtener_servicio_drive()
    request = svc.files().export_media(
        fileId=DRIVE_HORARIOS_ID,
        mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _status, done = downloader.next_chunk()
    destino.parent.mkdir(parents=True, exist_ok=True)
    destino.write_bytes(buf.getvalue())
    return destino


def _resolver_node_id(
    titulo: str,
    nombres: Dict[str, str],
) -> Optional[str]:
    slug = _slug(titulo)
    if slug in CPA_ALIAS:
        return CPA_ALIAS[slug]
    # alias parcial
    for key, nid in CPA_ALIAS.items():
        if key in slug or slug in key:
            return nid
    tokens = {t for t in slug.split() if len(t) >= 4}
    if not tokens:
        return None
    best_nid = None
    best_score = 0
    for nid, name in nombres.items():
        nslug = _slug(name)
        ntok = {t for t in nslug.split() if len(t) >= 4}
        score = len(tokens & ntok)
        if score > best_score:
            best_score = score
            best_nid = nid
    if best_score >= 2:
        return best_nid
    return None


def cargar_ids_cpa(
    nombres_nodos: Optional[Dict[str, str]] = None,
    *,
    excel: Optional[Path] = None,
) -> Tuple[Set[str], List[Tuple[str, str, str]]]:
    """
    Returns:
      set de nodeId CPA
      lista (titulo_excel, nodeId, nombre_wes) para auditoría
    """
    path = excel or CACHE_XLSX
    if excel is None:
        try:
            path = _descargar_drive(CACHE_XLSX)
            print(f"[INFO] Horarios CPA desde Drive → {path}")
        except Exception as exc:
            if path.is_file():
                print(f"[WARN] No se pudo bajar Drive ({exc}); uso caché {path}")
            else:
                raise

    nombres = nombres_nodos or {}
    titulos = _titulos_desde_xlsx(path)
    ids: Set[str] = set()
    detalle: List[Tuple[str, str, str]] = []
    for titulo in titulos:
        nid = _resolver_node_id(titulo, nombres)
        if not nid:
            detalle.append((titulo, "", "NO MAPEADO"))
            print(f"[WARN] Título CPA sin nodeId: {titulo}")
            continue
        ids.add(nid)
        detalle.append((titulo, nid, nombres.get(nid, "")))
    return ids, detalle


def clasificar_nodos(
    node_ids: Iterable[str],
    ids_cpa: Set[str],
) -> Tuple[List[str], List[str]]:
    cpa: List[str] = []
    solo: List[str] = []
    for nid in sorted(set(node_ids)):
        if nid in ids_cpa:
            cpa.append(nid)
        else:
            solo.append(nid)
    return cpa, solo

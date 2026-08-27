# -*- coding: utf-8 -*-
"""
Excel de accesos WES: una hoja por empresa con quienes tienen allowedNodes
en puntos activos de esa empresa.

Reutiliza la misma muestra de correos que contar_usuarios_por_nodo.py
(repo + contactos + receivers FILTRATION).

Uso:
  python usuarios_acceso_por_empresa.py
"""

from __future__ import annotations

import argparse
import re
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Set, Tuple
from zoneinfo import ZoneInfo

from contar_usuarios_por_nodo import (
    _allowed_nodes,
    _emails_desde_alertas,
    _fetch_user,
    _nodos_y_empresas,
    _recolectar_emails,
    _session,
    es_nodo_activo,
)

ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "reports" / "Usuarios" / "acceso_por_empresa"
TZ_CL = ZoneInfo("America/Santiago")
SHEET_BAD = re.compile(r"[\\/*?:\[\]]+")
NOMBRE_EMPRESA_FALLBACK = {
    "000026": "UDD",
}


def _nombre_hoja(cid: str, empresa: str, usados: Set[str]) -> str:
    base = SHEET_BAD.sub(" ", f"{cid} {empresa}").strip()
    base = re.sub(r"\s+", " ", base)[:31] or cid
    name = base
    n = 2
    while name.casefold() in usados:
        suffix = f"_{n}"
        name = (base[: 31 - len(suffix)] + suffix)[:31]
        n += 1
    usados.add(name.casefold())
    return name


def _cargar_usuarios(extra_email_files: List[Path], workers: int):
    session = _session()
    print("[INFO] Obteniendo empresas y nodos...")
    nombres, company_de, companies = _nodos_y_empresas(session)
    for cid, name in list(companies.items()):
        if not name or name.strip() == cid:
            companies[cid] = NOMBRE_EMPRESA_FALLBACK.get(cid, name or cid)
    print(f"[OK] {len(companies)} empresas · {len(nombres)} nodeId")

    emails = _recolectar_emails(extra_email_files)
    extra = _emails_desde_alertas(session, nombres.keys(), workers=workers)
    if extra:
        n0 = len(emails)
        emails = sorted(set(emails) | extra)
        print(f"[INFO] Correos extra por alertas: {len(emails) - n0} (total {len(emails)})")
    print(f"[INFO] Consultando {len(emails)} correo(s)...")

    usuarios: Dict[str, dict] = {}
    with ThreadPoolExecutor(max_workers=workers) as ex:
        futs = [ex.submit(_fetch_user, session, e) for e in emails]
        for fut in as_completed(futs):
            email, user, _obs = fut.result()
            if not user:
                continue
            email_api = str(user.get("username") or email).strip().lower()
            usuarios[email_api] = user
            print(f"[OK] {email_api}")
    return nombres, company_de, companies, usuarios


def _accesos_por_empresa(
    nombres: Dict[str, str],
    company_de: Dict[str, str],
    companies: Dict[str, str],
    usuarios: Dict[str, dict],
    solo_activos: bool,
) -> Tuple[Dict[str, List[str]], Dict[str, List[dict]]]:
    nodos_empresa: Dict[str, List[str]] = defaultdict(list)
    for nid, cid in company_de.items():
        cname = companies.get(cid, "")
        if solo_activos and not es_nodo_activo(nid, cid, cname):
            continue
        nodos_empresa[cid].append(nid)
    for cid in nodos_empresa:
        nodos_empresa[cid] = sorted(set(nodos_empresa[cid]))

    filas: Dict[str, List[dict]] = defaultdict(list)
    for email, user in usuarios.items():
        nombre = f"{user.get('name', '')} {user.get('lastName', '')}".strip()
        cid_user = str(user.get("companyId") or "").strip()
        switch = user.get("switchEnabled")
        por_cia: Dict[str, List[str]] = defaultdict(list)
        for nid in _allowed_nodes(user):
            cid = company_de.get(nid, "")
            if not cid:
                continue
            cname = companies.get(cid, "")
            if solo_activos and not es_nodo_activo(nid, cid, cname):
                continue
            por_cia[cid].append(nid)
        for cid, nids in por_cia.items():
            nids = sorted(set(nids))
            filas[cid].append(
                {
                    "email": email,
                    "nombre": nombre,
                    "company_id_cuenta": cid_user,
                    "empresa_cuenta": companies.get(cid_user, cid_user),
                    "switch_on_off": "Sí" if switch else "No",
                    "nodos_count": len(nids),
                    "node_ids": ", ".join(nids),
                    "nombres_puntos": "; ".join(
                        f"{nid} ({nombres.get(nid) or nid})" for nid in nids
                    ),
                    "es_cuenta_de_esta_empresa": "Sí" if cid_user == cid else "No",
                }
            )
    for cid in filas:
        filas[cid].sort(
            key=lambda r: (
                0 if r["es_cuenta_de_esta_empresa"] == "Sí" else 1,
                r["email"].casefold(),
            )
        )
    return dict(nodos_empresa), dict(filas)


def _escribir_xlsx(
    path: Path,
    companies: Dict[str, str],
    nodos_empresa: Dict[str, List[str]],
    filas: Dict[str, List[dict]],
    generado: str,
) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)
    wrap = Alignment(wrap_text=True, vertical="top")

    cols_res = [
        "company_id",
        "empresa",
        "puntos_activos",
        "usuarios_con_acceso",
        "hoja",
    ]
    usados: Set[str] = {"resumen"}
    hojas: List[Tuple[str, str, str]] = []  # cid, empresa, sheet name
    for cid, empresa in sorted(companies.items(), key=lambda x: (x[1].casefold(), x[0])):
        if cid not in nodos_empresa:
            continue
        sheet = _nombre_hoja(cid, empresa, usados)
        hojas.append((cid, empresa, sheet))

    ws = wb.active
    ws.title = "Resumen"
    ws["A1"] = "Accesos WES por empresa (puntos activos)"
    ws["A1"].font = Font(bold=True, size=14, color="1F4E79")
    ws["A2"] = f"Generado {generado} hora Chile · una hoja por empresa"
    ws.merge_cells("A1:E1")
    ws.merge_cells("A2:E2")
    for c, h in enumerate(cols_res, 1):
        cell = ws.cell(row=4, column=c, value=h)
        cell.fill = header_fill
        cell.font = header_font
    for i, (cid, empresa, sheet) in enumerate(hojas, 5):
        ws.cell(row=i, column=1, value=cid)
        ws.cell(row=i, column=2, value=empresa)
        ws.cell(row=i, column=3, value=len(nodos_empresa.get(cid, [])))
        ws.cell(row=i, column=4, value=len(filas.get(cid, [])))
        ws.cell(row=i, column=5, value=sheet)
    for i, w in enumerate((12, 32, 16, 20, 28), 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:E{4 + max(1, len(hojas))}"

    cols = [
        "email",
        "nombre",
        "es_cuenta_de_esta_empresa",
        "empresa_cuenta",
        "switch_on_off",
        "nodos_count",
        "node_ids",
        "nombres_puntos",
    ]
    titulos = [
        "Email",
        "Nombre",
        "Cuenta de esta empresa",
        "Empresa de la cuenta",
        "Switch on/off",
        "Puntos con acceso",
        "Node IDs",
        "Puntos (nombre)",
    ]
    widths = (34, 28, 22, 28, 14, 16, 40, 70)

    for cid, empresa, sheet in hojas:
        wsc = wb.create_sheet(sheet)
        wsc["A1"] = f"{empresa} ({cid})"
        wsc["A1"].font = Font(bold=True, size=14, color="1F4E79")
        wsc.merge_cells("A1:H1")
        n_pts = len(nodos_empresa.get(cid, []))
        n_usr = len(filas.get(cid, []))
        wsc["A2"] = f"{n_usr} usuario(s) con acceso · {n_pts} punto(s) activo(s)"
        wsc.merge_cells("A2:H2")
        for c, h in enumerate(titulos, 1):
            cell = wsc.cell(row=4, column=c, value=h)
            cell.fill = header_fill
            cell.font = header_font
        rows = filas.get(cid, [])
        if not rows:
            wsc.cell(row=5, column=1, value="Sin usuarios en la muestra consultada")
        for r_i, row in enumerate(rows, 5):
            for c, key in enumerate(cols, 1):
                cell = wsc.cell(row=r_i, column=c, value=row.get(key, ""))
                cell.alignment = wrap
        for i, w in enumerate(widths, 1):
            wsc.column_dimensions[get_column_letter(i)].width = w
        wsc.freeze_panes = "A5"
        last = 4 + max(1, len(rows))
        wsc.auto_filter.ref = f"A4:H{last}"
        wsc.row_dimensions[4].height = 18

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Excel: una hoja por empresa con usuarios que tienen acceso"
    )
    parser.add_argument("--emails", type=Path, action="append", default=[])
    parser.add_argument("--salida", type=Path, default=OUT_DIR)
    parser.add_argument("--workers", type=int, default=10)
    parser.add_argument(
        "--incluir-inactivos",
        action="store_true",
        help="Incluir puntos dados de baja / fuera de operación",
    )
    args = parser.parse_args()

    nombres, company_de, companies, usuarios = _cargar_usuarios(
        args.emails, max(1, min(args.workers, 16))
    )
    nodos_empresa, filas = _accesos_por_empresa(
        nombres, company_de, companies, usuarios, solo_activos=not args.incluir_inactivos
    )

    args.salida.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now(TZ_CL).strftime("%Y%m%d_%H%M")
    generado = datetime.now(TZ_CL).strftime("%d/%m/%Y %H:%M")
    xlsx = args.salida / f"accesos_por_empresa_{stamp}.xlsx"
    latest = args.salida / "accesos_por_empresa.xlsx"
    _escribir_xlsx(xlsx, companies, nodos_empresa, filas, generado)
    latest.write_bytes(xlsx.read_bytes())

    n_emp = sum(1 for cid in companies if cid in nodos_empresa)
    n_usr = len(usuarios)
    print(f"\n[OK] Empresas con puntos activos: {n_emp}")
    print(f"[OK] Usuarios encontrados: {n_usr}")
    for cid, empresa in sorted(companies.items(), key=lambda x: x[1].casefold()):
        if cid not in nodos_empresa:
            continue
        print(
            f"  {cid} {empresa}: {len(filas.get(cid, []))} usuario(s) · "
            f"{len(nodos_empresa[cid])} punto(s)"
        )
    print(f"[OK] XLSX: {xlsx}")
    print(f"[OK] XLSX: {latest}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

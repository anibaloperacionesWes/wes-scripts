"""
Extracción de períodos de facturación desde PDF Aguas Andinas.

- Boleta / factura electrónica con capa de texto: mismos patrones que los otros informes Renca.
- Historial de consumo (típicamente una página escaneada o renderizada a imagen): OCR con RapidOCR
  y heurística de tabla «HISTORIAL DE CONSUMO» (fechas pegadas DD/MM/YYYY + DD-MM-YYYY).
  Use ``listar_filas_historial_consumo_facturacion_desde_pdf`` para Excel con columnas
  Fecha Lectura, Lectura, M3 Consumos, Facturación del servicio, Total de cuenta.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import List, Optional

from pypdf import PdfReader

_MESES = {
    "ENE": 1,
    "JAN": 1,
    "FEB": 2,
    "MAR": 3,
    "ABR": 4,
    "APR": 4,
    "MAY": 5,
    "JUN": 6,
    "JUL": 7,
    "AGO": 8,
    "AUG": 8,
    "SEP": 9,
    "OCT": 10,
    "NOV": 11,
    "DIC": 12,
    "DEC": 12,
}

_RAPID_OCR = None


def _get_rapid_ocr():
    global _RAPID_OCR
    if _RAPID_OCR is None:
        from rapidocr_onnxruntime import RapidOCR

        _RAPID_OCR = RapidOCR()
    return _RAPID_OCR


def _parse_fecha_es_dd_mmm_yyyy(s: str) -> datetime:
    p = s.strip().upper().split("-")
    if len(p) != 3:
        raise ValueError(f"Fecha inválida: {s!r}")
    dd = int(p[0])
    mm = _MESES[p[1][:3]]
    yy = int(p[2])
    return datetime(yy, mm, dd)


def _parse_dd_mm_yyyy_slash(s: str) -> datetime:
    p = s.strip().split("/")
    if len(p) != 3:
        raise ValueError(f"Fecha inválida: {s!r}")
    dd, mm, yy = int(p[0]), int(p[1]), int(p[2])
    return datetime(yy, mm, dd)


def _parse_dd_mm_yyyy_dash(s: str) -> datetime:
    p = s.strip().split("-")
    if len(p) != 3:
        raise ValueError(f"Fecha inválida: {s!r}")
    dd, mm, yy = int(p[0]), int(p[1]), int(p[2])
    return datetime(yy, mm, dd)


def extraer_texto_pdf(path: Path) -> str:
    reader = PdfReader(str(path))
    parts = [(pg.extract_text() or "") for pg in reader.pages]
    txt = "\n".join(parts)
    if len(re.sub(r"\s+", "", txt)) >= 80:
        return txt
    try:
        import fitz
    except ImportError as e:
        raise ValueError(
            "El PDF no tiene texto extraíble. Instale PyMuPDF y rapidocr-onnxruntime: pip install pymupdf rapidocr-onnxruntime"
        ) from e
    ocr = _get_rapid_ocr()
    out_chunks: List[str] = []
    doc = fitz.open(path)
    try:
        mat = fitz.Matrix(2.0, 2.0)
        for page in doc:
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img_bytes = pix.tobytes("png")
            result, _elapsed = ocr(img_bytes)
            if not result:
                continue
            lines = [item[1] for item in result]
            out_chunks.append("\n".join(lines))
    finally:
        doc.close()
    return "\n".join(out_chunks)


@dataclass(frozen=True)
class PeriodoFacturacionAA:
    """Un período comparable (lecturas + m³) extraído de PDF/CSV."""

    pdf: Path
    boleta: str
    cuenta: Optional[str]
    medidor: Optional[str]
    emision: datetime
    lectura_anterior: datetime
    lectura_actual: datetime
    m3_cuenta: int


def _parse_boleta_estandar(txt: str, pdf: Path) -> Optional[PeriodoFacturacionAA]:
    boleta_m = re.search(
        r"(?:FACTURA|BOLETA)\s+ELECTR[ÓO]NICA\s*\n?\s*N[º°]\s*([0-9]{6,})",
        txt,
        flags=re.IGNORECASE,
    )
    if not boleta_m:
        boleta_m = re.search(r"\nN[º°]\s*([0-9]{6,})\n", txt, flags=re.IGNORECASE)
    emision_m = re.search(
        r"FECHA\s+EMISI[ÓO]N[:\s]*([0-9]{2}-[A-Z]{3}-[0-9]{4})",
        txt,
        flags=re.IGNORECASE,
    )
    cuenta_m = re.search(r"\n([0-9]{7,}-[0-9])\n", txt)
    medidor_m = re.search(r"N[úu]mero\s+de\s+Medidor\s+([0-9\.]+)", txt, flags=re.IGNORECASE)
    consumo_m = re.search(r"CONSUMO\s+TOTAL\s+([0-9\.\,]+)\s*m3", txt, flags=re.IGNORECASE)
    la_m = re.search(
        r"LECTURA\s+ACTUAL\s+([0-9]{2}-[A-Z]{3}-[0-9]{4})",
        txt,
        flags=re.IGNORECASE,
    )
    lan_m = re.search(
        r"LECTURA\s+ANTERIOR\s+([0-9]{2}-[A-Z]{3}-[0-9]{4})",
        txt,
        flags=re.IGNORECASE,
    )
    if not (emision_m and consumo_m and la_m and lan_m):
        return None
    dt_emision = _parse_fecha_es_dd_mmm_yyyy(emision_m.group(1))
    dt_actual = _parse_fecha_es_dd_mmm_yyyy(la_m.group(1))
    dt_anterior = _parse_fecha_es_dd_mmm_yyyy(lan_m.group(1))
    m3_cuenta = int(float(consumo_m.group(1).replace(".", "").replace(",", ".")))
    return PeriodoFacturacionAA(
        pdf=pdf,
        boleta=(boleta_m.group(1) if boleta_m else pdf.stem),
        cuenta=(cuenta_m.group(1) if cuenta_m else None),
        medidor=(medidor_m.group(1) if medidor_m else None),
        emision=dt_emision,
        lectura_anterior=dt_anterior,
        lectura_actual=dt_actual,
        m3_cuenta=m3_cuenta,
    )


def _es_historial_consumo(txt: str) -> bool:
    u = re.sub(r"\s+", "", txt.upper())
    return "HISTORIALDECONSUMO" in u or ("HISTORIAL" in u and "CONSUMO" in u and "AGUAS" in u)


def _m3_tras_fecha(txt: str, pos: int) -> Optional[int]:
    sub = txt[pos : pos + 450]
    lines = [ln.strip() for ln in sub.split("\n")]
    seen_reading = False
    for ln in lines[:18]:
        if not ln or ln.upper() in ("CORPORATIVO", "NO"):
            if ln.upper() == "CORPORATIVO":
                break
            continue
        if re.fullmatch(r"\d{1,3}\.\d{3}", ln):
            seen_reading = True
            continue
        if re.fullmatch(r"\d{2,4}", ln):
            val = int(ln)
            if 40 <= val <= 6000:
                if seen_reading or val <= 2500:
                    return val
        if re.fullmatch(r"\d{5,7}", ln):
            seen_reading = True
    return None


def _parse_entero_chile_linea(ln: str) -> int | None:
    """Entero desde línea OCR tipo ``1.188.194`` o ``878``."""
    ln = ln.strip().replace(" ", "")
    if not ln or ln.upper() in ("CORPORATIVO", "NO"):
        return None
    if not re.fullmatch(r"[\d\.]+", ln):
        return None
    return int(ln.replace(".", ""))


@dataclass(frozen=True)
class HistorialFilaConsumoFacturacion:
    """Fila del historial: columnas Fecha Lectura, Lectura, M3, Facturación del servicio, Total cuenta."""

    fecha_lectura: datetime
    lectura: int
    m3_consumos: int
    facturacion_servicio: int
    total_cuenta: int


def _filas_historial_cinco_columnas_desde_texto(txt: str) -> list[HistorialFilaConsumoFacturacion]:
    """Parsea bloques ``DD/MM/AAAA`` + ``DD-MM-AAAA`` y líneas numéricas siguientes (layout OCR Aguas Andinas)."""
    u = txt.upper()
    cut = u.find("FACTURACIONES")
    if cut != -1:
        txt = txt[cut:]
    pat = re.compile(r"(\d{2}/\d{2}/\d{4})\s*(\d{2}-\d{2}-\d{4})")
    matches = list(pat.finditer(txt))
    filas: list[HistorialFilaConsumoFacturacion] = []
    visto_fin: set[tuple[int, int, int]] = set()

    for i, m in enumerate(matches):
        try:
            lec_dt = _parse_dd_mm_yyyy_dash(m.group(2))
        except ValueError:
            continue
        key = (lec_dt.year, lec_dt.month, lec_dt.day)
        if key in visto_fin:
            continue
        end = matches[i + 1].start() if i + 1 < len(matches) else len(txt)
        chunk = txt[m.end() : end]
        nums: list[int] = []
        for ln in chunk.splitlines():
            v = _parse_entero_chile_linea(ln)
            if v is not None:
                nums.append(v)
        if len(nums) < 3:
            continue
        lectura, m3, fact = nums[0], nums[1], nums[2]
        rest = nums[3:]
        while rest and rest[-1] == 0:
            rest.pop()
        total = rest[-1] if rest else 0
        # Descarta coincidencias fuera de la tabla de montos (p. ej. ruido OCR)
        if fact < 1_000 or lectura < 10:
            continue
        visto_fin.add(key)
        filas.append(
            HistorialFilaConsumoFacturacion(
                fecha_lectura=lec_dt,
                lectura=lectura,
                m3_consumos=m3,
                facturacion_servicio=fact,
                total_cuenta=total,
            )
        )
    filas.sort(key=lambda f: f.fecha_lectura)
    return filas


def listar_filas_historial_consumo_facturacion_desde_pdf(
    path: Path,
) -> tuple[Optional[str], Optional[str], list[HistorialFilaConsumoFacturacion]] | None:
    """Historial de consumo/facturaciones: cuenta, medidor y filas con las 5 columnas estándar del documento."""
    txt = extraer_texto_pdf(path)
    if not _es_historial_consumo(txt):
        return None
    cuenta_m = re.search(r"Nro\.?\s*Cuenta\s*([0-9]{6,}-[0-9])", txt, flags=re.IGNORECASE)
    medidor_m = re.search(r"Nro\.?\s*Medidor\s*([0-9]{7,})", txt, flags=re.IGNORECASE)
    cuenta = cuenta_m.group(1).strip() if cuenta_m else None
    medidor = medidor_m.group(1).strip() if medidor_m else None
    filas = _filas_historial_cinco_columnas_desde_texto(txt)
    if not filas:
        return None
    return cuenta, medidor, filas


def _historial_dedup_triplets(txt: str) -> list[tuple[datetime, datetime, int]]:
    """Filas únicas (emisión, fecha lectura, m³) en orden cronológico por fecha de lectura."""
    hits: list[tuple[int, datetime, datetime, int]] = []
    for m in re.finditer(r"(\d{2}/\d{2}/\d{4})\s*(\d{2}-\d{2}-\d{4})", txt):
        try:
            emi = _parse_dd_mm_yyyy_slash(m.group(1))
            fin = _parse_dd_mm_yyyy_dash(m.group(2))
        except ValueError:
            continue
        m3 = _m3_tras_fecha(txt, m.end())
        if m3 is None:
            continue
        hits.append((m.start(), emi, fin, m3))

    hits.sort(key=lambda x: x[0])
    dedup: list[tuple[datetime, datetime, int]] = []
    seen_fin: set[tuple[int, int, int]] = set()
    for _pos, emi, fin, m3 in hits:
        key = (fin.year, fin.month, fin.day)
        if key in seen_fin:
            continue
        seen_fin.add(key)
        dedup.append((emi, fin, m3))

    dedup.sort(key=lambda x: x[1])
    return dedup


def _parse_historial_consumo(txt: str, pdf: Path) -> List[PeriodoFacturacionAA]:
    """Filas del PDF historial: documento suele ir del período más nuevo al más viejo."""
    cuenta_m = re.search(r"Nro\.?\s*Cuenta\s*([0-9]{6,}-[0-9])", txt, flags=re.IGNORECASE)
    medidor_m = re.search(r"Nro\.?\s*Medidor\s*([0-9]{7,})", txt, flags=re.IGNORECASE)
    cuenta = cuenta_m.group(1).strip() if cuenta_m else None
    medidor = medidor_m.group(1).strip() if medidor_m else None

    dedup = _historial_dedup_triplets(txt)
    out: List[PeriodoFacturacionAA] = []
    for i in range(1, len(dedup)):
        _emi_prev, fin_prev, _m3_prev = dedup[i - 1]
        emi_i, fin_i, m3_i = dedup[i]
        out.append(
            PeriodoFacturacionAA(
                pdf=pdf,
                boleta=f"HIST-{fin_i.strftime('%Y%m%d')}",
                cuenta=cuenta,
                medidor=medidor,
                emision=emi_i,
                lectura_anterior=fin_prev,
                lectura_actual=fin_i,
                m3_cuenta=m3_i,
            )
        )
    return out


def listar_periodos_desde_pdf(path: Path) -> List[PeriodoFacturacionAA]:
    txt = extraer_texto_pdf(path)
    std = _parse_boleta_estandar(txt, path)
    if std is not None:
        return [std]
    if _es_historial_consumo(txt):
        rows = _parse_historial_consumo(txt, path)
        if rows:
            return rows
    raise ValueError(
        f"No se pudieron extraer períodos de facturación desde {path.name} "
        "(ni boleta estándar ni historial de consumo reconocible)."
    )

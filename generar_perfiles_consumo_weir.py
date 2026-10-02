"""
Perfil de consumo Weir / Vulco (San Bernardo) a partir de las boletas
Aguas Andinas de septiembre 2026.

El gráfico «Consumo últimos 13 meses» de cada boleta no trae tabla numérica.
Se mide la altura de cada barra y se escala para que el último mes coincida
con los m³ facturados en esa boleta.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pymupdf
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Inches, Pt, RGBColor
from PIL import Image
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

ROOT = Path(__file__).resolve().parent
BOLETAS = ROOT / "reports" / "Weir" / "boletas_202609"
OUT = ROOT / "reports" / "Weir" / "Perfil_consumo"
UPLOADS = Path("/home/ubuntu/.cursor/projects/workspace/uploads")

MESES = [
    "Sep-25",
    "Oct-25",
    "Nov-25",
    "Dic-25",
    "Ene-26",
    "Feb-26",
    "Mar-26",
    "Abr-26",
    "May-26",
    "Jun-26",
    "Jul-26",
    "Ago-26",
    "Sep-26",
]

pdfmetrics.registerFont(TTFont("DejaVu", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"))
pdfmetrics.registerFont(TTFont("DejaVuBold", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"))

NAVY = (14 / 255, 58 / 255, 93 / 255)
TEAL = (0 / 255, 122 / 255, 140 / 255)
BAR = "#1F6AA5"
NAVY_HEX = RGBColor(14, 58, 93)


@dataclass
class Cuenta:
    src_name: str
    archivo: str
    titulo: str
    sitio: str
    direccion: str
    titular: str
    rut: str
    cuenta: str
    medidor: str
    documento: str
    lectura: str
    consumo_m3: int
    total_pagar: str
    clave: str
    nota_extra: str
    lectura_perfil: str
    es_control: bool = False
    dias_ultimo: int | None = None


CUENTAS = [
    Cuenta(
        src_name="Baquedano_1215_9189.pdf",
        archivo="weir Baquedano 1215",
        titulo="Baquedano 1215",
        sitio="Estadio Vulco",
        direccion="Baquedano 1215, San Bernardo",
        titular="Vulco S.A.",
        rut="91.619.000-K",
        cuenta="295661-6",
        medidor="110176970",
        documento="Factura electrónica N° 9455958",
        lectura="31-07-2026 al 31-08-2026",
        consumo_m3=397,
        total_pagar="$ 550.552",
        clave="Consumo real · lectura ratificada",
        nota_extra="Despacho electrónico: weirbills01@weir.urjanet.com. Arranque 40 mm.",
        lectura_perfil=(
            "Es el medidor de control de Estadio Vulco. El invierno queda cerca del límite "
            "y febrero, marzo y abril lo superan cerca de cuatro veces. "
            "Ese exceso, en período punta, paga tarifa de sobreconsumo."
        ),
        es_control=True,
    ),
    Cuenta(
        src_name="Almirante_Riveros_7621.pdf",
        archivo="weir Almirante Riveros",
        titulo="Almirante Riveros 1050",
        sitio="Vulco",
        direccion="Almirante Riveros 1050, San Bernardo",
        titular="Gómez González, Alejandro Javier",
        rut="No figura en la boleta (boleta, no factura)",
        cuenta="2713987-6",
        medidor="2.128",
        documento="Boleta electrónica N° 322122961",
        lectura="31-07-2026 al 31-08-2026 (sin lecturas de medidor)",
        consumo_m3=253,
        total_pagar="$ 351.198",
        clave="Consumo estimado · medidor cerrado",
        nota_extra=(
            "La boleta no trae lectura actual ni anterior. Informa consumo promedio "
            "descontable de 253 m³ y 12.144 m³ por abonar en la próxima facturación. "
            "Despacho: weirbills01@weir.urjanet.com."
        ),
        lectura_perfil=(
            "La serie del gráfico es plana: los 13 meses quedan en 253 m³. "
            "No es un consumo medido. Aguas Andinas facturó por estimación porque "
            "la clave de lectura es CERRADO. Hasta que haya una lectura real, "
            "este perfil no sirve para comparar meses ni para dimensionar un medidor WES "
            "contra la facturación."
        ),
    ),
    Cuenta(
        src_name="San_Jos__0895_850e.pdf",
        archivo="weir San José 0895",
        titulo="San José 0895",
        sitio="Matriz San José 895",
        direccion="San José 0895, San Bernardo",
        titular="Vulco S.A.",
        rut="91.619.000-K",
        cuenta="296813-4",
        medidor="123707699",
        documento="Factura electrónica N° 9455953",
        lectura="31-07-2026 al 31-08-2026",
        consumo_m3=4014,
        total_pagar="$ 5.466.709",
        clave="Consumo real · lectura normal",
        nota_extra=(
            "Arranque 50 mm. Indemnización por corte de agua potable el 01-09-2026 "
            "(2 h 20 min): −$ 91.255. Despacho: weirbills01@weir.urjanet.com."
        ),
        lectura_perfil=(
            "Es la cuenta de mayor volumen. Tras un septiembre 2025 más bajo "
            "(cerca de 2.600 m³), el resto del año se mueve entre unos 3.700 y 7.100 m³. "
            "El máximo está en marzo 2026. Los últimos seis meses se estabilizan "
            "alrededor de 3.700–4.400 m³. Septiembre 2026 facturó 4.014 m³ con lectura real."
        ),
    ),
    Cuenta(
        src_name="San_Jos__0815_b1f8.pdf",
        archivo="weir San José 0815",
        titulo="San José 0815",
        sitio="Medidor fundición",
        direccion="San José 0815, San Bernardo",
        titular="Vulco S.A.",
        rut="91.619.000-K",
        cuenta="296816-9",
        medidor="131163780",
        documento="Factura electrónica N° 9455955",
        lectura="30-05-2026 al 02-09-2026 (95 días)",
        consumo_m3=3646,
        total_pagar="$ 4.690.886",
        clave="Consumo real · lectura normal",
        nota_extra=(
            "Diferencia de lecturas 4.803 m³, abono de 1.157 m³ descontables, "
            "consumo facturado 3.646 m³. Indemnización por corte el 11-08-2026 "
            "(20 h 01 min): −$ 357.614. Arranque 40 mm. "
            "Despacho: hector.almendares@mail.weir."
        ),
        lectura_perfil=(
            "La barra de septiembre no es un mes calendario: la factura cubre del "
            "30-05-2026 al 02-09-2026 (95 días) y concentra 3.646 m³. "
            "Mayo y junio también aparecen altos (cerca de 3.100 y 2.300 m³). "
            "Agosto queda en torno a 280 m³. Conviene no usar septiembre como "
            "consumo mensual típico al comparar con las otras direcciones."
        ),
        dias_ultimo=95,
    ),
    Cuenta(
        src_name="San_Jos__0783_b9dd.pdf",
        archivo="weir San José 0783",
        titulo="San José 0783",
        sitio="Servicio San José 0783",
        direccion="San José 0783, San Bernardo",
        titular="Vulco S.A.",
        rut="91.619.000-K",
        cuenta="296817-7",
        medidor="122714081",
        documento="Factura electrónica N° 9457198",
        lectura="31-07-2026 al 31-08-2026",
        consumo_m3=1072,
        total_pagar="$ 1.485.313",
        clave="Consumo real · lectura normal",
        nota_extra=(
            "El membrete de la factura dice San José 0815 (domicilio de Vulco S.A.). "
            "En el cuerpo, Aguas Andinas imprime DIRECCIÓN: SAN JOSE 0783. "
            "Arranque 50 mm. Indemnización por corte el 11-08-2026 (4 h 55 min)."
        ),
        lectura_perfil=(
            "Consumo relativamente estable, entre unos 660 y 1.330 m³. "
            "El máximo es enero 2026 y el mínimo agosto 2026. "
            "Septiembre 2026 facturó 1.072 m³ con lectura real, en línea con el promedio del año."
        ),
    ),
]


def _boleta_path(cuenta: Cuenta) -> Path:
    local = BOLETAS / cuenta.src_name
    if local.is_file():
        return local
    return UPLOADS / cuenta.src_name


def medir_alturas(pdf: Path) -> list[int]:
    """Altura en píxeles de las 13 barras, de izquierda a derecha."""
    doc = pymupdf.open(pdf)
    page = doc[0]
    clip = pymupdf.Rect(18, 312, 172, 348)
    pix = page.get_pixmap(matrix=pymupdf.Matrix(6, 6), clip=clip, alpha=False)
    doc.close()
    arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    if arr.shape[2] == 4:
        arr = arr[:, :, :3]
    r = arr[:, :, 0].astype(int)
    g = arr[:, :, 1].astype(int)
    b = arr[:, :, 2].astype(int)
    bar = (np.abs(r - g) < 12) & (np.abs(r - b) < 12) & (r >= 75) & (r <= 120)
    col = bar.sum(axis=0)
    spans: list[tuple[int, int]] = []
    inicio = None
    for i, activo in enumerate(col > 6):
        if activo and inicio is None:
            inicio = i
        elif not activo and inicio is not None:
            if 12 <= i - inicio <= 80:
                spans.append((inicio, i - 1))
            inicio = None
    barras: list[tuple[int, int]] = []
    bottoms: list[int] = []
    for x0, x1 in spans:
        densidad = bar[:, x0 : x1 + 1].sum(axis=1)
        rows = np.where(densidad > (x1 - x0) * 0.4)[0]
        if len(rows) < 3:
            continue
        barras.append((x0, x1))
        bottoms.append(int(rows.max()))
    if len(barras) != 13:
        raise RuntimeError(f"{pdf.name}: se esperaban 13 barras y hay {len(barras)}")
    base = int(np.median(bottoms))
    corregidas = []
    for x0, x1 in barras:
        densidad = bar[:, x0 : x1 + 1].sum(axis=1)
        rows = np.where(densidad > (x1 - x0) * 0.4)[0]
        top = int(rows.min())
        corregidas.append(max(base - top, 1))
    return corregidas


def serie_m3(alturas: list[int], consumo_ultimo: int, plana: bool) -> list[int]:
    if plana or max(alturas) - min(alturas) <= 2:
        return [consumo_ultimo] * 13
    escala = consumo_ultimo / alturas[-1]
    valores = [int(round(h * escala / 10.0) * 10) for h in alturas[:-1]]
    valores.append(consumo_ultimo)
    return valores


def fmt_m3(n: int) -> str:
    return f"{n:,}".replace(",", ".")


def stats(valores: list[int]) -> dict:
    ultimos_12 = valores[1:]
    return {
        "promedio_12": int(round(sum(ultimos_12) / 12)),
        "suma_12": sum(ultimos_12),
        "max": max(valores),
        "min": min(valores),
        "mes_max": MESES[valores.index(max(valores))],
        "mes_min": MESES[valores.index(min(valores))],
    }


# Tarifa con IVA publicada en la boleta (lecturas desde el 11-05-2026), grupo AA Gran Santiago.
CARGO_FIJO = 944.0
AGUA_NO_PUNTA = 611.48
AGUA_PUNTA = 605.25
AGUA_SOBRE = 1737.17
RECOLECCION = 461.64
TRATAMIENTO = 311.29
# La barra de la boleta es el mes de facturación. En estas cuentas la lectura cae
# a fin del mes anterior: ene–may son las cinco lecturas del 31-dic al 30-abr (período punta).
PUNTA_IDX = {4, 5, 6, 7, 8}
# Lecturas 31-ago a 30-nov 2025 (barras sep a dic): parte del no punta que fija el límite.
LIMITE_IDX = (0, 1, 2, 3)
SOBRE_COLOR = "#C0392B"


@dataclass
class MesCosto:
    nombre: str
    m3: int
    punta: bool
    sobre: int
    costo_normal: int
    costo_con_sobre: int

    @property
    def diferencia(self) -> int:
        return self.costo_con_sobre - self.costo_normal

    @property
    def dentro(self) -> int:
        return self.m3 - self.sobre


def _pesos(valor: float) -> int:
    return int(round(valor))


def limite_sobreconsumo(valores: list[int]) -> int:
    base = [valores[i] for i in LIMITE_IDX]
    return max(30, int(round(sum(base) / len(base))))


def analizar(valores: list[int], dias_ultimo: int | None) -> tuple[int, list[MesCosto]]:
    limite = limite_sobreconsumo(valores)
    meses: list[MesCosto] = []
    for i, m3 in enumerate(valores):
        punta = i in PUNTA_IDX
        factor = 1.0
        if i == len(valores) - 1 and dias_ultimo:
            # Lectura del 02-09-2026: queda fuera de punta. El límite y el cargo fijo se prorratean.
            punta = False
            factor = dias_ultimo / 30.0
        tope = limite * factor
        sobre = max(0, m3 - tope) if punta else 0
        agua = AGUA_PUNTA if punta else AGUA_NO_PUNTA
        variable_normal = agua + RECOLECCION + TRATAMIENTO
        costo_normal = CARGO_FIJO * factor + m3 * variable_normal
        costo_con_sobre = costo_normal + sobre * (AGUA_SOBRE - agua)
        meses.append(
            MesCosto(
                nombre=MESES[i],
                m3=m3,
                punta=punta,
                sobre=int(round(sobre)),
                costo_normal=_pesos(costo_normal),
                costo_con_sobre=_pesos(costo_con_sobre),
            )
        )
    return limite, meses


def fmt_clp(n: int) -> str:
    signo = "-" if n < 0 else ""
    return signo + "$ " + f"{abs(n):,}".replace(",", ".")


def fmt_tarifa(n: float) -> str:
    entero = int(round(n * 100))
    pesos, cents = divmod(abs(entero), 100)
    return f"{pesos:,}".replace(",", ".") + f",{cents:02d}"


def grafico(meses: list[MesCosto], limite: int, titulo: str, destino: Path) -> None:
    nombres = [m.nombre for m in meses]
    dentro = [m.dentro for m in meses]
    sobre = [m.sobre for m in meses]
    fig, ax = plt.subplots(figsize=(10.2, 3.55), dpi=140)
    ax.bar(nombres, dentro, color=BAR, width=0.72, zorder=3, label="Hasta el límite, o mes no punta")
    ax.bar(
        nombres,
        sobre,
        bottom=dentro,
        color=SOBRE_COLOR,
        width=0.72,
        zorder=3,
        label="Sobreconsumo (período punta)",
    )
    ax.axhline(limite, color=SOBRE_COLOR, linestyle="--", linewidth=1.0, zorder=2, label=f"Límite {fmt_m3(limite)} m³")
    ax.set_ylabel("m³")
    ax.set_title(titulo, loc="left", fontsize=11, color="#0E3A5D", pad=8)
    ax.yaxis.grid(True, linestyle="--", linewidth=0.6, color="#D0D7DE", zorder=0)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.tick_params(axis="x", labelrotation=40, labelsize=8)
    ax.tick_params(axis="y", labelsize=8)
    ymax = max(m.m3 for m in meses) * 1.22
    ax.set_ylim(0, ymax)
    for i, m in enumerate(meses):
        ax.text(i, m.m3 + ymax * 0.015, fmt_m3(m.m3), ha="center", va="bottom", fontsize=6.5, color="#243140")
        if m.sobre > ymax * 0.08:
            ax.text(
                i,
                m.dentro + m.sobre / 2,
                fmt_m3(m.sobre),
                ha="center",
                va="center",
                fontsize=6.5,
                color="white",
                fontweight="bold",
            )
    ax.legend(loc="upper left", frameon=False, fontsize=7.5, ncol=2)
    fig.tight_layout()
    fig.savefig(destino, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def _wrap(c: canvas.Canvas, text: str, font: str, size: float, width: float) -> list[str]:
    words = text.split()
    lines: list[str] = []
    cur = ""
    for w in words:
        trial = w if not cur else cur + " " + w
        if c.stringWidth(trial, font, size) <= width:
            cur = trial
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


def _resumen_clp(meses: list[MesCosto], limite: int) -> str:
    punta = [m for m in meses if m.punta]
    dif = sum(m.diferencia for m in punta)
    normal = sum(m.costo_normal for m in punta)
    real = sum(m.costo_con_sobre for m in punta)
    sobre_m3 = sum(m.sobre for m in punta)
    recargo = AGUA_SOBRE - AGUA_PUNTA
    return (
        f"En los cinco meses punta el costo a tarifa normal sería {fmt_clp(normal)} "
        f"y con sobreconsumo queda en {fmt_clp(real)}. "
        f"La diferencia es {fmt_clp(dif)} "
        f"({fmt_m3(sobre_m3)} m³ sobre el límite de {fmt_m3(limite)} m³, "
        f"a ${fmt_tarifa(recargo)} más por m³ de agua potable)."
    )


def _header(c: canvas.Canvas, cuenta: Cuenta, subtitulo: str) -> None:
    w, h = A4
    c.setFillColorRGB(*NAVY)
    c.rect(0, h - 64, w, 64, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)
    c.setFont("DejaVuBold", 14)
    c.drawString(36, h - 30, f"Perfil de consumo · Weir · {cuenta.titulo}")
    c.setFont("DejaVu", 8.5)
    c.drawString(36, h - 48, subtitulo)
    c.setFont("DejaVu", 8)
    c.drawRightString(w - 36, h - 48, "Octubre 2026")


def _pie(c: canvas.Canvas) -> None:
    c.setFont("DejaVu", 8)
    c.setFillColorRGB(*NAVY)
    c.drawString(36, 22, "WES · Perfil de consumo Weir Minerals · San Bernardo")


def escribir_pdf(cuenta: Cuenta, meses: list[MesCosto], limite: int, png: Path, destino: Path) -> None:
    s = stats([m.m3 for m in meses])
    c = canvas.Canvas(str(destino), pagesize=A4)
    w, h = A4
    rol = "Medidor de control · " if cuenta.es_control else ""
    _header(c, cuenta, f"{rol}Vulco S.A. · San Bernardo · Aguas Andinas · boleta septiembre 2026")

    y = h - 84
    c.setFillColorRGB(*NAVY)
    c.setFont("DejaVuBold", 11)
    c.drawString(36, y, "Cuenta")
    y -= 14
    c.setFillColorRGB(0.12, 0.16, 0.2)
    filas = [
        ("Dirección", cuenta.direccion),
        ("Sitio", cuenta.sitio + (" · medidor de control" if cuenta.es_control else "")),
        ("Titular", f"{cuenta.titular} · RUT {cuenta.rut}"),
        ("N° de cuenta", cuenta.cuenta),
        ("Medidor", cuenta.medidor),
        ("Documento", cuenta.documento),
        ("Período de lectura", cuenta.lectura),
        ("Consumo facturado", f"{fmt_m3(cuenta.consumo_m3)} m³ · {cuenta.total_pagar}"),
        ("Clave", cuenta.clave),
        ("Límite de sobreconsumo", f"{fmt_m3(limite)} m³/mes"),
    ]
    for etiqueta, valor in filas:
        c.setFont("DejaVuBold", 8)
        c.drawString(36, y, etiqueta)
        c.setFont("DejaVu", 8)
        c.drawString(158, y, valor)
        y -= 11

    y -= 4
    c.setFillColorRGB(*NAVY)
    c.setFont("DejaVuBold", 11)
    c.drawString(36, y, "Últimos 13 meses")
    y -= 2
    img = ImageReader(str(png))
    iw, ih = Image.open(png).size
    draw_w = w - 72
    draw_h = draw_w * ih / iw
    y -= draw_h
    c.drawImage(img, 36, y, width=draw_w, height=draw_h, mask="auto")

    y -= 14
    c.setFillColorRGB(0.12, 0.16, 0.2)
    c.setFont("DejaVu", 8)
    resumen = (
        f"Últimos 12 meses (oct-25 a sep-26): {fmt_m3(s['suma_12'])} m³, "
        f"promedio {fmt_m3(s['promedio_12'])} m³/mes. "
        f"Máximo {fmt_m3(s['max'])} m³ ({s['mes_max']}). "
        f"Mínimo {fmt_m3(s['min'])} m³ ({s['mes_min']})."
    )
    for line in _wrap(c, resumen, "DejaVu", 8, w - 72):
        c.drawString(36, y, line)
        y -= 11
    y -= 2
    for line in _wrap(c, _resumen_clp(meses, limite), "DejaVu", 8, w - 72):
        c.drawString(36, y, line)
        y -= 11
    y -= 3
    c.setFont("DejaVuBold", 10)
    c.setFillColorRGB(*NAVY)
    c.drawString(36, y, "Cómo se lee")
    y -= 12
    c.setFillColorRGB(0.12, 0.16, 0.2)
    c.setFont("DejaVu", 8)
    for parrafo in (cuenta.lectura_perfil, cuenta.nota_extra):
        for line in _wrap(c, parrafo, "DejaVu", 8, w - 72):
            if y < 78:
                break
            c.drawString(36, y, line)
            y -= 10
        y -= 2
    c.setFont("DejaVu", 7.5)
    c.setFillColorRGB(0.35, 0.4, 0.45)
    c.drawString(36, 40, "El detalle en pesos, mes a mes, está en la página siguiente.")
    _pie(c)
    c.showPage()

    _header(c, cuenta, "Costo normal y costo con sobreconsumo · tarifa con IVA de la boleta")
    y = h - 84
    c.setFillColorRGB(0.12, 0.16, 0.2)
    c.setFont("DejaVu", 8)
    intro = (
        f"Límite de trabajo: {fmt_m3(limite)} m³/mes. Es el promedio de las barras sep-25 a dic-25 "
        "(lecturas de agosto a noviembre 2025). El decreto usa el promedio de todas las lecturas "
        "entre el 1 de mayo y el 30 de noviembre; en el gráfico faltan mayo, junio y julio 2025. "
        "El período punta son las lecturas del 1 de diciembre al 30 de abril: en este calendario, las barras ene-26 a may-26. "
        "Fuera de esos meses no hay sobreconsumo, aunque el consumo supere el límite. "
        "Costo normal: todo el m³ de agua a tarifa punta o no punta. "
        "Costo con sobreconsumo: solo el exceso de los meses punta pasa de "
        f"${fmt_tarifa(AGUA_PUNTA)} a ${fmt_tarifa(AGUA_SOBRE)} por m³. Recolección y tratamiento no cambian. "
        "Los pesos están revalorizados a la tarifa del 11-05-2026, no son el monto histórico de cada boleta."
    )
    for line in _wrap(c, intro, "DejaVu", 8, w - 72):
        c.drawString(36, y, line)
        y -= 10
    y -= 8

    cols = [36, 88, 148, 196, 258, 360, 468]
    headers = ["Mes", "Período", "m³", "m³ sobre", "Costo normal", "Con sobreconsumo", "Diferencia"]
    c.setFillColorRGB(*NAVY)
    c.rect(36, y - 4, w - 72, 16, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)
    c.setFont("DejaVuBold", 7.5)
    for x, texto in zip(cols, headers):
        c.drawString(x + 2, y, texto)
    y -= 16
    c.setFillColorRGB(0.12, 0.16, 0.2)
    for i, m in enumerate(meses):
        if i % 2 == 0:
            c.setFillColorRGB(0.94, 0.96, 0.98)
            c.rect(36, y - 3, w - 72, 13, fill=1, stroke=0)
        c.setFillColorRGB(0.12, 0.16, 0.2)
        c.setFont("DejaVu", 7.5)
        periodo = "Punta" if m.punta else "No punta"
        vals = [
            m.nombre,
            periodo,
            fmt_m3(m.m3),
            fmt_m3(m.sobre) if m.sobre else "—",
            fmt_clp(m.costo_normal),
            fmt_clp(m.costo_con_sobre),
            fmt_clp(m.diferencia) if m.diferencia else "—",
        ]
        for x, texto in zip(cols, vals):
            if texto == fmt_clp(m.diferencia) and m.diferencia:
                c.setFillColorRGB(0.75, 0.22, 0.17)
                c.setFont("DejaVuBold", 7.5)
            else:
                c.setFillColorRGB(0.12, 0.16, 0.2)
                c.setFont("DejaVu", 7.5)
            c.drawString(x + 2, y, texto)
        y -= 13

    y -= 6
    c.setFillColorRGB(*NAVY)
    c.setFont("DejaVuBold", 8)
    total_n = sum(m.costo_normal for m in meses)
    total_r = sum(m.costo_con_sobre for m in meses)
    total_d = sum(m.diferencia for m in meses)
    total_s = sum(m.sobre for m in meses)
    total_txt = (
        f"Total 13 meses: normal {fmt_clp(total_n)} · con sobreconsumo {fmt_clp(total_r)} · "
        f"diferencia {fmt_clp(total_d)} · {fmt_m3(total_s)} m³ en sobreconsumo."
    )
    for line in _wrap(c, total_txt, "DejaVuBold", 8, w - 72):
        c.drawString(36, y, line)
        y -= 11
    y -= 8

    if cuenta.es_control:
        y = _bloque_control(c, cuenta, meses, limite, y, w)

    c.setFont("DejaVu", 7.5)
    c.setFillColorRGB(0.35, 0.4, 0.45)
    nota = (
        "La barra roja es solo el m³ que, en período punta, supera el límite. "
        "La diferencia en pesos es únicamente el recargo de agua potable de ese exceso "
        f"(${fmt_tarifa(AGUA_SOBRE - AGUA_PUNTA)} por m³). No incluye el costo de recolección ni de tratamiento, "
        "que se pagan igual en la tarifa normal."
    )
    for line in _wrap(c, nota, "DejaVu", 7.5, w - 72):
        if y < 40:
            break
        c.drawString(36, y, line)
        y -= 10
    _pie(c)
    c.save()


def _bloque_control(
    c: canvas.Canvas,
    cuenta: Cuenta,
    meses: list[MesCosto],
    limite: int,
    y: float,
    w: float,
) -> float:
    punta = [m for m in meses if m.punta]
    sobre_m3 = sum(m.sobre for m in punta)
    dif = sum(m.diferencia for m in punta)
    # Costo completo de los m³ por encima del límite (agua a tarifa sobre + recolección + tratamiento).
    variable_sobre = AGUA_SOBRE + RECOLECCION + TRATAMIENTO
    costo_exceso = _pesos(sobre_m3 * variable_sobre)
    costo_exceso_normal = _pesos(sobre_m3 * (AGUA_PUNTA + RECOLECCION + TRATAMIENTO))
    c.setFillColorRGB(*NAVY)
    c.setFont("DejaVuBold", 11)
    c.drawString(36, y, "Revisión del medidor de control")
    y -= 13
    c.setFillColorRGB(0.12, 0.16, 0.2)
    c.setFont("DejaVu", 8)
    parrafos = [
        (
            f"{cuenta.titulo} es el medidor de control de Estadio Vulco, no la suma de San José. "
            "San José 0783, 0815 y 0895 y Almirante Riveros tienen cuenta y medidor propios."
        ),
        (
            f"El límite de {fmt_m3(limite)} m³ sale de sep-25 {fmt_m3(meses[0].m3)}, "
            f"oct-25 {fmt_m3(meses[1].m3)}, nov-25 {fmt_m3(meses[2].m3)} y dic-25 {fmt_m3(meses[3].m3)}. "
            "Si mayo, junio y julio 2025 fueron más altos o más bajos, el límite oficial se mueve y el recargo también."
        ),
        (
            "Febrero, marzo y abril quedan cerca de 3.870 m³, unas cuatro veces el límite. "
            f"En los cinco meses punta hay {fmt_m3(sobre_m3)} m³ por encima del límite."
        ),
        (
            f"Diferencia de tarifa (lo que pide la boleta de sobreconsumo respecto de la tarifa punta normal): "
            f"{fmt_clp(dif)}. "
            f"Si esos mismos m³ se hubieran cobrado a tarifa punta, costarían {fmt_clp(costo_exceso_normal)} "
            f"entre agua, recolección y tratamiento. A tarifa de sobreconsumo cuestan {fmt_clp(costo_exceso)}. "
            f"El recargo de agua es la diferencia entre ambos: {fmt_clp(dif)}."
        ),
        (
            "Mantener el control cerca del límite en diciembre–abril evita ese recargo. "
            "El salto no está en el invierno: septiembre 2026, con lectura real, son 397 m³, bajo el límite."
        ),
    ]
    for texto in parrafos:
        for line in _wrap(c, texto, "DejaVu", 8, w - 72):
            c.drawString(36, y, line)
            y -= 10
        y -= 3
    return y


def escribir_docx(cuenta: Cuenta, meses: list[MesCosto], limite: int, png: Path, destino: Path) -> None:
    s = stats([m.m3 for m in meses])
    doc = Document()
    estilo = doc.styles["Normal"]
    estilo.font.name = "Calibri"
    estilo.font.size = Pt(11)
    titulo = doc.add_heading(f"Perfil de consumo · Weir · {cuenta.titulo}", level=1)
    for run in titulo.runs:
        run.font.color.rgb = NAVY_HEX
    p = doc.add_paragraph("Vulco S.A. · San Bernardo · Aguas Andinas · boleta de septiembre 2026")
    p.runs[0].italic = True
    if cuenta.es_control:
        doc.add_paragraph("Medidor de control de Estadio Vulco (Baquedano 1215).")

    tabla = doc.add_table(rows=11, cols=2)
    tabla.style = "Table Grid"
    datos = [
        ("Dirección", cuenta.direccion),
        ("Sitio", cuenta.sitio),
        ("Titular", f"{cuenta.titular} · RUT {cuenta.rut}"),
        ("N° de cuenta", cuenta.cuenta),
        ("Medidor", cuenta.medidor),
        ("Documento", cuenta.documento),
        ("Período de lectura", cuenta.lectura),
        ("Consumo facturado", f"{fmt_m3(cuenta.consumo_m3)} m³"),
        ("Total a pagar", f"{cuenta.total_pagar} · vence 26-09-2026"),
        ("Clave", cuenta.clave),
        ("Límite de sobreconsumo", f"{fmt_m3(limite)} m³/mes"),
    ]
    for i, (k, v) in enumerate(datos):
        tabla.rows[i].cells[0].text = k
        tabla.rows[i].cells[1].text = v

    doc.add_paragraph()
    doc.add_picture(str(png), width=Inches(6.4))
    doc.add_paragraph(
        f"Últimos 12 meses (oct-25 a sep-26): {fmt_m3(s['suma_12'])} m³ · "
        f"promedio {fmt_m3(s['promedio_12'])} m³/mes · "
        f"máximo {fmt_m3(s['max'])} m³ ({s['mes_max']}) · "
        f"mínimo {fmt_m3(s['min'])} m³ ({s['mes_min']})."
    )
    doc.add_paragraph(_resumen_clp(meses, limite))

    doc.add_heading("Costo normal y costo con sobreconsumo", level=2)
    t2 = doc.add_table(rows=1 + len(meses) + 1, cols=7)
    t2.style = "Table Grid"
    headers = ["Mes", "Período", "m³", "m³ sobre", "Costo normal", "Con sobreconsumo", "Diferencia"]
    for i, htxt in enumerate(headers):
        t2.rows[0].cells[i].text = htxt
    for i, m in enumerate(meses, start=1):
        fila = [
            m.nombre,
            "Punta" if m.punta else "No punta",
            fmt_m3(m.m3),
            fmt_m3(m.sobre) if m.sobre else "—",
            fmt_clp(m.costo_normal),
            fmt_clp(m.costo_con_sobre),
            fmt_clp(m.diferencia) if m.diferencia else "—",
        ]
        for j, txt in enumerate(fila):
            t2.rows[i].cells[j].text = txt
    ultima = t2.rows[len(meses) + 1]
    ultima.cells[0].text = "Total"
    ultima.cells[3].text = fmt_m3(sum(m.sobre for m in meses))
    ultima.cells[4].text = fmt_clp(sum(m.costo_normal for m in meses))
    ultima.cells[5].text = fmt_clp(sum(m.costo_con_sobre for m in meses))
    ultima.cells[6].text = fmt_clp(sum(m.diferencia for m in meses))

    doc.add_heading("Cómo se lee", level=2)
    doc.add_paragraph(cuenta.lectura_perfil)
    doc.add_paragraph(cuenta.nota_extra)
    if cuenta.es_control:
        doc.add_heading("Revisión del medidor de control", level=2)
        punta = [m for m in meses if m.punta]
        sobre_m3 = sum(m.sobre for m in punta)
        dif = sum(m.diferencia for m in punta)
        doc.add_paragraph(
            "Baquedano 1215 no suma las cuentas de San José ni Almirante Riveros: "
            "cada una tiene su propio medidor en Aguas Andinas. "
            f"En los cinco meses punta hay {fmt_m3(sobre_m3)} m³ sobre el límite "
            f"y el recargo de tarifa es {fmt_clp(dif)}."
        )
    nota = doc.add_paragraph(
        "Los meses previos a septiembre 2026 se midieron en el gráfico de la boleta. "
        "El rojo de la barra es el sobreconsumo del período punta. "
        "La diferencia en pesos es solo el recargo de agua potable de ese exceso, "
        f"${fmt_tarifa(AGUA_SOBRE - AGUA_PUNTA)} por m³, valorizado a la tarifa del 11-05-2026."
    )
    for run in nota.runs:
        run.font.size = Pt(9)
        run.italic = True
    pie = doc.add_paragraph("WES · Perfil de consumo Weir Minerals · San Bernardo · octubre 2026")
    for run in pie.runs:
        run.font.color.rgb = NAVY_HEX
        run.font.size = Pt(9)
    doc.save(destino)


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    (OUT / "graficos").mkdir(exist_ok=True)
    for cuenta in CUENTAS:
        pdf_src = _boleta_path(cuenta)
        texto = "\n".join(page.get_text() for page in pymupdf.open(pdf_src))
        if cuenta.cuenta not in texto.replace(" ", ""):
            digitos = re.sub(r"\D", "", cuenta.cuenta)
            if digitos not in re.sub(r"\D", "", texto):
                raise RuntimeError(f"No está la cuenta {cuenta.cuenta} en {pdf_src.name}")
        alturas = medir_alturas(pdf_src)
        plana = "estimado" in cuenta.clave.lower()
        valores = serie_m3(alturas, cuenta.consumo_m3, plana)
        limite, meses = analizar(valores, cuenta.dias_ultimo)
        png = OUT / "graficos" / f"{cuenta.archivo}.png"
        grafico(meses, limite, f"{cuenta.titulo} · m³ por mes", png)
        pdf_out = OUT / f"{cuenta.archivo}.pdf"
        docx_out = OUT / f"{cuenta.archivo}.docx"
        escribir_pdf(cuenta, meses, limite, png, pdf_out)
        escribir_docx(cuenta, meses, limite, png, docx_out)
        print(cuenta.archivo, "límite", limite)
        for m in meses:
            if m.sobre or m.nombre in ("Ene-26", "May-26", "Sep-26"):
                print(
                    f"  {m.nombre} punta={m.punta} m3={m.m3} sobre={m.sobre} "
                    f"normal={m.costo_normal} real={m.costo_con_sobre} dif={m.diferencia}"
                )
        print("  recargo", sum(m.diferencia for m in meses))
        print(" ", pdf_out)


if __name__ == "__main__":
    main()

"""
Informe de una página de las alertas de consumo de Fundo Zapallar.

Mismo formato que la revisión 23-09 a 14-10: gráfico de Etapa N°5,
tarjetas y las reglas que se cumplieron. Agrega el aviso de punto en cero.
"""

from __future__ import annotations

import csv
import re
from datetime import date, timedelta
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Patch

CARPETA = Path("reports/Fundo_Zapallar/alerta_nocturna")
INICIO = date(2026, 10, 15)
FIN = date(2026, 10, 26)

NAVY = "#1B3A4B"
BLUE = "#2E7AC8"
RED = "#C0392B"
AMBER = "#D68910"
INK = "#1F2A30"
MUTED = "#5C6B73"
GRID = "#E6ECF0"
OK = "#D5DDE3"

_ETAPAS = re.compile(r"Etapas 1-4 ([0-9.,]+) m³")
_MAXIMO = re.compile(r"máximo ([0-9.,]+) m³/h")


def _fmt(valor: float) -> str:
    return f"{valor:.2f}".replace(".", ",")


def _num(texto: str) -> float:
    limpio = texto.strip()
    if "," in limpio:
        limpio = limpio.replace(".", "").replace(",", ".")
    return float(limpio)


def _color_volumen(valor: float | None) -> str:
    if valor is None or valor <= 1:
        return OK
    if valor > 10:
        return RED
    if valor > 5:
        return AMBER
    return BLUE


def _color_nivel(nivel: str) -> str:
    if nivel.startswith("Crítica"):
        return RED
    if nivel.startswith("Preventiva"):
        return AMBER
    if "cero" in nivel.lower():
        return NAVY
    if nivel.startswith("Sin consumo"):
        return MUTED
    return BLUE


def _leer() -> dict[date, list[dict]]:
    por_dia: dict[date, list[dict]] = {}
    cursor = INICIO
    while cursor <= FIN:
        path = CARPETA / f"alerta_fundo_zapallar_{cursor.strftime('%Y%m%d')}.csv"
        if path.is_file():
            with path.open(encoding="utf-8", newline="") as fh:
                por_dia[cursor] = list(csv.DictReader(fh))
        cursor += timedelta(days=1)
    return por_dia


def _fila(por_dia: dict[date, list[dict]], noche: date, regla: str) -> dict | None:
    for fila in por_dia.get(noche, []):
        if fila["regla"] == regla and date.fromisoformat(fila["fecha"]) == (
            noche if "noche" in regla or regla.startswith("cero") else noche - timedelta(days=1)
        ):
            return fila
    for fila in por_dia.get(noche, []):
        if fila["regla"] == regla:
            return fila
    return None


def _noches(por_dia: dict[date, list[dict]]) -> list[date]:
    return [dia for dia in sorted(por_dia) if INICIO <= dia <= FIN]


def _etiqueta(fila: dict) -> str:
    nivel = fila["nivel"]
    nombre = {"seguimiento": "Seguimiento", "preventiva": "Preventiva", "critica": "Crítica"}.get(nivel, nivel)
    if "se repite" in fila.get("detalle", "") and nivel == "preventiva":
        return "Preventiva, se repite"
    return nombre


def _maximo(fila: dict) -> str:
    hallazgo = _MAXIMO.search(fila.get("detalle", ""))
    if not hallazgo:
        return ""
    return f" · máx. {_fmt(_num(hallazgo.group(1)))}"


def construir() -> Path:
    por_dia = _leer()
    dias = _noches(por_dia)
    if not dias:
        raise FileNotFoundError("No hay CSV de alerta para el informe de consumo.")
    etapa5 = []
    for dia in dias:
        fila = _fila(por_dia, dia, "etapa5-noche")
        etapa5.append(None if fila is None else float(fila["valor"]))

    def reglas(regla: str) -> list[tuple[str, str, str]]:
        salida = []
        for dia in dias:
            fila = _fila(por_dia, dia, regla)
            if fila is None or fila["nivel"] == "normal":
                continue
            fecha_fila = date.fromisoformat(fila["fecha"])
            if regla.startswith("balance-") and not (INICIO <= fecha_fila < FIN):
                continue
            fecha = fecha_fila.strftime("%d-%m")
            valor = f"{_fmt(float(fila['valor']))} m³{_maximo(fila)}"
            salida.append((fecha, valor, _etiqueta(fila)))
        return salida

    ceros = []
    for dia in dias:
        fila = _fila(por_dia, dia, "balance-estanque-etapas")
        if fila is None:
            continue
        fecha_fila = date.fromisoformat(fila["fecha"])
        if not (INICIO <= fecha_fila < FIN):
            continue
        hallazgo = _ETAPAS.search(fila.get("detalle", ""))
        if hallazgo and _num(hallazgo.group(1)) == 0:
            ceros.append(date.fromisoformat(fila["fecha"]))

    alertas = 0
    preventiva_e5 = 0
    for dia in dias:
        for regla in ("etapa5-noche", "etapas14-noche", "matriz-noche", "balance-matriz-estanque", "balance-estanque-etapas"):
            fila = _fila(por_dia, dia, regla)
            if fila is None or fila["nivel"] == "normal":
                continue
            fecha_fila = date.fromisoformat(fila["fecha"])
            if regla.startswith("balance-") and not (INICIO <= fecha_fila < FIN):
                continue
            alertas += 1
            if regla == "etapa5-noche" and fila["nivel"] == "preventiva":
                preventiva_e5 += 1
    max_e5 = max(v for v in etapa5 if v is not None)
    dia_max_e5 = dias[etapa5.index(max_e5)].strftime("%d-%m")

    destino = CARPETA / f"Alerta_Fundo_Zapallar_{INICIO.strftime('%Y%m%d')}_{FIN.strftime('%Y%m%d')}.pdf"
    destino.parent.mkdir(parents=True, exist_ok=True)
    fig = plt.figure(figsize=(8.27, 11.69), facecolor="white")

    encabezado = fig.add_axes([0, 0.935, 1, 0.065])
    encabezado.set_xlim(0, 1)
    encabezado.set_ylim(0, 1)
    encabezado.axis("off")
    encabezado.add_patch(plt.Rectangle((0, 0), 1, 1, color=NAVY))
    encabezado.text(0.045, 0.62, "WES", color="white", fontsize=15, fontweight="bold", va="center")
    encabezado.text(0.12, 0.64, "Seguimiento técnico", color="#D6E4EE", fontsize=8, va="center")
    encabezado.text(0.955, 0.64, "Fundo Zapallar", color="white", fontsize=11, fontweight="bold", va="center", ha="right")
    encabezado.text(
        0.955,
        0.30,
        f"{INICIO.strftime('%d-%m-%Y')}  ·  {FIN.strftime('%d-%m-%Y')}",
        color="#D6E4EE",
        fontsize=8,
        va="center",
        ha="right",
    )

    fig.text(0.055, 0.908, "Alertas nocturnas y balances diarios", fontsize=15, color=NAVY, fontweight="bold")
    fig.text(
        0.055,
        0.888,
        "Tres niveles de aviso: seguimiento, preventiva y crítica. Ventana 00:00 a 06:00. La categoría crítica se reserva para\n"
        "consumos excepcionalmente altos. Si una alerta de seguimiento se repite al día siguiente, pasa a preventiva.\n"
        "Un punto que marca 0,00 m³ más de un día completo también se avisa. Estanque del Medio = Estanque Inferior (000027-02).",
        fontsize=7.4,
        color=INK,
        va="top",
        linespacing=1.35,
    )

    tarjetas = [
        (str(alertas), "alertas de consumo en el período"),
        (f"{preventiva_e5} / {len(dias)}", "noches con alerta preventiva · Etapa 5"),
        (f"{_fmt(max_e5)} m³", f"máximo nocturno Etapa N°5, {dia_max_e5}"),
        (str(len(ceros)), "días en cero · Etapas 1-4, desde el 17-10"),
    ]
    acentos = [BLUE, AMBER, RED, NAVY]
    for i, (grande, etiqueta) in enumerate(tarjetas):
        ax = fig.add_axes([0.055 + i * 0.232, 0.798, 0.215, 0.055])
        ax.set_xlim(0, 1)
        ax.set_ylim(0, 1)
        ax.axis("off")
        ax.add_patch(plt.Rectangle((0, 0), 1, 1, facecolor="#F4F7F9", edgecolor="#D5DEE5", lw=0.6))
        ax.add_patch(plt.Rectangle((0, 0), 0.025, 1, color=acentos[i]))
        ax.text(0.10, 0.62, grande, fontsize=12, color=NAVY, fontweight="bold", va="center")
        ax.text(0.10, 0.26, etiqueta, fontsize=6.2, color=MUTED, va="center")

    ax = fig.add_axes([0.055, 0.610, 0.62, 0.158])
    xs = list(range(len(dias)))
    colores = [_color_volumen(valor) for valor in etapa5]
    ax.bar(xs, [0 if v is None else v for v in etapa5], color=colores, width=0.72, zorder=3)
    for x, valor in zip(xs, etapa5):
        if valor is None:
            continue
        ax.text(x, valor + 0.25, _fmt(valor), ha="center", va="bottom", rotation=90, fontsize=6.2, color=INK, clip_on=False)
    ax.axhline(1, color=BLUE, lw=0.8, ls=":", zorder=2)
    ax.axhline(5, color=AMBER, lw=0.8, ls="--", zorder=2)
    ax.axhline(10, color=RED, lw=0.9, ls="-", zorder=2)
    ax.set_xlim(-0.6, len(dias) - 0.4)
    ax.set_ylim(0, 16.2)
    ax.set_ylabel("m³ por noche", fontsize=8, color=INK)
    ax.set_title("Etapa N°5 — m³ entre las 00:00 y las 06:00", loc="left", fontsize=9, color=NAVY, pad=8)
    ax.set_xticks(xs)
    ax.set_xticklabels([d.strftime("%d-%m") for d in dias], rotation=0, fontsize=6.0, color=MUTED)
    ax.tick_params(axis="y", labelsize=7, colors=MUTED)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#D5DEE5")
    ax.spines["bottom"].set_color("#D5DEE5")
    ax.yaxis.grid(True, color=GRID, zorder=0)
    ax.set_axisbelow(True)
    ax.legend(
        handles=[
            Patch(facecolor=BLUE, label="Seguimiento: >1 a 5 m³"),
            Patch(facecolor=AMBER, label="Preventiva: >5 a 10 m³"),
            Patch(facecolor=RED, label="Crítica: >10 m³ o >4 m³/h"),
            Patch(facecolor=OK, label="Sin alerta"),
        ],
        frameon=False,
        fontsize=6.4,
        loc="center left",
        bbox_to_anchor=(1.03, 0.55),
        borderaxespad=0,
        handlelength=1.1,
        handleheight=0.8,
        labelspacing=0.55,
    )

    fig.text(0.055, 0.575, "Otras reglas que se cumplieron", fontsize=11, color=NAVY, fontweight="bold")

    filas_cero = []
    for i, dia in enumerate(ceros):
        marca = "Primer día en cero" if i == 0 else "En cero, se avisa"
        filas_cero.append((dia.strftime("%d-%m"), "0,00 m³ día", marca))
    if ceros:
        filas_cero.append(("17-10 al 26-10", "0,00 m³ noche", "Sin consumo nocturno"))

    bloques = [
        (
            "Etapas 1-4  ·  en cero",
            "Se avisa si el día completo marca 0,00 m³ más de un día.\nEl 16-10 todavía tuvo consumo.",
            filas_cero,
        ),
        (
            "Matriz ESVAL  ·  noche",
            "Seguimiento >55 a 65 m³ o >25 a 30 m³/h. Preventiva >65 a 80 m³\no >30 a 35 m³/h. Crítica >80 m³ o >35 m³/h.",
            reglas("matriz-noche"),
        ),
        (
            "Matriz − Estanque Inferior  ·  día",
            "Seguimiento >21 a 25 m³. Preventiva >25 a 42 m³ o >21 m³ durante\n2 días. Crítica >42 m³.",
            reglas("balance-matriz-estanque"),
        ),
        (
            "Estanque − (Etapa N°5 + Etapas 1-4)  ·  día",
            "Seguimiento >30 a 36 m³. Preventiva >36 a 60 m³ o >30 m³ durante\n2 días. Crítica >60 m³.",
            reglas("balance-estanque-etapas"),
        ),
    ]
    cajas = [
        (0.055, 0.292, 0.44, 0.258),
        (0.515, 0.292, 0.43, 0.258),
        (0.055, 0.042, 0.44, 0.210),
        (0.515, 0.042, 0.43, 0.210),
    ]
    for (titulo, criterio, filas), (x0, y0, ancho, alto) in zip(bloques, cajas):
        axb = fig.add_axes([x0, y0, ancho, alto])
        axb.set_xlim(0, 1)
        axb.set_ylim(0, 1)
        axb.axis("off")
        axb.add_patch(plt.Rectangle((0, 0), 1, 1, facecolor="#F8FAFB", edgecolor="#E1E8ED", lw=0.6))
        axb.text(0.04, 0.90, titulo, fontsize=8, color=NAVY, fontweight="bold", va="center")
        axb.text(0.04, 0.76, criterio, fontsize=5.5, color=MUTED, va="top", linespacing=1.25)
        if not filas:
            axb.text(0.04, 0.40, "Sin alerta de consumo en el período.", fontsize=7, color=MUTED, va="center")
            continue
        paso = 0.62 / max(len(filas), 1)
        tam_fecha = 6.2 if len(filas) <= 8 else 5.7
        for i, (fecha, valor, nivel) in enumerate(filas):
            yy = 0.62 - i * paso
            color = _color_nivel(nivel)
            axb.plot([0.045], [yy], marker="o", color=color, markersize=4.2)
            axb.text(0.09, yy, f"{fecha}    {valor}", fontsize=tam_fecha, color=INK, va="center")
            axb.text(0.97, yy, nivel, fontsize=tam_fecha - 0.3, color=color, va="center", ha="right")

    fig.text(
        0.055,
        0.012,
        "WES SpA  ·  Documento de una página para revisión interna. Seguimiento o preventiva: contrastar con riegos, llenados u operaciones extraordinarias\n"
        "antes de escalar al fundo. Crítica: aviso prioritario. El 26-10 el día calendario aún no estaba cerrado; el balance de ese día no se incluye.",
        fontsize=6.1,
        color=MUTED,
        va="bottom",
    )
    fig.savefig(destino, format="pdf")
    plt.close(fig)
    return destino


if __name__ == "__main__":
    print(construir())

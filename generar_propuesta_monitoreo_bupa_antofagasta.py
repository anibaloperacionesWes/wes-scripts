"""
Propuesta de monitoreo hídrico — Clínica Bupa Antofagasta.

Arma el Word de alcance a partir de la ficha de visita
(fichas/Clinica_Bupa_Antofagasta/).

Uso:
  python generar_propuesta_monitoreo_bupa_antofagasta.py
"""

from __future__ import annotations

import tempfile
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor

ROOT = Path(__file__).resolve().parent
FICHA = (
    ROOT
    / "fichas"
    / "Clinica_Bupa_Antofagasta"
    / "Ficha_tecnica_Clinica_Bupa_Antofagasta.docx"
)
OUT_DIR = ROOT / "reports" / "Bupa_Antofagasta" / "PROPUESTA"
OUT_DOCX = OUT_DIR / "Propuesta_monitoreo_Clinica_Bupa_Antofagasta.docx"

AZUL = RGBColor(0, 51, 102)
NEGRO = RGBColor(0, 0, 0)

# Orden de fotos en la ficha (word/media).
FOTO_LOGO = "word/media/image5.png"
FOTOS = {
    "principal": "word/media/image1.png",
    "sala2": "word/media/image6.png",
    "sexto": "word/media/image2.png",
    "sanitaria": "word/media/image3.png",
}

PUNTOS = [
    {
        "n": "1",
        "clave": "principal",
        "nombre": "Sala de impulsión Principal",
        "actividad": "50 % de la clínica y llenado del estanque S.N°2",
        "electrica": "10 m",
        "foto_pie": (
            "Sala de bombas con manifold de cobre, válvulas y tramos verticales "
            "y horizontales accesibles para el par de sensores."
        ),
    },
    {
        "n": "2",
        "clave": "sala2",
        "nombre": "Sala de impulsión N°2",
        "actividad": "Abastece la sala de bombas del sexto piso (S.B. N°3)",
        "electrica": "10 m",
        "foto_pie": (
            "Cuarto de impulsión con tuberías aisladas. En la fotografía de la "
            "visita hay una nota manuscrita: «Sala de impulsión 1 no se monitorea»."
        ),
    },
    {
        "n": "3",
        "clave": "sexto",
        "nombre": "Sala de impulsión Sexto Piso (S.B. N°3)",
        "actividad": "50 % de la clínica",
        "electrica": "15 m",
        "foto_pie": (
            "Bomba del sexto piso sobre base de concreto, con aislación en la "
            "tubería. El piso de la sala estaba mojado el día de la visita."
        ),
    },
    {
        "n": "4",
        "clave": "sanitaria",
        "nombre": "Medidor principal Sanitaria",
        "actividad": "100 % de la clínica (cuenta de agua)",
        "electrica": "5 m",
        "foto_pie": (
            "Medidor mecánico de Aguas Antofagasta. El ultrasonido se instala "
            "en el tramo recto de cobre de 2 pulgadas que corre sobre el medidor."
        ),
    },
]


def _shade(cell, hex_color: str) -> None:
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tc_pr.append(shd)


def _set_run(run, *, size=11, bold=False, color=NEGRO, name="Calibri") -> None:
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    run.font.color.rgb = color


def _p(doc: Document, text: str, *, bold=False, size=11, center=False, space_after=8, color=NEGRO):
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.JUSTIFY
    para.paragraph_format.space_after = Pt(space_after)
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    run = para.add_run(text)
    _set_run(run, size=size, bold=bold, color=color)
    return para


def _heading(doc: Document, text: str) -> None:
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(12)
    para.paragraph_format.space_after = Pt(6)
    para.paragraph_format.keep_with_next = True
    run = para.add_run(text.upper())
    _set_run(run, size=13, bold=True, color=AZUL)


def _bullet(doc: Document, text: str) -> None:
    para = doc.add_paragraph(style="List Bullet")
    para.paragraph_format.space_after = Pt(2)
    para.clear()
    run = para.add_run(text)
    _set_run(run, size=11)


def _kv_table(doc: Document, rows: list[tuple[str, str]]) -> None:
    table = doc.add_table(rows=len(rows), cols=2)
    table.autofit = True
    for i, (k, v) in enumerate(rows):
        c0, c1 = table.rows[i].cells
        c0.text = ""
        c1.text = ""
        r0 = c0.paragraphs[0].add_run(k)
        r1 = c1.paragraphs[0].add_run(v)
        _set_run(r0, size=10, bold=True, color=AZUL)
        _set_run(r1, size=10)
        _shade(c0, "E8EEF6")
        if i % 2 == 1:
            _shade(c1, "F7F9FC")
    doc.add_paragraph().paragraph_format.space_after = Pt(4)


def _resumen_table(doc: Document) -> None:
    headers = ["Punto", "Actividad hídrica", "220 V", "Sensores", "Señal M2M"]
    data = [
        (
            p["nombre"],
            p["actividad"],
            p["electrica"],
            "2 × 5 m",
            "Buena",
        )
        for p in PUNTOS
    ]
    table = doc.add_table(rows=1 + len(data), cols=len(headers))
    table.style = "Table Grid"
    for j, h in enumerate(headers):
        cell = table.rows[0].cells[j]
        cell.text = ""
        run = cell.paragraphs[0].add_run(h)
        _set_run(run, size=9, bold=True, color=RGBColor(255, 255, 255))
        _shade(cell, "003366")
    for i, row in enumerate(data, start=1):
        for j, val in enumerate(row):
            cell = table.rows[i].cells[j]
            cell.text = ""
            run = cell.paragraphs[0].add_run(val)
            _set_run(run, size=8, bold=(j == 0))
            if i % 2 == 0:
                _shade(cell, "F2F6FC")
    doc.add_paragraph().paragraph_format.space_after = Pt(4)


def _foto(doc: Document, path: Path, width_in: float, pie: str) -> None:
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.paragraph_format.space_before = Pt(6)
    para.paragraph_format.space_after = Pt(2)
    run = para.add_run()
    run.add_picture(str(path), width=Inches(width_in))
    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.space_after = Pt(8)
    r = cap.add_run(pie)
    _set_run(r, size=9, color=RGBColor(80, 80, 80))


def _footer(section) -> None:
    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("WES  ·  Propuesta de monitoreo hídrico  ·  Clínica Bupa Antofagasta  ·  ")
    _set_run(r, size=8, color=RGBColor(90, 90, 90))
    run = p.add_run()
    _set_run(run, size=8, color=RGBColor(90, 90, 90))
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_end)


def _extraer_fotos(dest: Path) -> dict[str, Path]:
    dest.mkdir(parents=True, exist_ok=True)
    out: dict[str, Path] = {}
    with zipfile.ZipFile(FICHA) as zf:
        logo = dest / "logo_wes.png"
        logo.write_bytes(zf.read(FOTO_LOGO))
        out["logo"] = logo
        for clave, interno in FOTOS.items():
            path = dest / f"{clave}.png"
            path.write_bytes(zf.read(interno))
            out[clave] = path
    return out


def construir(fotos: dict[str, Path]) -> Document:
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)
    section.top_margin = Cm(1.6)
    section.bottom_margin = Cm(1.6)
    _footer(section)

    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)
    normal.font.color.rgb = NEGRO

    if fotos["logo"].exists():
        logo_p = doc.add_paragraph()
        logo_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        logo_p.add_run().add_picture(str(fotos["logo"]), width=Inches(1.55))

    _p(doc, "PROPUESTA DE MONITOREO HÍDRICO", bold=True, size=18, center=True, space_after=2, color=AZUL)
    _p(doc, "Clínica Bupa Antofagasta", bold=True, size=16, center=True, space_after=2, color=AZUL)
    _p(
        doc,
        "Alcance técnico según ficha de visita del 24-06-2026, 09:30 h",
        size=11,
        center=True,
        space_after=10,
    )

    _heading(doc, "1. Antecedentes de la visita")
    _kv_table(
        doc,
        [
            ("Cliente", "Clínica Bupa Antofagasta"),
            ("Fecha de la visita", "24-06-2026, 09:30 h"),
            ("Contacto en la visita", "Catherine Velasquez"),
            ("Proveedor de agua", "Aguas Antofagasta"),
            ("Puntos relevados", "4"),
            ("Matriz en los 4 puntos", "Cobre, 2 pulgadas"),
            ("Remarcador del cliente", "No existe en ninguno de los puntos"),
            ("Fuente", "Ficha técnica de visita (datos de terreno y fotografías)"),
        ],
    )
    _p(
        doc,
        "Este documento fija el alcance de medición y las condiciones de instalación "
        "registradas en la visita. La cotización comercial (valores, plazo y condiciones "
        "contractuales) se emite en un documento aparte.",
    )

    _heading(doc, "2. Objetivo")
    _p(
        doc,
        "Medir de forma continua el agua que entra a la clínica y el régimen de las "
        "tres salas de impulsión, para distinguir el consumo de la cuenta sanitaria "
        "del caudal que impulsa cada sala (incluido el llenado del estanque S.N°2) "
        "y detectar consumos fuera del horario de operación.",
    )

    _heading(doc, "3. Cómo se leen los cuatro puntos")
    _p(
        doc,
        "La ficha describe la actividad hídrica de cada sala. Con esa lectura, el "
        "esquema de la propuesta queda así:",
    )
    _bullet(doc, "Medidor principal Sanitaria: referencia de la cuenta. Cubre el 100 % de la clínica.")
    _bullet(
        doc,
        "Sala de impulsión Principal: impulsa el 50 % de la clínica y llena el estanque S.N°2.",
    )
    _bullet(
        doc,
        "Sala de impulsión N°2: abastece la sala de bombas del sexto piso (S.B. N°3).",
    )
    _bullet(doc, "Sala de impulsión Sexto Piso (S.B. N°3): impulsa el otro 50 % de la clínica.")
    _p(
        doc,
        "La Sala N°2 y la sala del sexto piso quedan en el mismo recorrido: la primera "
        "alimenta a la segunda. Cada punto se informa por separado. La cuenta sanitaria "
        "es la referencia del total de la clínica. Las salas describen impulsión y llenado "
        "de estanque, y su caudal se contrasta con esa cuenta punto por punto.",
    )

    _heading(doc, "4. Alcance propuesto")
    _p(
        doc,
        "Cuatro puntos de monitoreo ultrasónico no invasivo, uno por cada matriz de "
        "cobre de 2 pulgadas relevada. En cada punto la visita dejó las mismas condiciones "
        "de comunicación y de sensores:",
    )
    _bullet(doc, "Par de sensores de ultrasonido, con canalización de 2 × 5 m.")
    _bullet(doc, "Enlace M2M 3G / 4G / 5G. En los cuatro puntos la señal se registró como buena.")
    _bullet(doc, "Canalización de señal: no aplica. La antena queda en el punto, sin tendido adicional.")
    _bullet(doc, "Alimentación 220 V desde un punto cercano. La distancia figura en cada ficha.")
    _bullet(doc, "El cliente no tiene remarcador en estos puntos: la medición WES es la del tramo.")
    _p(doc, "Resumen de factibilidad (todos los puntos quedan habilitados para instalar):", space_after=4)
    _resumen_table(doc)

    _heading(doc, "5. Ficha de cada punto")
    for punto in PUNTOS:
        _p(doc, f"Punto {punto['n']}. {punto['nombre']}", bold=True, size=12, color=AZUL, space_after=4)
        _kv_table(
            doc,
            [
                ("Diámetro", "2 pulgadas"),
                ("Material de la matriz", "Cobre"),
                ("Actividad hídrica", punto["actividad"]),
                ("Remarcador del cliente", "No existe"),
                ("Factibilidad eléctrica 220 V", punto["electrica"]),
                ("Canalización de señal", "No aplica"),
                ("Canalización sensores ultrasonido", "2 × 5 m"),
                ("Señal M2M 3G, 4G, 5G", "Buena"),
            ],
        )
        _foto(doc, fotos[punto["clave"]], 4.6, punto["foto_pie"])

    _heading(doc, "6. Condiciones de instalación a cuidar en terreno")
    _bullet(
        doc,
        "Medidor Sanitaria: el sensor va en el tramo recto de cobre de 2 pulgadas visible "
        "sobre el medidor de Aguas Antofagasta, de modo que la lectura represente el 100 % de la cuenta.",
    )
    _bullet(
        doc,
        "Sala de impulsión N°2: cuarto estrecho, tuberías con aislación y piso con restos "
        "de material. Conviene dejar el tramo de cobre libre en los 2 × 5 m de canalización "
        "antes de fijar sensores.",
    )
    _bullet(
        doc,
        "Sexto piso: el día de la visita el piso de la sala estaba mojado. La alimentación "
        "de 15 m y la fijación del equipo se hacen con ese recinto seco y con el tablero 220 V confirmado.",
    )
    _bullet(
        doc,
        "Sala Principal: hay manifold, válvulas y tramos verticales. El par de sensores "
        "se instala en un tramo de cobre de 2 pulgadas que impulse el 50 % de la clínica "
        "y el llenado del estanque S.N°2.",
    )

    _heading(doc, "7. Confirmar antes de instalar")
    _p(
        doc,
        "En la fotografía de la Sala de impulsión N°2 la visita anotó a mano "
        "«Sala de impulsión 1 no se monitorea». La ficha, en cambio, incluye la Sala de "
        "impulsión Principal como punto 1. Hay que confirmar con Catherine Velasquez si "
        "esa nota deja fuera un equipo dentro de la sala o si modifica el alcance de "
        "cuatro puntos. Mientras esa confirmación no cambie la ficha, la propuesta "
        "mantiene los cuatro puntos.",
    )
    _p(
        doc,
        "En la ficha original el medidor de la cuenta figura como «Santinaria». En esta "
        "propuesta queda como Medidor principal Sanitaria, alineado con el proveedor "
        "Aguas Antofagasta.",
    )

    _heading(doc, "8. Qué queda habilitado con estos cuatro puntos")
    _bullet(doc, "Consumo del medidor de la cuenta (100 % de la clínica) y de cada sala de impulsión.")
    _bullet(doc, "Perfil horario y consumo en madrugada en cada punto.")
    _bullet(doc, "Régimen de llenado del estanque S.N°2, visto desde la Sala de impulsión Principal.")
    _bullet(
        doc,
        "Contraste entre la cuenta sanitaria y las impulsiones, tratando la Sala N°2 y "
        "el sexto piso como tramos del mismo recorrido.",
    )

    _heading(doc, "9. Referencia interna WES")
    _p(
        doc,
        "Cuando estos puntos se dan de alta en la plataforma, la correspondencia usada "
        "en Clínica Bupa Antofagasta (empresa 000029) es:",
    )
    _bullet(doc, "Sala de impulsión Principal → 000029-07 Sala de Bomba Principal.")
    _bullet(doc, "Sala de impulsión Sexto Piso (S.B. N°3) → 000029-08 Sala de Bomba Sexto Piso.")
    _bullet(doc, "Medidor principal Sanitaria → 000029-09 Medidor Principal Sanitaria.")
    _bullet(doc, "Sala de impulsión N°2 → 000029-10 Sala de Bomba N°2.")
    _p(
        doc,
        "Esa correspondencia es de uso interno del equipo. La propuesta hacia el cliente "
        "se presenta con los nombres de sala de la ficha de visita.",
        size=10,
    )
    return doc


def main() -> None:
    if not FICHA.is_file():
        raise SystemExit(f"No está la ficha de visita: {FICHA}")
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        fotos = _extraer_fotos(Path(tmp))
        doc = construir(fotos)
    doc.save(OUT_DOCX)
    print(f"[OK] {OUT_DOCX}")


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""La comuna del formulario coincide con el catálogo de clientes y sitios."""

from __future__ import annotations

import json
import unittest
from pathlib import Path

from comunas_clientes import cargar_comunas, comuna_de

ROOT = Path(__file__).resolve().parent
SITIOS = ROOT / "catalogos" / "clientes_maquinas.json"
HTMLS = [
    ROOT / "apps_script_export" / "Formulario.html",
    ROOT / "formulario_visita.html",
]
GS = ROOT / "apps_script_export" / "Codigo.gs"


def _embed(path: Path) -> dict:
    text = path.read_text(encoding="utf-8")
    marca = "COMUNAS_WES = "
    inicio = text.find(marca)
    if inicio < 0:
        raise AssertionError(f"sin mapa embebido en {path.name}")
    inicio += len(marca)
    fin = text.find(";\n", inicio)
    if fin < 0:
        raise AssertionError(f"mapa sin cierre en {path.name}")
    return json.loads(text[inicio:fin])


class ComunasClientesTest(unittest.TestCase):
    def test_cada_sitio_del_catalogo_tiene_comuna(self):
        sitios = json.loads(SITIOS.read_text(encoding="utf-8"))
        vacios = []
        for cliente, maquinas in sitios.items():
            for maquina in maquinas:
                if not comuna_de(cliente, maquina):
                    vacios.append(f"{cliente} / {maquina}")
        self.assertEqual(vacios, [])

    def test_casos_conocidos(self):
        self.assertEqual(comuna_de("CORMUP", "TOBALABA"), "Peñalolén")
        self.assertEqual(comuna_de("COR. PUENTE", "LICEO MAIPO"), "Puente Alto")
        self.assertEqual(comuna_de("RENCA", "GIMNASIO MUNICIPAL"), "Renca")
        self.assertEqual(comuna_de("GENCHI", ""), "")
        self.assertEqual(comuna_de("GENCHI", "CDP PUENTE ALTO"), "Puente Alto")
        self.assertEqual(comuna_de("GENCHI", "COLINA 2 RED SUR"), "Colina")
        self.assertEqual(comuna_de("GENCHI", "CPF SAN MIGUEL MATRIZ PRINCIPAL"), "San Miguel")
        self.assertEqual(comuna_de("GENCHI", "ESFORPEN BOSQUE ADMINISTRACION"), "El Bosque")
        self.assertEqual(comuna_de("GENCHI", "CPA OVALO"), "Santiago")
        self.assertEqual(comuna_de("AGUNSA", "DEPOSITO"), "Lampa")
        self.assertEqual(comuna_de("AGUNSA", "INTERMODAL"), "San Antonio")
        self.assertEqual(comuna_de("UDD", "SALA DE BOMBAS HONDURAS"), "Las Condes")
        self.assertEqual(comuna_de("PAK", "DL KENNEDY"), "Las Condes")
        self.assertEqual(comuna_de("MAE", "PIZZA HUT"), "Estación Central")
        self.assertEqual(comuna_de("BOM", "SAN IGNACIO 500"), "Quilicura")
        self.assertEqual(comuna_de("CUR", "ANILLO NORTE"), "Valparaíso")
        self.assertEqual(comuna_de("DERCO", "QUILICURA - CASINO"), "Quilicura")
        self.assertEqual(comuna_de("DERCO", "LO BOZA - MATRIZ"), "Lampa")
        self.assertEqual(comuna_de("BUPA ANTOFAGASTA", "MEDIDOR"), "Antofagasta")
        self.assertEqual(comuna_de("COPEC", "COSTANERA"), "Vitacura")
        self.assertEqual(comuna_de(""), "")

    def test_html_y_apps_script_embeben_el_mismo_mapa(self):
        esperado = cargar_comunas()
        for path in HTMLS + [GS]:
            self.assertEqual(_embed(path), esperado, path.name)


if __name__ == "__main__":
    unittest.main()

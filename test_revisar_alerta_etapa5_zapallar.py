"""Primera alerta: Etapa N°5 no debe pasar de 60 m³/h."""

import unittest
from datetime import date, datetime
from zoneinfo import ZoneInfo

from revisar_alerta_etapa5_zapallar import (
    armar_correo,
    excesos,
    horas_de_la_revision,
    revision_para,
)


ZONA = ZoneInfo("America/Santiago")
DIA = date(2026, 9, 28)


class VentanasTests(unittest.TestCase):
    def test_08_mira_desde_las_23(self):
        self.assertEqual(
            horas_de_la_revision(DIA, "08:00"),
            [(date(2026, 9, 27), 23), *[(DIA, hora) for hora in range(0, 8)]],
        )

    def test_16_mira_la_manana_tarde(self):
        self.assertEqual(horas_de_la_revision(DIA, "16:00"), [(DIA, hora) for hora in range(8, 16)])

    def test_23_mira_hasta_las_22(self):
        self.assertEqual(horas_de_la_revision(DIA, "23:00"), [(DIA, hora) for hora in range(16, 23)])

    def test_programado_solo_en_esas_horas(self):
        self.assertIsNone(revision_para(datetime(2026, 9, 28, 10, 0, tzinfo=ZONA), True))
        self.assertEqual(
            revision_para(datetime(2026, 9, 28, 16, 5, tzinfo=ZONA), True),
            (DIA, "16:00"),
        )


class UmbralTests(unittest.TestCase):
    def test_60_no_avisa_y_mas_de_60_si(self):
        pedidas = [(DIA, 14), (DIA, 15)]
        lecturas = {(DIA, 14): 60.0, (DIA, 15): 60.1}
        self.assertEqual(excesos(lecturas, pedidas, 60), [(DIA, 15, 60.1)])

    def test_correo_nombra_las_horas(self):
        alerta = {
            "punto": "Etapa N°5",
            "cliente": "Fundo Zapallar",
            "nodeId": "000027-03",
            "umbral_m3h": 60,
            "motivo": "Una tubería de 3 pulgadas entrega como máximo 60 m³/h.",
        }
        asunto, cuerpo = armar_correo(alerta, DIA, "16:00", [(DIA, 14, 72.5)])
        self.assertIn("de 14:00 a 15:00", asunto)
        self.assertIn("de 14:00 a 15:00: 72,500 m³/h", cuerpo)
        self.assertIn("Revisión de las 16:00", cuerpo)
        self.assertIn("Aníbal y Juan", cuerpo)


if __name__ == "__main__":
    unittest.main()

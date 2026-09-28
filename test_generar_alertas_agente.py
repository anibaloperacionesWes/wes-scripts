"""Reglas propias del agente: umbral, horario y catálogo por cliente."""

import tempfile
import unittest
from datetime import date
from pathlib import Path

from generar_alertas_agente import (
    Hallazgo,
    config_por_defecto,
    evaluar_punto,
    evaluar_serie,
    horas_desde_csv,
    regla_para,
    sembrar_clientes,
    escribir_salida,
)


def _regla(**cambios):
    config = config_por_defecto()
    config.update(cambios)
    return regla_para(config, "000031", "000031-01")


class EvaluarSerieTests(unittest.TestCase):
    def test_noche_sobre_umbral(self):
        horas = {h: 0.0 for h in range(24)}
        horas[3] = 0.8
        hallazgos = evaluar_serie(horas, _regla())
        self.assertEqual([h[0] for h in hallazgos], ["consumo_fuera_de_horario"])
        self.assertIn("03:00", hallazgos[0][1])
        self.assertEqual(hallazgos[0][2], 0.8)

    def test_noche_bajo_umbral_no_alerta(self):
        horas = {h: 0.1 for h in range(24)}
        self.assertEqual(evaluar_serie(horas, _regla()), [])

    def test_umbral_del_cliente(self):
        config = config_por_defecto()
        config["clientes"] = {"000031": {"umbral_m3h": 2.0}}
        horas = {h: 0.0 for h in range(24)}
        horas[1] = 1.2
        self.assertEqual(evaluar_serie(horas, regla_para(config, "000031", "000031-01")), [])

    def test_umbral_del_punto(self):
        config = config_por_defecto()
        config["clientes"] = {
            "000031": {"umbral_m3h": 2.0, "puntos": {"000031-01": {"umbral_m3h": 0.3}}}
        }
        horas = {h: 0.0 for h in range(24)}
        horas[1] = 0.4
        nombres = [h[0] for h in evaluar_serie(horas, regla_para(config, "000031", "000031-01"))]
        self.assertEqual(nombres, ["consumo_fuera_de_horario"])

    def test_dia_en_cero(self):
        horas = {h: 0.0 for h in range(24)}
        nombres = [h[0] for h in evaluar_serie(horas, _regla())]
        self.assertEqual(nombres, ["dia_sin_consumo"])

    def test_sin_datos(self):
        self.assertEqual(evaluar_serie(None, _regla())[0][0], "sin_datos")
        self.assertEqual(evaluar_serie({1: 0.0, 2: 0.0}, _regla())[0][0], "sin_datos")

    def test_regla_apagada(self):
        config = config_por_defecto()
        config["reglas"]["sin_datos"] = False
        self.assertEqual(evaluar_serie(None, regla_para(config, "000031", "000031-01")), [])


class CatalogoTests(unittest.TestCase):
    def test_siembra_conserva_umbral_editado(self):
        config = config_por_defecto()
        config["clientes"] = {"000031": {"nombre": "Viejo", "activa": True, "umbral_m3h": 1.5}}
        empresas = [
            {
                "companyId": "000031",
                "name": "Club Providencia",
                "nodes": [{"nodeId": "000031-01", "name": "Piscina"}],
            },
            {
                "companyId": "000000",
                "name": "Wes Spa",
                "nodes": [{"nodeId": "000000-01", "name": "Interno"}],
            },
        ]
        nuevo = sembrar_clientes(config, empresas)
        self.assertEqual(list(nuevo["clientes"]), ["000031"])
        self.assertEqual(nuevo["clientes"]["000031"]["nombre"], "Club Providencia")
        self.assertEqual(nuevo["clientes"]["000031"]["umbral_m3h"], 1.5)

    def test_escribe_un_csv_por_cliente(self):
        hallazgo = Hallazgo(
            fecha="2026-09-27",
            company_id="000031",
            cliente="Club Providencia",
            node_id="000031-01",
            punto="Piscina",
            regla="consumo_fuera_de_horario",
            detalle="03:00 (0.800 m³/h)",
            valor_m3h=0.8,
            umbral_m3h=0.5,
        )
        with tempfile.TemporaryDirectory() as tmp:
            carpeta = escribir_salida([hallazgo], date(2026, 9, 27), Path(tmp))
            texto = (carpeta / "000031.csv").read_text(encoding="utf-8")
            resumen = (carpeta / "resumen.txt").read_text(encoding="utf-8")
        self.assertIn("000031-01", texto)
        self.assertIn("Club Providencia", resumen)


class CsvTests(unittest.TestCase):
    def test_hora_de_la_marca_y_ultimo_valor(self):
        csv = (
            "TIME,VALUE\n"
            "2026-09-27T03:00:00.000Z,0.200\n"
            "2026-09-27T03:00:00.000Z,0.800\n"
            "2026-09-26T03:00:00.000Z,9.000\n"
        )
        horas = horas_desde_csv(csv, date(2026, 9, 27))
        self.assertEqual(horas, {3: 0.8})


class PuntoTests(unittest.TestCase):
    def test_evaluar_punto_copia_identidad(self):
        horas = {h: 0.0 for h in range(24)}
        horas[4] = 1.0
        filas = evaluar_punto(
            horas=horas,
            regla=_regla(),
            fecha=date(2026, 9, 27),
            company_id="000031",
            cliente="Club Providencia",
            node_id="000031-01",
            punto="Piscina",
        )
        self.assertEqual(filas[0].cliente, "Club Providencia")
        self.assertEqual(filas[0].regla, "consumo_fuera_de_horario")


if __name__ == "__main__":
    unittest.main()

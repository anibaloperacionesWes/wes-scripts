"""Pruebas de resolución de empresa/punto y del plan de alta de alertas."""

import io
import unittest
from contextlib import redirect_stdout

from configurar_alerta import (
    AlertaError,
    buscar_empresa,
    buscar_puntos,
    cmd_crear,
    normalizar,
    parse_umbral,
)


EMPRESAS = [
    {
        "companyId": "000025",
        "name": "Parque Arauco",
        "nodes": [
            {"nodeId": "000025-01", "name": "Estanque Norte Locales Mall", "configuration": {"threshold": "30"}},
            {"nodeId": "000025-02", "name": "Estanque Sur", "configuration": {"threshold": None}},
        ],
    },
    {"companyId": "000016", "name": "Renca", "nodes": []},
    {"companyId": "000010", "name": "Corporación Puente Alto", "nodes": []},
]


class ResolucionTests(unittest.TestCase):
    def test_normaliza_acentos_y_espacios(self):
        self.assertEqual(normalizar("  Corporación   Puente "), normalizar("corporacion puente"))

    def test_empresa_por_id_nombre_o_parcial(self):
        self.assertEqual(buscar_empresa(EMPRESAS, "000025")["name"], "Parque Arauco")
        self.assertEqual(buscar_empresa(EMPRESAS, "parque arauco")["companyId"], "000025")
        self.assertEqual(buscar_empresa(EMPRESAS, "puente alto")["companyId"], "000010")

    def test_punto_ambiguo(self):
        nodos = EMPRESAS[0]["nodes"]
        self.assertEqual(buscar_puntos(nodos, "000025-02")[0]["name"], "Estanque Sur")
        with self.assertRaises(AlertaError):
            buscar_puntos(nodos, "estanque")

    def test_umbral_con_coma(self):
        self.assertEqual(parse_umbral("2,5"), "2.5")
        self.assertEqual(parse_umbral("30"), "30")
        with self.assertRaises(AlertaError):
            parse_umbral("-1")


class _FalsoCliente:
    def __init__(self):
        self.creadas = []
        self.correos = []
        self.umbrales = []

    def listar_empresas(self):
        return EMPRESAS

    def obtener_alerta(self, company_id, node_id, tipo):
        if node_id == "000025-01":
            return {"venueName": "Estanque Norte", "receiverList": [{"email": "ya@wes.cl"}]}
        return None

    def crear_alerta(self, company_id, node_id, tipo, correos):
        self.creadas.append((company_id, node_id, tipo, correos))
        return {}

    def agregar_correos(self, company_id, node_id, tipo, correos):
        self.correos.append((node_id, correos))
        return {}

    def definir_umbral(self, company_id, node_id, umbral):
        self.umbrales.append((node_id, umbral))
        return {}


class PlanCrearTests(unittest.TestCase):
    def _args(self, confirmar: bool):
        return type(
            "Args",
            (),
            {
                "empresa": "Parque Arauco",
                "punto": "Estanque Sur",
                "todos_los_puntos": False,
                "tipo": "filtración",
                "correo": ["nuevo@cliente.cl, ya@wes.cl"],
                "umbral": "12,5",
                "confirmar": confirmar,
            },
        )()

    def test_sin_confirmar_no_escribe(self):
        cliente = _FalsoCliente()
        with redirect_stdout(io.StringIO()):
            codigo = cmd_crear(cliente, self._args(False))
        self.assertEqual(codigo, 0)
        self.assertEqual(cliente.creadas, [])
        self.assertEqual(cliente.umbrales, [])

    def test_confirmar_crea_y_guarda_umbral(self):
        cliente = _FalsoCliente()
        with redirect_stdout(io.StringIO()):
            codigo = cmd_crear(cliente, self._args(True))
        self.assertEqual(codigo, 0)
        self.assertEqual(cliente.creadas, [("000025", "000025-02", "FILTRATION", ["nuevo@cliente.cl", "ya@wes.cl"])])
        self.assertEqual(cliente.umbrales, [("000025-02", "12.5")])

    def test_si_ya_existe_solo_suma_correos_nuevos(self):
        cliente = _FalsoCliente()
        args = self._args(True)
        args.punto = "000025-01"
        args.umbral = "30"
        with redirect_stdout(io.StringIO()):
            cmd_crear(cliente, args)
        self.assertEqual(cliente.creadas, [])
        self.assertEqual(cliente.correos, [("000025-01", ["nuevo@cliente.cl"])])
        self.assertEqual(cliente.umbrales, [])


if __name__ == "__main__":
    unittest.main()

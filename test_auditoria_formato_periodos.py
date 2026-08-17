"""Pruebas del formato de auditoría: etiquetas de leyenda y nombre de Excel."""

from datetime import date

from generar_auditoria_wes_cliente import (
    dias_inclusive,
    periodo_desde_rango,
    xlsx_nombre_consolidado,
)
from generar_graficos_comparativos_desde_excel_consolidado import (
    _partir_dos_periodos,
    etiqueta_rango_wes,
)


def test_etiqueta_misma_mes():
    assert (
        etiqueta_rango_wes("Con WES", date(2026, 8, 10), date(2026, 8, 16))
        == "Con WES 10–16 ago 2026"
    )
    assert (
        etiqueta_rango_wes("Sin WES", date(2026, 4, 6), date(2026, 4, 12))
        == "Sin WES 06–12 abr 2026"
    )


def test_partir_14_usa_fechas_del_excel_no_marzo_fijo():
    """Las 14 columnas de abril eran Con 13–19 / Sin 06–12; no 23–29 mar."""
    fechas = [date(2026, 4, d) for d in list(range(13, 20)) + list(range(6, 13))]
    mats = [[0.0] * 24 for _ in fechas]
    (lab1, v1), (lab2, v2) = _partir_dos_periodos(fechas, mats)
    assert lab1 == "Con WES 13–19 abr 2026"
    assert lab2 == "Sin WES 06–12 abr 2026"
    assert len(v1) == 7 and len(v2) == 7


def test_partir_agosto_7_mas_7():
    fechas = [date(2026, 8, d) for d in list(range(10, 17)) + list(range(17, 24))]
    mats = [[0.0] * 24 for _ in fechas]
    (lab1, _), (lab2, _) = _partir_dos_periodos(fechas, mats)
    assert lab1 == "Con WES 10–16 ago 2026"
    assert lab2 == "Sin WES 17–23 ago 2026"


def test_xlsx_nombre_agosto():
    ref = periodo_desde_rango("con", date(2026, 8, 10), date(2026, 8, 16))
    aud = periodo_desde_rango("sin", date(2026, 8, 17), date(2026, 8, 23))
    assert xlsx_nombre_consolidado(ref, aud) == (
        "consumo_consolidado_parseo_filas_con_ago10-16_sin_ago17-23_2026.xlsx"
    )


def test_dias_inclusive_lunes_domingo():
    dias = dias_inclusive(date(2026, 8, 17), date(2026, 8, 23))
    assert len(dias) == 7
    assert dias[0].weekday() == 0  # lunes
    assert dias[-1].weekday() == 6  # domingo


if __name__ == "__main__":
    test_etiqueta_misma_mes()
    test_partir_14_usa_fechas_del_excel_no_marzo_fijo()
    test_partir_agosto_7_mas_7()
    test_xlsx_nombre_agosto()
    test_dias_inclusive_lunes_domingo()
    print("ok")

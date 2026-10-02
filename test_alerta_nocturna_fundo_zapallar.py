"""Criterio de la alerta nocturna de Fundo Zapallar, sin llamar a la API."""

from datetime import date, datetime
from zoneinfo import ZoneInfo

from alerta_nocturna_fundo_zapallar import (
    CHILE_TZ,
    consumo_ventana,
    exceso_porcentaje,
    noches_base,
    promedio_noches,
    supera_umbral,
    ultima_noche_completa,
)


def test_ventana_suma_00_a_06():
    horas = {h: float(h) for h in range(24)}
    assert consumo_ventana(horas) == float(sum(range(0, 7)))


def test_promedio_desde_el_23_de_septiembre():
    noche = date(2026, 10, 2)
    dias = noches_base(noche, 90)
    assert dias[0] == date(2026, 9, 23)
    assert dias[-1] == date(2026, 10, 1)
    assert date(2026, 7, 2) not in dias
    assert noche not in dias
    assert len(dias) == 9

    # Con más de 90 noches desde el 23-09, el tope sigue siendo 90.
    tarde = date(2026, 12, 31)
    dias_tarde = noches_base(tarde, 90)
    assert len(dias_tarde) == 90
    assert dias_tarde[0] == date(2026, 10, 2)
    assert dias_tarde[-1] == date(2026, 12, 30)

    assert noches_base(date(2026, 9, 23), 90) == []
    assert noches_base(date(2026, 9, 24), 90) == [date(2026, 9, 23)]
    assert promedio_noches([10.0, 20.0, 30.0]) == 20.0


def test_supera_el_promedio_en_25_porciento():
    assert supera_umbral(125.01, 100.0, 1.25)
    assert not supera_umbral(125.0, 100.0, 1.25)
    assert not supera_umbral(100.0, 100.0, 1.25)
    assert supera_umbral(0.01, 0.0, 1.25)
    assert not supera_umbral(0.0, 0.0, 1.25)
    assert exceso_porcentaje(150.0, 100.0) == 50.0
    assert exceso_porcentaje(10.0, 0.0) is None


def test_noche_completa_despues_de_las_07():
    tarde = datetime(2026, 9, 30, 18, 14, tzinfo=CHILE_TZ)
    assert ultima_noche_completa(tarde) == date(2026, 9, 30)
    madrugada = datetime(2026, 10, 1, 5, 30, tzinfo=CHILE_TZ)
    assert ultima_noche_completa(madrugada) == date(2026, 9, 30)
    cerrada = datetime(2026, 10, 1, 7, 0, tzinfo=ZoneInfo("America/Santiago"))
    assert ultima_noche_completa(cerrada) == date(2026, 10, 1)

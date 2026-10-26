"""Reglas de alerta de Fundo Zapallar: seguimiento, preventiva y crítica."""

from datetime import date, datetime
from pathlib import Path

from alerta_nocturna_fundo_zapallar import (
    CHILE_TZ,
    aplicar_persistencia,
    consumo_ventana,
    maximo_hora,
    nivel_balance,
    nivel_etapa5,
    nivel_etapas_14,
    nivel_matriz,
    racha_flujo,
    ultima_noche_completa,
)


def test_ventana_00_a_06():
    horas = {h: 1.0 for h in range(24)}
    assert consumo_ventana(horas) == 7.0
    assert maximo_hora({0: 1.0, 4: 2.4, 8: 9.0}) == 2.4
    assert racha_flujo({0: 0.1, 1: 0.1, 2: 0.0, 3: 1.0, 4: 1.0, 5: 1.0, 6: 1.0}) == 4


def test_etapa5_bandas_del_pdf():
    assert nivel_etapa5(1.0, 1.0) == "normal"
    assert nivel_etapa5(3.2, 3.2) == "seguimiento"
    assert nivel_etapa5(5.0, 3.7) == "seguimiento"
    assert nivel_etapa5(5.01, 1.0) == "preventiva"
    assert nivel_etapa5(9.0, 1.3) == "preventiva"
    assert nivel_etapa5(10.0, 3.5) == "preventiva"
    assert nivel_etapa5(10.30, 3.5) == "critica"
    assert nivel_etapa5(3.0, 4.0) == "seguimiento"
    assert nivel_etapa5(3.0, 4.01) == "critica"
    assert nivel_etapa5(3.2, 1.1, "seguimiento") == "preventiva"
    assert nivel_etapa5(3.2, 1.1, "preventiva") == "preventiva"
    assert nivel_etapa5(9.0, 1.3, "seguimiento") == "preventiva"
    assert nivel_etapa5(0.4, 0.2, "seguimiento") == "normal"


def test_etapas_14_y_matriz():
    assert nivel_etapas_14(40.0, 10.0) == "normal"
    assert nivel_etapas_14(40.36, 10.0) == "seguimiento"
    assert nivel_etapas_14(45.0, 10.0) == "seguimiento"
    assert nivel_etapas_14(45.1, 10.0) == "preventiva"
    assert nivel_etapas_14(47.05, 10.0) == "preventiva"
    assert nivel_etapas_14(60.0, 10.0) == "preventiva"
    assert nivel_etapas_14(60.1, 10.0) == "critica"
    assert nivel_etapas_14(39.13, 12.24) == "seguimiento"
    assert nivel_etapas_14(37.02, 13.08) == "seguimiento"
    assert nivel_etapas_14(33.37, 12.49, "preventiva") == "preventiva"
    assert nivel_etapas_14(44.72, 10.0, "seguimiento") == "preventiva"
    assert nivel_etapas_14(10.0, 12.0) == "normal"
    assert nivel_etapas_14(10.0, 15.0) == "seguimiento"
    assert nivel_etapas_14(10.0, 15.1) == "preventiva"
    assert nivel_etapas_14(10.0, 20.0) == "preventiva"
    assert nivel_etapas_14(10.0, 20.1) == "critica"

    assert nivel_matriz(55.0, 10.0) == "normal"
    assert nivel_matriz(58.20, 10.0) == "seguimiento"
    assert nivel_matriz(65.0, 10.0) == "seguimiento"
    assert nivel_matriz(70.82, 10.0) == "preventiva"
    assert nivel_matriz(80.0, 10.0) == "preventiva"
    assert nivel_matriz(80.1, 10.0) == "critica"
    assert nivel_matriz(60.83, 10.0, "preventiva") == "preventiva"
    assert nivel_matriz(57.30, 10.0) == "seguimiento"
    assert nivel_matriz(10.0, 25.0) == "normal"
    assert nivel_matriz(10.0, 30.0) == "seguimiento"
    assert nivel_matriz(10.0, 30.1) == "preventiva"
    assert nivel_matriz(10.0, 35.0) == "preventiva"
    assert nivel_matriz(10.0, 35.1) == "critica"


def test_balance_diario():
    assert nivel_balance(21.0, 21, 25, 42) == "normal"
    assert nivel_balance(23.02, 21, 25, 42) == "seguimiento"
    assert nivel_balance(23.02, 21, 25, 42, previo_abs=28.66) == "preventiva"
    assert nivel_balance(25.0, 21, 25, 42) == "seguimiento"
    assert nivel_balance(25.96, 21, 25, 42) == "preventiva"
    assert nivel_balance(36.43, 21, 25, 42) == "preventiva"
    assert nivel_balance(42.0, 21, 25, 42) == "preventiva"
    assert nivel_balance(42.01, 21, 25, 42) == "critica"
    assert nivel_balance(23.31, 21, 25, 42, previo_abs=10.0) == "seguimiento"

    assert nivel_balance(30.0, 30, 36, 60) == "normal"
    assert nivel_balance(31.92, 30, 36, 60) == "seguimiento"
    assert nivel_balance(35.35, 30, 36, 60) == "seguimiento"
    assert nivel_balance(32.88, 30, 36, 60, previo_abs=40.28) == "preventiva"
    assert nivel_balance(41.12, 30, 36, 60) == "preventiva"
    assert nivel_balance(60.0, 30, 36, 60) == "preventiva"
    assert nivel_balance(60.01, 30, 36, 60) == "critica"


def test_persistencia_no_sube_un_dia_normal():
    assert aplicar_persistencia("normal", "preventiva") == "normal"
    assert aplicar_persistencia("seguimiento", None) == "seguimiento"
    assert aplicar_persistencia("seguimiento", "seguimiento") == "preventiva"
    assert aplicar_persistencia("preventiva", "preventiva") == "preventiva"
    assert aplicar_persistencia("critica", "seguimiento") == "critica"


def test_pdf_de_un_dia_marca_la_critica():
    import pypdfium2 as pdfium

    from generar_pdf_alerta_fundo_zapallar import pagina_dia

    filas = [
        {
            "fecha": "2026-10-23",
            "periodo": "23-10-2026 00:00 a 06:00",
            "regla": "etapa5-noche",
            "punto": "Etapa N°5 — noche",
            "nivel": "preventiva",
            "valor": "3.5",
            "umbral": "Seguimiento: >1 a 5 m³/noche. Preventiva: >5 a 10 m³/noche. Crítica: >10 m³/noche o >4 m³/h.",
            "detalle": "Total 3,50 m³; máximo 3,50 m³/h; racha con flujo 1 h. 00:00=0,00; 06:00=3,50. La alerta de seguimiento se repite, por eso el nivel pasa a preventiva.",
            "accion": "Revisar riego abierto o pérdida.",
        },
        {
            "fecha": "2026-10-22",
            "periodo": "22-10-2026 día completo",
            "regla": "balance-estanque-etapas",
            "punto": "Estanque Inferior − (Etapa N°5 + Etapas 1-4)",
            "nivel": "critica",
            "valor": "102.5",
            "umbral": "Seguimiento >30 a 36 m³/día. Preventiva >36 a 60 m³/día o >30 m³ por 2 días. Crítica >60 m³/día.",
            "detalle": "Estanque 116,40 m³; Etapa N°5 13,90 m³; Etapas 1-4 0,00 m³; ∆ 102,50 m³.",
            "accion": "Sobre 30 m³ debe revisarse el balance de los dos estanques de etapa.",
        },
    ]
    fig = pagina_dia(filas)
    destino = "/tmp/alerta_zapallar_test.pdf"
    fig.savefig(destino, format="pdf")
    texto = pdfium.PdfDocument(destino)[0].get_textpage().get_text_bounded()
    assert "Alerta crítica" in texto
    assert "102,50 m³" in texto
    assert "3,50 m³" in texto
    assert "se repite" in texto


def test_el_correo_adjunta_el_pdf():
    from alerta_nocturna_fundo_zapallar import _mensaje_correo

    pdf = Path("/tmp/alerta_zapallar_adjunto.pdf")
    pdf.write_bytes(b"%PDF-1.4\n%alerta\n")
    filas = [
        {
            "fecha": "2026-10-26",
            "periodo": "26-10-2026 00:00 a 06:00",
            "regla": "etapa5-noche",
            "punto": "Etapa N°5 — noche",
            "nivel": "preventiva",
            "valor": 3.4,
            "umbral": "Seguimiento: >1 a 5 m³/noche.",
            "detalle": "Total 3,40 m³.",
            "accion": "Revisar riego abierto o pérdida.",
        }
    ]
    msg = _mensaje_correo(filas, pdf)
    partes = msg.get_payload()
    assert any(parte.get_content_type() == "application/pdf" for parte in partes)
    assert msg.get_payload()[1].get_filename() == "alerta_zapallar_adjunto.pdf"
    assert "Se adjunta el PDF" in msg.get_payload()[0].get_payload(decode=True).decode("utf-8")


def test_noche_completa_despues_de_las_07():
    assert ultima_noche_completa(datetime(2026, 10, 14, 18, 0, tzinfo=CHILE_TZ)) == date(2026, 10, 14)
    assert ultima_noche_completa(datetime(2026, 10, 15, 5, 0, tzinfo=CHILE_TZ)) == date(2026, 10, 14)

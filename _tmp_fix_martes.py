from pathlib import Path

p = Path("generar_informe_reparacion_fugas_honduras.py")
t = p.read_text(encoding="utf-8")

repls = [
    ("LUNES_ANTES", "MARTES_ANTES"),
    ("LUNES_DESPUES", "MARTES_DESPUES"),
    ("horas_lun_antes", "horas_mar_antes"),
    ("horas_lun_desp", "horas_mar_desp"),
    ("lunes_desp_sin_datos", "martes_desp_sin_datos"),
    ("total_lun_antes", "total_mar_antes"),
    ("total_lun_desp", "total_mar_desp"),
    ("noct_lun_antes", "noct_mar_antes"),
    ("noct_lun_desp", "noct_mar_desp"),
    ("img_lun_antes", "img_mar_antes"),
    ("img_lun_desp", "img_mar_desp"),
    ("img_lun_comp", "img_mar_comp"),
    ("_grafico_comparativo_lunes_lineas", "_grafico_comparativo_martes_lineas"),
    ("03_lunes_antes_estilo_app.png", "03_martes_antes_estilo_app.png"),
    ("04_lunes_despues_estilo_app.png", "04_martes_despues_estilo_app.png"),
    ("fallback si el lunes después está vacío", "fallback si el martes después está vacío"),
    ("lunes antes vs día después", "martes antes vs martes después"),
    ('label=f"Lunes {dia_antes.strftime', 'label=f"Martes {dia_antes.strftime'),
    ("p. ej. lunes antes vs lunes después", "p. ej. martes antes vs martes después"),
    ("Comparativo lunes (estilo app línea)", "Comparativo martes (estilo app línea)"),
    ("Cargando perfiles horarios lunes", "Cargando perfiles horarios martes"),
    ("Si el lunes después no tiene CSV", "Si el martes después no tiene CSV"),
    ("[AVISO] Lunes {MARTES_DESPUES}", "[AVISO] Martes {MARTES_DESPUES}"),
    ("primer día con datos post-reparación; lunes", "primer día con datos post-reparación; martes"),
    ("para el lunes {MARTES_DESPUES", "para el martes {MARTES_DESPUES}"),
    ("Comparativo de lunes (gráfica día estilo app WES)", "Comparativo de martes (gráfica día estilo app WES)"),
    ("los dos lunes de las ventanas", "los dos martes de las ventanas"),
    ('("Indicador", f"Lunes {MARTES_ANTES', '("Indicador", f"Martes {MARTES_ANTES'),
    ('f"Lunes {MARTES_DESPUES.strftime', 'f"Martes {MARTES_DESPUES.strftime'),
    ("Perfil día — lunes antes", "Perfil día — martes antes"),
    ("Perfil día — lunes después", "Perfil día — martes después"),
    ("rojo = lunes antes", "rojo = martes antes"),
    ('f"El lunes {MARTES_DESPUES', 'f"El martes {MARTES_DESPUES'),
    ("El lunes {MARTES_DESPUES.strftime", "El martes {MARTES_DESPUES.strftime"),
    ('f"Lunes {MARTES_ANTES}:', 'f"Martes {MARTES_ANTES}:'),
    ('f"Lunes {MARTES_DESPUES}:', 'f"Martes {MARTES_DESPUES}:'),
    ("etiqueta_perfil_desp = f\"Lunes", "etiqueta_perfil_desp = f\"Martes"),
    ("etiqueta_perfil_desp = f'Lunes", "etiqueta_perfil_desp = f'Martes"),
]

for a, b in repls:
    t = t.replace(a, b)

# catch remaining lowercase lunes referring to the comparison day
t = t.replace("el lunes después", "el martes después")
t = t.replace("El lunes ", "El martes ")
t = t.replace("el lunes ", "el martes ")

p.write_text(t, encoding="utf-8")
print("LUNES_", t.count("LUNES_"))
print("MARTES_", t.count("MARTES_"))
print("lunes remaining contexts:")
for i, line in enumerate(t.splitlines(), 1):
    if "lunes" in line.lower() and "DIAS_ES" not in line and '"Lunes"' not in line:
        if "lunes" in line.lower():
            print(i, line.strip()[:120])

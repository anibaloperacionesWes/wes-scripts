# Formulario técnico WES (vistoso) · PDF al cliente · Excel de análisis

## Acceso para técnicos (usar este link)
Planilla permanente — una fila por visita, listas desplegables:

https://docs.google.com/spreadsheets/d/1B5gGXua055WO5V9Ff4Tm-ur4msN4fA5WG1XEiKs-RSE/edit?usp=sharing

Compartida con `@wes.cl` y con el link (pueden escribir). Carpeta:

https://drive.google.com/drive/folders/1RCtWP1hK4fKzjgjyvzzSbttWJZiNhtKC

```bash
python3 mantenimiento_wes/compartir_acceso_tecnicos.py
```

## Formulario web con firma (Apps Script)
```bash
python3 mantenimiento_wes/publicar_formulario_permanente.py
```
Eso crea/actualiza el Apps Script y la carpeta Drive `Tecnicos_WES_Formulario`.

**Activación (1 vez):** abrí el Script → Implementar → Aplicación web → acceso **Cualquiera** → copiá la URL `/exec`.

Ese `/exec` es el link fijo del celular (firma + PDF + correo + fila en Excel).

## Flujo local / demo (túnel temporal)
```bash
cd mantenimiento_wes
python3 servir_formulario_visita.py
```
Celular (misma WiFi): `http://IP-DEL-PC:8787`

Al completar: PDF → correo al cliente (acusar recibo) → fila en Excel `Datos`.

## Clientes / máquinas que faltan
1. Agregá los clientes/máquinas en el Sheet **Registro de fallas WES** (hoja `Base1`)
   o en tu `analisis de falla.xlsx`.
2. Refrescá catálogos del formulario:
```bash
python sincronizar_catalogos_desde_maestro.py --desde-drive
```
3. Reiniciá `servir_formulario_visita.py`.

El sync también **completa solo** pares que ya existen en el historial `Datos`
pero faltaban en `Base1` (ej. MOLYMET, PAE).

## Mejorar el Excel digital
```bash
python mejorar_excel_digital.py
```
Agrega hoja **Resumen**, columnas extra (email, recibido por, PDF/Drive, Año, Mes)
y deja compatible la hoja `Datos` con el Sheet.

## Dónde se guarda cada visita
1. `FORMULARIO_MANTENCION_WES_DIGITAL.xlsx` → hoja `Datos`
2. Copia local `maestro/analisis_de_falla.xlsx` y `maestro/analisis_falla_google.xlsx`
3. Subida a Drive `mantenimiento wes/maestro/` (fallback si Sheets API no está habilitada)

Sheet vivo (catálogo oficial):
https://docs.google.com/spreadsheets/d/1GlRn7QXWEre7ziau29ojR5lTl-bZ8T3mCT3cD93HZgM/edit

## Correo
Secret: `WES_SMTP_PASSWORD` (app password de `agente.ia@wes.cl`).

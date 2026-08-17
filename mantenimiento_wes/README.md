# Formulario técnico WES (vistoso) · PDF al cliente · Excel de análisis

## Flujo en terreno
```bash
cd mantenimiento_wes
python servir_formulario_visita.py
```
Celular (misma WiFi): `http://IP-DEL-PC:8787`

Al completar: PDF → correo al cliente (acusar recibo) → fila en Excel `Datos`.

## Clientes / máquinas (puntos) — dónde editarlos
Fuente oficial del desplegable del formulario:

**`CONTACTOS_ENVIOS_ACTAS` → hoja `Clientes_catalogo`**
(no Base1 del Registro de fallas: ahí quedan nombres viejos, ej. RENCA).

```bash
python sincronizar_catalogos_desde_maestro.py --desde-drive
```

Reiniciá `servir_formulario_visita.py` después del sync.

- Puntos: `Clientes_catalogo`
- Emails: hoja `Contactos` del mismo Excel
- Fallas: Registro de fallas · `Base3`

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

## Contactos / emails (fuente ÚNICA)

Archivo oficial: **`CONTACTOS_ENVIOS_ACTAS`**

- Windows: `G:\Mi unidad\Agente WES\wes-scripts\mantenimiento wes\CONTACTOS_ENVIOS_ACTAS`
- Drive: https://docs.google.com/spreadsheets/d/1Tpjm1eXRXKuKvxachtbYVr9503wICJdsYDTjkbm__o8/edit

Columnas: `Cliente | Máquina | Rol | Nombre | Cargo | Email | Actualizado`

| Rol | Uso |
|---|---|
| `general` | TO (puede haber varios: JO, Linkes, etc.) |
| `CC` / `punto` | CC automático (Máquina vacía = todo el cliente) |

No edites correos en `FORMULARIO_MANTENCION_WES_DIGITAL` ni en otras planillas.

```bash
python contactos_cliente.py --desde-drive
```

Reiniciá el servidor del formulario después del sync. Aníbal queda siempre en CC adicional.

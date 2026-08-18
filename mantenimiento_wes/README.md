# Formulario técnico WES (vistoso) · PDF al cliente · Excel de análisis

## Arquitectura (2 Excel)

1. **`MAESTRO_FORMULARIO_WES`** — lo que el formulario necesita  
   Puntos · Contactos · Técnicos · Fallas · Opciones  
   Path: `mantenimiento wes\MAESTRO_FORMULARIO_WES.xlsx`

2. **`Registro de fallas WES`** — historial + evaluaciones  
   Hoja `Datos` = cada visita · ahí se analizan KPIs/pareto  
   https://docs.google.com/spreadsheets/d/1GlRn7QXWEre7ziau29ojR5lTl-bZ8T3mCT3cD93HZgM/edit

Detalle: [`ARQUITECTURA_DATOS.md`](ARQUITECTURA_DATOS.md)

## Flujo en terreno
```bash
cd mantenimiento_wes
python servir_formulario_visita.py
```
Celular (misma WiFi): `http://IP-DEL-PC:8787`

Al completar: PDF → correo al cliente (acusar recibo) → fila en Excel `Datos`.

## Actualizar catálogos del formulario
Editá el **maestro**, no el Registro:
```bash
python sincronizar_desde_maestro_formulario.py --desde-drive
```
Reiniciá el servidor (o, en el `/exec`, usá el Codigo que lee en vivo).

## Contactos / emails
En el maestro, hoja **Contactos**:

| Rol | Uso |
|---|---|
| `general` | TO |
| `CC` / `punto` | CC (Máquina vacía = todo el cliente) |

Aníbal queda siempre en CC adicional.

## Mejorar el Excel digital local
```bash
python mejorar_excel_digital.py
```

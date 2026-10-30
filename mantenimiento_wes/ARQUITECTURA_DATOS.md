# Arquitectura de datos — Formulario WES

Dos Excel. No más.

## 1) Configuración del formulario → `MAESTRO_FORMULARIO_WES`

Path: `G:\Mi unidad\Agente WES\wes-scripts\mantenimiento wes\MAESTRO_FORMULARIO_WES.xlsx`

| Hoja | Contenido |
|------|-----------|
| **Puntos** | Cliente + Máquina/sitio |
| **Contactos** | Quién recibe el PDF (`general`=TO, `CC`/`punto`=CC) |
| **Tecnicos** | Lista de técnicos |
| **Fallas** | Árbol Tipo → Falla específica |
| **Opciones** | Tipos mtto, motivos, tecnologías, checklist |

Acá **solo** se edita lo que el técnico elige en el form y a quién se manda el acta.

```bash
python sincronizar_desde_maestro_formulario.py --desde-drive
```

## 2) Historial + evaluaciones → `Registro de fallas WES`

https://docs.google.com/spreadsheets/d/1GlRn7QXWEre7ziau29ojR5lTl-bZ8T3mCT3cD93HZgM/edit

| Hoja | Uso |
|------|-----|
| **Datos** | Cada visita queda acá (fila nueva) |
| Resúmenes / Pareto / KPIs | Evaluaciones sobre `Datos` |

Acá **no** se editan puntos ni emails del form. Solo visitas y análisis.

## Flujo

```
MAESTRO_FORMULARIO_WES ──sync──► formulario (/exec o local)
         │
         └─ Contactos, Puntos, Técnicos, Fallas

Técnico completa visita ──► PDF + email
                         └─► Registro de fallas · Datos
                                   └─► evaluaciones / reportes
```

## Legacy

`CONTACTOS_ENVIOS_ACTAS` sigue válido como origen de contactos/puntos hasta
que el maestro esté adoptado; el generador lo consolidó dentro del maestro.

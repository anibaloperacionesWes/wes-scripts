# Formulario técnico WES (vistoso) · PDF al cliente · Excel de análisis

Reemplaza el Excel “amarillo” como canal de terreno. El técnico completa un
**formulario web en el teléfono**, firma en pantalla, y al enviar:

1. Se genera el **PDF del acta**
2. Se **envía al correo del cliente** pidiendo **acusar recibo**
3. Se agrega la fila al **Excel** `FORMULARIO_MANTENCION_WES_DIGITAL.xlsx` (hoja `Datos`) para seguir analizando

Sin Google Form.

## Cómo usarlo (en terreno)

En el PC (misma WiFi que el teléfono):

```bash
cd mantenimiento_wes
python servir_formulario_visita.py
```

Abrí en el celular la URL que imprime, por ejemplo:

`http://192.168.x.x:8787`

Completá cliente/máquina, checklist CIR·CPA·SAB, firma y correo del cliente → **Completar · PDF · Enviar**.

## Correo (SMTP)

Configurá en el entorno (o secretos del Cloud Agent):

- `WES_SMTP_PASSWORD` o `SMTP_PASSWORD` → contraseña de aplicación de `agente.ia@wes.cl`
- Opcional: `WES_SMTP_USUARIO` (default `agente.ia@wes.cl`)

Si falta la contraseña, igual se generan PDF + Excel; el correo se omite con aviso.

## Archivos clave

| Archivo | Rol |
|---|---|
| `formulario_visita.html` | UI móvil |
| `servir_formulario_visita.py` | Servidor + orquestación |
| `generar_pdf_acta_visita.py` | PDF acta |
| `enviar_acta_cliente_pdf.py` | Mail con pedido de acuse |
| `registrar_visita_excel.py` | Volcado a hoja `Datos` |
| `FORMULARIO_MANTENCION_WES_DIGITAL.xlsx` | Maestro local de análisis |
| `catalogos/*.json` | Listas Cliente→Máquina y fallas |
| `firma_visita.html` + `servir_firma.py` | Solo firma (legado) |

## Demo sin servidor

```bash
cd mantenimiento_wes
python demo_visita_completa.py
```

Genera un PDF de ejemplo en `salidas/` y `reports/Mantenimientos/formulario_visita/`.

## Drive

Las actas suben a `reports/Mantenimientos/actas_visita` cuando hay secretos `GOOGLE_DRIVE_*`.
También:

```bash
python ../subir_archivo_drive.py reports/Mantenimientos/formulario_visita/Acta_....pdf --subcarpeta "Mantenimientos/actas_visita"
```

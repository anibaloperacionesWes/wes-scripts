# Automation Cursor — Correo WES 8–18 (aprobación Aníbal)

Nombre sugerido en Cursor: **Correo 8-18**  
(si preferís “9 a 7”, renombralo; el horario operativo es **lun–vie 08:00–18:00 Chile**).

## Qué hace

1. Cada ~30 min (lun–vie), un Cloud Agent revisa `agente.ia@wes.cl`.
2. Si hay correos que necesitan respuesta, arma un **borrador** y te lo manda **solo a vos** (`anibal.aoperaciones@wes.cl`).
3. El cliente **no** recibe nada hasta que respondás al aviso con:

```text
DALE NOMÁS <id>
```

4. Opcional: debajo de esa línea pegá el texto editado; si no, se envía la respuesta sugerida.
5. Cuando se envía, te llega un `[ENVIADO]` de confirmación.

## Paso 1 — Secreto de Gmail (obligatorio)

Hoy en Cloud Agents solo están los secretos de Drive. Falta el de correo:

1. Abrí https://cursor.com/dashboard → **Cloud Agents** → **Secrets**
2. Agregá:

```text
WES_GMAIL_APP_PASSWORD=<contraseña de aplicación de agente.ia@wes.cl>
```

Misma contraseña de aplicación que usan los scripts de envío (16 caracteres, sin espacios).

## Paso 2 — Crear la Automation

1. Abrí https://cursor.com/automations
2. **New automation**
3. **Name:** `Correo 8-18`
4. **Trigger:** Schedule  
   - Recurring every **30 minutes**  
   - **Monday–Friday** only  
   - (El script igual se corta solo fuera de 08:00–18:00 Chile; podés dejar el cron todo el día lun–vie.)
5. **Repo / environment:** `wes-scripts` (el mismo que puntos en cero).
6. **Tools:** permitir terminal / shell. Slack opcional (avisar “hay N borradores”).
7. Pegá este **prompt**:

```text
Sos el agente de correo WES (cuenta agente.ia@wes.cl).

Objetivo: revisar el inbox en horario laboral Chile y preparar respuestas con aprobación humana de Aníbal. NUNCA envíes respuestas al cliente externo sin aprobación.

Pasos:
1. Ejecutá:
   python revisar_correo_wes_borradores.py

2. Si el script dice que está fuera de ventana laboral, terminá sin cambios.

3. Si faltan secretos (WES_GMAIL_APP_PASSWORD), avisá por Slack/resumen y no inventes credenciales.

4. Revisá el resumen del script (borradores_nuevos / enviados_aprobados).
   - Si hubo borradores nuevos: listá IDs y remitentes en el mensaje final.
   - Si hubo envíos aprobados: listá a quién se envió.

5. No abras PRs ni modifiques código salvo que el script esté roto y el arreglo sea mínimo.
6. No subas reportes a Drive en esta Automation.
7. Idioma: español, conciso.
```

8. Guardá y **Enable**.

## Paso 3 — Probar ahora (opcional)

Desde un Cloud Agent o PC con el secreto:

```bash
python revisar_correo_wes_borradores.py --forzar
```

## Cómo aprobás desde el celular

1. Te llega un mail: `[BORRADOR abc123…] Asunto original…`
2. Leés el original + la respuesta sugerida (y el informe si lo adjuntaron aparte).
3. Respondés ese mismo hilo:

```text
DALE NOMÁS abc123
```

o, si querés cambiar el texto:

```text
DALE NOMÁS abc123

Hola… (tu versión final)
```

4. En la próxima corrida (≤ ~30 min) se envía al destinatario y te llega `[ENVIADO]`.

## Relación con el monitor del PC

- El servicio Windows `WESMonitoreoCorreos` **sigue siendo otra cosa** (auto-genera reportes).
- Esta Automation es el reemplazo en la nube para **leer + borrador + tu OK**.
- Si ambos corren a la vez sobre los mismos UNSEEN, pueden pisarse: preferí **una sola** fuente de verdad para inbox, o apagá el monitor del PC cuando esta Automation esté estable.

## Variables opcionales

| Variable | Default | Uso |
|----------|---------|-----|
| `WES_APROBADOR_EMAIL` | `anibal.aoperaciones@wes.cl` | Quién aprueba |
| `WES_CORREO_HORA_INI` | `8` | Inicio ventana Chile |
| `WES_CORREO_HORA_FIN` | `18` | Fin (exclusivo) |
| `WES_CORREO_MAX_UNSEEN` | `15` | Máx. mails por corrida |

Estado: en Cloud Agent el disco es efímero; por eso el aviso `[BORRADOR id]` incluye un bloque
`META BORRADOR` (no lo borres al responder). Con eso la próxima corrida puede aprobar sin archivo local.

Si el monitor Windows `WESMonitoreoCorreos` sigue activo, pueden pisarse los UNSEEN: apagalo cuando esta Automation esté estable.

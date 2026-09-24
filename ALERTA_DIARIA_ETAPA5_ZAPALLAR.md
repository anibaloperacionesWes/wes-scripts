# Alerta diaria matutina — Fundo Zapallar Etapa N°5

**Punto:** `000027-03` — Etapa N°5 (Fundo Zapallar / `000027`)  
**Umbral ALERTA:** ≥ **60 m³/h** (techo hidráulico DN90 fierro dúctil)  
**Umbral AVISO:** ≥ **45 m³/h**  
**Script:** `python alerta_diaria_etapa5_zapallar.py --dias 2`

## Qué hacer cada mañana (timer Cloud Agent)

1. Ejecutar: `python alerta_diaria_etapa5_zapallar.py --dias 2`
2. Si `estado=OK`: responder al usuario en 2–3 líneas con máximo m³/h del período.
3. Si `estado=AVISO` o `ALERTA`: listar las horas/picos, recordar el criterio DN90 (~60 m³/h) y el antecedente de falla de memoria de placa (sep-2026); sugerir revisión de placa/sensor.
4. Si `ERROR` o `SIN_DATOS`: indicar el problema de API/datos.

No crear PR ni commits solo por la corrida diaria; solo avisar al usuario.

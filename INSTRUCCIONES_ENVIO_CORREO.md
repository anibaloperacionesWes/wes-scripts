# Instrucciones para Enviar Reportes por Correo

## ⚠️ Problema Detectado

El envío de correo falló porque:

1. **Si usas Google Workspace (@wes.cl)**: Necesitas una **"Contraseña de aplicación"**, no tu contraseña normal.
2. **Si usas servidor SMTP corporativo**: Necesitamos conocer el servidor SMTP correcto de tu empresa.

## 🔧 Solución 1: Usar Contraseña de Aplicación (Google Workspace)

Si tu correo `joseotarola@wes.cl` usa Google Workspace, sigue estos pasos:

### Paso 1: Activar Verificación en 2 Pasos
1. Ve a: https://myaccount.google.com/security
2. Activa "Verificación en 2 pasos" si no está activada

### Paso 2: Crear Contraseña de Aplicación
1. Ve a: https://myaccount.google.com/apppasswords
2. Selecciona "Correo" y "Otro (nombre personalizado)"
3. Escribe "WES Reportes" como nombre
4. Copia la contraseña de 16 caracteres que te genera (ejemplo: `abcd efgh ijkl mnop`)

### Paso 3: Usar la Contraseña de Aplicación

```bash
python generar_reporte_word.py \
  --company-id 000029 \
  --node-id 000029-01 \
  --start-date 01112025 \
  --end-date 30112025 \
  --enviar-correo \
  --destinatario joseotarola@wes.cl \
  --smtp-usuario joseotarola@wes.cl \
  --smtp-password "abcd efgh ijkl mnop" \
  --smtp-servidor smtp.gmail.com \
  --smtp-puerto 587
```

**Nota:** Usa la contraseña de aplicación (sin espacios o con espacios, ambos funcionan).

## 🔧 Solución 2: Usar Servidor SMTP Corporativo

Si tu empresa tiene su propio servidor SMTP, necesitamos:

1. **Servidor SMTP**: (ejemplo: `mail.wes.cl`, `smtp.wes.cl`, `outlook.office365.com`)
2. **Puerto**: (generalmente 587 para TLS o 465 para SSL)
3. **Confirmar si requiere autenticación especial**

Pregunta a tu departamento de TI por:
- "¿Cuál es el servidor SMTP saliente para enviar correos?"
- "¿Qué puerto usa? (587 o 465)"
- "¿Necesito alguna configuración especial?"

Una vez tengas esta información, el comando sería:

```bash
python generar_reporte_word.py \
  --company-id 000029 \
  --node-id 000029-01 \
  --start-date 01112025 \
  --end-date 30112025 \
  --enviar-correo \
  --destinatario joseotarola@wes.cl \
  --smtp-usuario joseotarola@wes.cl \
  --smtp-password tu_contraseña \
  --smtp-servidor SERVIDOR_SMTP_AQUI \
  --smtp-puerto PUERTO_AQUI
```

## 🧪 Prueba Rápida

Una vez tengas la contraseña de aplicación o el servidor SMTP correcto:

```bash
python enviar_reporte_prueba.py
```

Este script te pedirá toda la información interactivamente.

## 📋 Resumen de lo que Necesitas

- ✅ Correo remitente: `joseotarola@wes.cl`
- ✅ Correo destinatario: `joseotarola@wes.cl`
- ❓ Contraseña de aplicación (si Google Workspace) O servidor SMTP corporativo
- ❓ Servidor SMTP y puerto (si no es Gmail)

## 🔐 Seguridad

**IMPORTANTE:** 
- Nunca compartas tu contraseña de aplicación públicamente
- Considera usar variables de entorno para producción
- La contraseña de aplicación es específica para esta aplicación



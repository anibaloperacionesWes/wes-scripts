# 🤖 Guía: Bot de WhatsApp para Reportes WES

Esta guía explica cómo configurar y usar el bot de WhatsApp que permite solicitar reportes WES mediante mensajes de texto.

## 📋 Índice

1. [Requisitos](#requisitos)
2. [Configuración de Twilio](#configuración-de-twilio)
3. [Instalación](#instalación)
4. [Configuración del Bot](#configuración-del-bot)
5. [Uso del Bot](#uso-del-bot)
6. [Alternativas](#alternativas)
7. [Solución de Problemas](#solución-de-problemas)

---

## 🔧 Requisitos

### Software
- Python 3.10 o superior
- Cuenta de Twilio con WhatsApp Business API habilitada
- Acceso a internet para recibir webhooks

### Librerías Python
```bash
pip install twilio flask python-dotenv requests
```

---

## 📱 Configuración de Twilio

### Paso 1: Crear cuenta en Twilio

1. Ve a [twilio.com](https://www.twilio.com) y crea una cuenta
2. Verifica tu número de teléfono
3. Obtén tu **Account SID** y **Auth Token** desde el dashboard

### Paso 2: Habilitar WhatsApp Sandbox

1. En el dashboard de Twilio, ve a **Messaging** > **Try it out** > **Send a WhatsApp message**
2. Sigue las instrucciones para unirte al Sandbox de WhatsApp
3. Envía el código que te proporciona Twilio a su número de WhatsApp

### Paso 3: Obtener número de WhatsApp

Una vez en el Sandbox, Twilio te proporcionará un número de WhatsApp (formato: `whatsapp:+14155238886`)

### Paso 4: Configurar Webhook

1. En el dashboard de Twilio, ve a **Messaging** > **Settings** > **WhatsApp Sandbox Settings**
2. Configura el webhook URL para apuntar a tu servidor:
   - **Development**: `https://tu-dominio-ngrok.ngrok.io/whatsapp`
   - **Production**: `https://tu-dominio.com/whatsapp`

---

## 🚀 Instalación

### 1. Instalar dependencias

```bash
pip install twilio flask python-dotenv requests
```

### 2. Configurar variables de entorno

Crea un archivo `.env` en la raíz del proyecto:

```env
TWILIO_ACCOUNT_SID=ACxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
TWILIO_AUTH_TOKEN=tu_auth_token_aqui
TWILIO_WHATSAPP_FROM=whatsapp:+14155238886
TWILIO_WHATSAPP_TO=whatsapp:+56912345678
```

**Nota**: Reemplaza los valores con tus credenciales reales de Twilio.

### 3. Para desarrollo local (usando ngrok)

Si quieres probar el bot localmente, necesitas exponer tu servidor local:

```bash
# Instalar ngrok
# Windows: Descarga desde https://ngrok.com/download
# O usando chocolatey: choco install ngrok

# En una terminal, ejecuta:
ngrok http 5000

# Copia la URL HTTPS que te proporciona (ej: https://abc123.ngrok.io)
# Úsala para configurar el webhook en Twilio
```

---

## ⚙️ Configuración del Bot

### Ejecutar el bot

```bash
python bot_whatsapp_reportes.py
```

El bot iniciará un servidor Flask en `http://localhost:5000` y escuchará mensajes en `/whatsapp`.

### Verificar que funciona

1. Abre tu navegador en `http://localhost:5000/health`
2. Deberías ver: `{"status": "ok", "service": "whatsapp-bot-wes"}`

---

## 💬 Uso del Bot

### Comandos disponibles

#### 1. Reporte Individual

**Formato 1:**
```
reporte empresa 000025 nodo 000025-12 desde 01/12/2025 hasta 15/12/2025
```

**Formato 2 (simplificado):**
```
reporte 000025 000025-12 01/12/2025 15/12/2025
```

**Formato 3 (últimos N días):**
```
reporte empresa 000025 nodo 000025-12 ultimos 7 dias
```

#### 2. Reporte Agregado

```
reporte agregado empresa 000025 desde 01/12/2025 hasta 15/12/2025
```

#### 3. Ayuda

```
ayuda
```

### Ejemplos de uso

1. **Reporte de los últimos 7 días:**
   ```
   reporte 000025 000025-12 ultimos 7 dias
   ```

2. **Reporte de un periodo específico:**
   ```
   reporte empresa 000025 nodo 000025-12 desde 01/12/2025 hasta 15/12/2025
   ```

3. **Reporte agregado:**
   ```
   reporte agregado 000025 desde 01/12/2025 hasta 15/12/2025
   ```

### Flujo de trabajo

1. El usuario envía un mensaje con la solicitud de reporte
2. El bot confirma la solicitud
3. El bot genera el reporte en segundo plano
4. El bot envía un mensaje con la ubicación del archivo generado

**Nota**: Actualmente, el bot envía la ubicación del archivo. Para enviar el archivo directamente por WhatsApp, necesitas configurar un servidor web que aloje los archivos temporalmente (ver sección de mejoras).

---

## 🔄 Alternativas

### Opción 1: Twilio WhatsApp API (Recomendada - Actual)

✅ **Ventajas:**
- Oficial y estable
- No requiere mantener sesión activa
- Escalable
- Documentación completa

❌ **Desventajas:**
- Requiere cuenta de pago (aunque tiene tier gratuito)
- Requiere servidor web para webhooks
- Los archivos grandes necesitan URL pública

### Opción 2: whatsapp-web.js con Python

Esta opción usa `whatsapp-web.js` (Node.js) con un wrapper en Python.

**Requisitos:**
- Node.js instalado
- Librería `whatsapp-web.js`

**Ventajas:**
- Gratis (usa WhatsApp Web)
- Puede enviar archivos directamente

**Desventajas:**
- No oficial (puede romperse con actualizaciones de WhatsApp)
- Requiere mantener sesión activa (escaneo QR)
- Menos estable

### Opción 3: pywhatkit

Librería Python simple pero limitada.

**Ventajas:**
- Muy fácil de usar
- No requiere configuración compleja

**Desventajas:**
- Solo puede enviar mensajes programados
- No puede recibir mensajes
- No es adecuado para bots interactivos

---

## 🛠️ Solución de Problemas

### Error: "Twilio no configurado"

**Solución**: Verifica que el archivo `.env` existe y contiene las credenciales correctas.

### Error: "Número no autorizado"

**Solución**: Verifica que `TWILIO_WHATSAPP_TO` en `.env` coincide con el número que envía el mensaje.

### El bot no recibe mensajes

**Solución**: 
1. Verifica que el webhook de Twilio esté configurado correctamente
2. Si usas ngrok, asegúrate de que esté corriendo
3. Verifica que el servidor Flask esté ejecutándose

### Error al generar reporte

**Solución**:
1. Verifica que `PYTHON_EXE` en el script apunte a tu instalación de Python
2. Verifica que `SCRIPT_PATH` apunte al archivo `generar_reporte_word.py`
3. Revisa los logs del servidor para ver el error específico

### No puedo enviar archivos por WhatsApp

**Solución**: Twilio requiere que los archivos estén en una URL pública. Opciones:
1. Subir archivos a un servicio de almacenamiento (AWS S3, Google Cloud Storage)
2. Usar un servidor web temporal (ver sección de mejoras)

---

## 🚀 Mejoras Futuras

### 1. Envío directo de archivos

Para enviar archivos directamente por WhatsApp, necesitas:

1. **Subir archivos a un servidor web:**
   ```python
   # Ejemplo usando Flask para servir archivos temporalmente
   @app.route('/files/<filename>')
   def serve_file(filename):
       return send_from_directory('reports', filename)
   ```

2. **O usar un servicio de almacenamiento:**
   - AWS S3
   - Google Cloud Storage
   - Azure Blob Storage

### 2. Cola de tareas

Para manejar múltiples solicitudes simultáneas, considera usar:
- **Celery** con Redis/RabbitMQ
- **RQ** (Redis Queue)
- **APScheduler** para tareas programadas

### 3. Autenticación de usuarios

Agregar sistema de autenticación para que solo usuarios autorizados puedan solicitar reportes.

### 4. Historial de solicitudes

Guardar un registro de todas las solicitudes de reportes para auditoría.

---

## 📞 Soporte

Si tienes problemas o preguntas:
1. Revisa los logs del servidor Flask
2. Verifica la configuración de Twilio
3. Consulta la documentación de Twilio: https://www.twilio.com/docs/whatsapp

---

## 📝 Notas Importantes

- ⚠️ **Seguridad**: Nunca compartas tus credenciales de Twilio. Usa variables de entorno.
- ⚠️ **Límites**: Twilio tiene límites en el número de mensajes gratuitos. Revisa tu plan.
- ⚠️ **Producción**: Para producción, usa un servidor web real (no ngrok) y HTTPS.
- ⚠️ **Archivos**: Los archivos grandes pueden requerir almacenamiento en la nube.

---

*Última actualización: Diciembre 2025*












# Guía de Envío de Reportes por Correo

## 📧 Forma de Usar

### Opción 1: Generar y Enviar en un Solo Comando

```bash
python generar_reporte_word.py \
  --company-id 000029 \
  --node-id 000029-01 \
  --start-date 01112025 \
  --end-date 30112025 \
  --enviar-correo \
  --destinatario destinatario@empresa.com \
  --smtp-usuario tu.correo@empresa.com \
  --smtp-password tu_contraseña
```

### Opción 2: Usar Script de Prueba Interactivo

```bash
python enviar_reporte_prueba.py
```

Este script te pedirá:
- Correo del remitente
- Contraseña o contraseña de aplicación
- Correo del destinatario
- Servidor SMTP (opcional, default: smtp.gmail.com)
- Puerto SMTP (opcional, default: 587)

## 🔧 Parámetros Disponibles

| Parámetro | Requerido | Descripción | Default |
|-----------|-----------|-------------|---------|
| `--enviar-correo` | No | Activa el envío por correo | - |
| `--destinatario` | Sí (si `--enviar-correo`) | Correo del destinatario | - |
| `--smtp-usuario` | Sí (si `--enviar-correo`) | Correo del remitente | - |
| `--smtp-password` | Sí (si `--enviar-correo`) | Contraseña o contraseña de aplicación | - |
| `--smtp-servidor` | No | Servidor SMTP | smtp.gmail.com |
| `--smtp-puerto` | No | Puerto SMTP | 587 |

## 📝 Ejemplos de Uso

### Ejemplo 1: Gmail

```bash
python generar_reporte_word.py \
  --company-id 000029 \
  --node-id 000029-01 \
  --start-date 01112025 \
  --end-date 30112025 \
  --enviar-correo \
  --destinatario cliente@empresa.com \
  --smtp-usuario tu.correo@gmail.com \
  --smtp-password xxxx-xxxx-xxxx-xxxx \
  --smtp-servidor smtp.gmail.com \
  --smtp-puerto 587
```

**Importante para Gmail:**
- Necesitas una **"Contraseña de aplicación"**, no tu contraseña normal
- Obténla en: https://myaccount.google.com/apppasswords
- Activa la verificación en 2 pasos primero

### Ejemplo 2: Outlook/Office 365

```bash
python generar_reporte_word.py \
  --company-id 000029 \
  --node-id 000029-01 \
  --start-date 01112025 \
  --end-date 30112025 \
  --enviar-correo \
  --destinatario cliente@empresa.com \
  --smtp-usuario tu.correo@empresa.com \
  --smtp-password tu_contraseña \
  --smtp-servidor smtp.office365.com \
  --smtp-puerto 587
```

### Ejemplo 3: Servidor SMTP Corporativo

```bash
python generar_reporte_word.py \
  --company-id 000029 \
  --node-id 000029-01 \
  --start-date 01112025 \
  --end-date 30112025 \
  --enviar-correo \
  --destinatario cliente@empresa.com \
  --smtp-usuario tu.correo@empresa.com \
  --smtp-password tu_contraseña \
  --smtp-servidor mail.empresa.com \
  --smtp-puerto 587
```

## 🔐 Configuración de Servidores SMTP Comunes

### Gmail
- **Servidor:** smtp.gmail.com
- **Puerto:** 587 (TLS) o 465 (SSL)
- **Requisito:** Contraseña de aplicación

### Outlook/Office 365
- **Servidor:** smtp.office365.com
- **Puerto:** 587
- **Requisito:** Contraseña normal

### Yahoo
- **Servidor:** smtp.mail.yahoo.com
- **Puerto:** 587 o 465
- **Requisito:** Contraseña de aplicación

## ⚠️ Notas de Seguridad

1. **No compartas tu contraseña**: Nunca compartas tu contraseña o contraseña de aplicación
2. **Usa variables de entorno**: Para producción, considera usar variables de entorno para las credenciales
3. **Contraseñas de aplicación**: Para Gmail y otros servicios, usa contraseñas de aplicación en lugar de tu contraseña principal

## 🧪 Prueba Rápida

Para hacer una prueba rápida sin generar un nuevo reporte:

```bash
python enviar_reporte_prueba.py
```

Este script:
1. Busca el reporte más reciente generado
2. Te pide los datos de correo
3. Envía el reporte como prueba

## 📋 Contenido del Correo

El correo incluye:
- **Asunto:** "Reporte de Consumo - [Empresa] ([Nodo]) - [Fecha inicio] a [Fecha fin]"
- **Cuerpo:** Mensaje con información del reporte
- **Adjunto:** Archivo Word del reporte (.docx)

## ❓ Solución de Problemas

### Error: "SMTPAuthenticationError"
- Verifica que el usuario y contraseña sean correctos
- Si usas Gmail, asegúrate de usar una contraseña de aplicación
- Verifica que la verificación en 2 pasos esté activada (Gmail)

### Error: "Connection refused"
- Verifica que el servidor SMTP y puerto sean correctos
- Verifica tu conexión a internet
- Algunos servidores corporativos requieren VPN

### Error: "Timeout"
- Verifica que el puerto sea correcto
- Algunos firewalls bloquean conexiones SMTP
- Intenta con un puerto alternativo (465 para SSL)



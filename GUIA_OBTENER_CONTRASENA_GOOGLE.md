# 🔐 Guía: Obtener Contraseña de Aplicación de Google

Esta guía explica cómo obtener una **Contraseña de Aplicación** de Google para usar con SMTP en el sistema de monitoreo de correos WES.

---

## 📋 Requisitos Previos

- ✅ Tener acceso a la cuenta de Google Workspace: `agente.ia@wes.cl`
- ✅ Tener permisos de administrador (si es cuenta corporativa)
- ✅ Verificación en dos pasos (2FA) activada en la cuenta

---

## 🚀 Proceso Paso a Paso

### **Paso 1: Activar Verificación en Dos Pasos (2FA)**

Si la cuenta `agente.ia@wes.cl` **NO tiene 2FA activado**, debes activarlo primero:

1. Ve a: https://myaccount.google.com/security
2. Inicia sesión con `agente.ia@wes.cl`
3. Busca la sección **"Verificación en dos pasos"**
4. Haz clic en **"Activar"** y sigue las instrucciones
5. Configura un método de verificación (teléfono, app autenticador, etc.)

> ⚠️ **IMPORTANTE**: Las contraseñas de aplicación solo funcionan si la cuenta tiene 2FA activado.

---

### **Paso 2: Generar Contraseña de Aplicación**

Una vez que tengas 2FA activado:

1. Ve a: https://myaccount.google.com/apppasswords
   - O desde: https://myaccount.google.com/security → **"Contraseñas de aplicaciones"**

2. Si no ves la opción "Contraseñas de aplicaciones":
   - Asegúrate de que 2FA esté activado
   - Si usas Google Workspace, el administrador puede haber deshabilitado esta función

3. En la página de **"Contraseñas de aplicaciones"**:
   - Selecciona la **aplicación**: `Correo`
   - Selecciona el **dispositivo**: `Otro (nombre personalizado)`
   - Escribe un nombre: `WES Monitoreo Correos` o `WES Scripts`
   - Haz clic en **"Generar"**

4. **Google mostrará una contraseña de 16 caracteres** (sin espacios):
   ```
   xxxx xxxx xxxx xxxx
   ```
   > ⚠️ **COPIA ESTA CONTRASEÑA INMEDIATAMENTE** - Solo se muestra una vez

5. **Guarda la contraseña en un lugar seguro** (será tu `SMTP_PASSWORD`)

---

### **Paso 3: Usar la Contraseña en el Script**

La contraseña de aplicación se usa así:

```python
SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_PASSWORD = "xxxx xxxx xxxx xxxx"  # ← La contraseña de aplicación (sin espacios)
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587
```

> 💡 **Nota**: Quita los espacios de la contraseña cuando la uses en el código.

---

## 🔍 Verificación

### Probar la Contraseña

Puedes probar la contraseña con este script de prueba:

```python
import smtplib
from email.mime.text import MIMEText

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_PASSWORD = "TU_CONTRASENA_AQUI"  # Sin espacios
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

try:
    server = smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO)
    server.starttls()
    server.login(SMTP_USUARIO, SMTP_PASSWORD)
    print("✅ Conexión exitosa!")
    server.quit()
except Exception as e:
    print(f"❌ Error: {e}")
```

---

## ⚠️ Solución de Problemas

### **Problema 1: No aparece "Contraseñas de aplicaciones"**

**Causas posibles:**
- 2FA no está activado
- La cuenta es de Google Workspace y el administrador deshabilitó esta función
- La cuenta es muy nueva

**Soluciones:**
1. Verifica que 2FA esté activado: https://myaccount.google.com/security
2. Si es Google Workspace, contacta al administrador
3. Espera 24-48 horas después de activar 2FA

---

### **Problema 2: Error "Username and Password not accepted"**

**Causas posibles:**
- La contraseña tiene espacios (quítalos)
- La contraseña fue copiada incorrectamente
- La contraseña fue revocada

**Soluciones:**
1. Verifica que no haya espacios en la contraseña
2. Genera una nueva contraseña de aplicación
3. Asegúrate de usar el correo completo: `agente.ia@wes.cl`

---

### **Problema 3: "Less secure app access" (Apps menos seguras)**

Si Google Workspace tiene deshabilitado el acceso de apps menos seguras:

**Opción A: Usar OAuth 2.0** (Recomendado)
- Requiere configurar credenciales OAuth
- Más seguro pero más complejo

**Opción B: Habilitar acceso de apps menos seguras** (No recomendado)
- Solo para cuentas personales
- Google Workspace puede no permitirlo

---

## 📝 Alternativa: OAuth 2.0 (Avanzado)

Si las contraseñas de aplicación no funcionan (por ejemplo, en Google Workspace con políticas restrictivas), puedes usar OAuth 2.0:

1. Ve a: https://console.cloud.google.com/
2. Crea un proyecto o selecciona uno existente
3. Habilita la API de Gmail
4. Crea credenciales OAuth 2.0
5. Descarga el archivo `credentials.json`
6. Usa el flujo OAuth en el script

> 💡 **Nota**: El script actual (`monitorear_correos_y_generar_reportes.py`) ya tiene soporte para OAuth con Gmail API, pero usa credenciales diferentes.

---

## ✅ Checklist Final

Antes de usar la contraseña en producción:

- [ ] 2FA está activado en `agente.ia@wes.cl`
- [ ] Contraseña de aplicación generada
- [ ] Contraseña copiada y guardada de forma segura
- [ ] Contraseña probada con script de prueba
- [ ] Contraseña sin espacios en el código
- [ ] Correo configurado correctamente: `agente.ia@wes.cl`

---

## 🔒 Seguridad

**Buenas prácticas:**
- ✅ Guarda la contraseña en un archivo de configuración separado (no en el código)
- ✅ Usa variables de entorno si es posible
- ✅ No compartas la contraseña públicamente
- ✅ Revoca la contraseña si sospechas que fue comprometida
- ✅ Revisa periódicamente las contraseñas de aplicación activas

**Para revocar una contraseña:**
1. Ve a: https://myaccount.google.com/apppasswords
2. Busca la contraseña que quieres revocar
3. Haz clic en el ícono de eliminar (🗑️)

---

## 📞 Soporte

Si tienes problemas:
1. Verifica que 2FA esté activado
2. Intenta generar una nueva contraseña
3. Revisa los logs del script para ver el error específico
4. Si es Google Workspace, contacta al administrador

---

*Última actualización: Diciembre 2025*

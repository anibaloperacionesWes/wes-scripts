# Prompt para Crear Usuarios en WES API

## 📋 Contexto para el Asistente

Cuando necesites crear un usuario en el sistema WES, usa este prompt con los datos del usuario:

---

## 🎯 PROMPT ESTÁNDAR

```
Crea un usuario en WES API con los siguientes datos:

- Nombre de la empresa: [NOMBRE_EMPRESA]
- Nombre: [NOMBRE]
- Apellido: [APELLIDO]
- Correo: [CORREO]

El usuario debe tener acceso a TODOS los nodos de la empresa especificada.
```

**Nota**: El ID de la empresa es opcional. Si no se proporciona, el script buscará automáticamente la empresa por nombre.

---

## 📝 Ejemplo de Uso

### Ejemplo 1: Solo con nombre de empresa (recomendado)

```
Crea un usuario en WES API con los siguientes datos:

- Nombre de la empresa: Lo valledor
- Nombre: Jose Luis
- Apellido: Otarola
- Correo: joseluisricardo@hotmail.com

El usuario debe tener acceso a TODOS los nodos de la empresa especificada.
```

### Ejemplo 2: Con nombre e ID de empresa (opcional)

```
Crea un usuario en WES API con los siguientes datos:

- Nombre de la empresa: Lo valledor
- ID de la empresa: 000002
- Nombre: Jose Luis
- Apellido: Otarola
- Correo: joseluisricardo@hotmail.com

El usuario debe tener acceso a TODOS los nodos de la empresa especificada.
```

---

## 🔧 Qué Hace el Script

El script `crear_usuario.ps1` automáticamente:

1. **Busca la empresa** (si no se proporciona ID):
   - Busca desde `000001` hasta `000100` usando el endpoint `GET /companies/{companyId}`
   - Compara el nombre de la empresa hasta encontrar coincidencia
   - Si no encuentra, muestra error y solicita el ID manualmente

2. **Valida la empresa** (si se proporciona ID):
   - Verifica que existe y que el nombre coincide con el proporcionado

3. **Obtiene todos los nodos** de la compañía:
   - Usa el endpoint `GET /companies/{companyId}`
   - Extrae todos los `nodeId` de la respuesta

4. **Crea el usuario** con acceso a todos esos nodos:
   - Usa `POST /configuration/users`
   - Con los datos proporcionados y el array completo de `allowedNodes`

---

## ⚡ Comando Directo

Si prefieres ejecutar directamente sin el asistente:

### Solo con nombre de empresa (recomendado):
```powershell
.\crear_usuario.ps1 -NombreEmpresa "Lo valledor" -Nombre "Jose Luis" -Apellido "Otarola" -Email "joseluisricardo@hotmail.com"
```

### Con nombre e ID de empresa (opcional):
```powershell
.\crear_usuario.ps1 -NombreEmpresa "Lo valledor" -CompanyId "000002" -Nombre "Jose Luis" -Apellido "Otarola" -Email "joseluisricardo@hotmail.com"
```

---

## 📊 Estructura de Datos

El script crea un usuario con esta estructura:

```json
{
    "username": "[CORREO]",
    "name": "[NOMBRE]",
    "lastName": "[APELLIDO]",
    "companyId": "[ID_EMPRESA]",
    "allowedNodes": [
        "nodeId1",
        "nodeId2",
        "nodeId3",
        ...
    ]
}
```

Donde `allowedNodes` se completa automáticamente con TODOS los nodos de la empresa.

---

## ✅ Checklist de Datos Necesarios

- [ ] **Nombre de la empresa** (obligatorio)
- [ ] ID de la empresa (opcional - se busca automáticamente si no se proporciona)
- [ ] Nombre de la persona
- [ ] Apellido de la persona
- [ ] Correo electrónico

**Notas**: 
- No necesitas proporcionar los nodos, se obtienen automáticamente
- Si no proporcionas el ID, el script buscará la empresa por nombre desde 000001 hasta 000100

---

## 🔗 Información Adicional

- **API Base URL**: http://104.248.53.141:7001/wes/api/acl-entities/v1
- **Documentación**: Ver `GUIA_CREACION_USUARIOS.md` para más detalles
- **Script**: `crear_usuario.ps1`

---

*Este prompt está diseñado para uso futuro con el asistente de IA*


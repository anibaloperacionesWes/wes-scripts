# Guía Completa: Creación de Usuarios en WES API

## 📋 Índice
1. [Introducción](#introducción)
2. [Entendiendo la API WES](#entendiendo-la-api-wes)
3. [Proceso de Creación de Usuarios](#proceso-de-creación-de-usuarios)
4. [Datos Requeridos](#datos-requeridos)
5. [Cómo Usar el Script](#cómo-usar-el-script)
6. [Ejemplos Prácticos](#ejemplos-prácticos)
7. [Solución de Problemas](#solución-de-problemas)

---

## 🎯 Introducción

Esta guía explica cómo crear usuarios en el sistema WES (Water Efficiency System) utilizando la API REST. El proceso consta de dos pasos principales:
1. **Obtener los nodos** de una compañía
2. **Crear el usuario** con acceso a esos nodos

---

## 🔍 Entendiendo la API WES

### Base URL
```
http://104.248.53.141:7001/wes/api/acl-entities/v1
```

### Documentación
- **Swagger UI**: http://104.248.53.141:7001/wes/api/acl-entities/v1/swagger-ui.html
- **API Docs JSON**: http://104.248.53.141:7001/wes/api/acl-entities/v1/v2/api-docs

### Controladores Principales
- **entity-controller**: Operaciones de lectura (GET) de entidades
- **configuration-controller**: Operaciones de configuración (POST, PUT, DELETE)

---

## 📝 Proceso de Creación de Usuarios

### Paso 1: Obtener Nodos de la Compañía

**Endpoint**: `GET /companies/{companyId}`

**Propósito**: Obtener todos los nodos disponibles de una compañía para asignarlos al usuario.

**Ejemplo de Request**:
```powershell
Invoke-RestMethod -Uri "http://104.248.53.141:7001/wes/api/acl-entities/v1/companies/000002" `
    -Method Get `
    -Headers @{
        "Accept" = "application/json"
        "Content-Type" = "application/json"
    }
```

**Ejemplo de Response**:
```json
{
    "companyId": "000002",
    "name": "Lo valledor",
    "nodes": [
        {
            "nodeId": "000002-01",
            "name": "Lo Valledor - P1"
        },
        {
            "nodeId": "000002-02",
            "name": "Lo valledor - Pozo"
        },
        {
            "nodeId": "000002-03",
            "name": "Lo Valledor - Barrio Norte"
        }
    ]
}
```

### Paso 2: Crear el Usuario

**Endpoint**: `POST /configuration/users`

**Propósito**: Crear un nuevo usuario con acceso a los nodos especificados.

**Operación**: `createUserUsingPOST`

---

## 📊 Datos Requeridos

### Estructura del Request (CreateUserRequest)

| Campo | Tipo | Requerido | Descripción |
|-------|------|-----------|-------------|
| `username` | string | ✅ Sí | Email o nombre de usuario único |
| `name` | string | ✅ Sí | Nombre del usuario |
| `lastName` | string | ✅ Sí | Apellido del usuario |
| `companyId` | string | ✅ Sí | ID de la compañía (ej: "000002") |
| `allowedNodes` | array[string] | ✅ Sí | Lista de IDs de nodos permitidos |

### Ejemplo de JSON para Crear Usuario

```json
{
    "username": "usuario@ejemplo.com",
    "name": "Juan",
    "lastName": "Pérez",
    "companyId": "000002",
    "allowedNodes": [
        "000002-01",
        "000002-02",
        "000002-03"
    ]
}
```

---

## 🚀 Cómo Usar el Script

### Script Actual: `create_user.ps1`

El script actual está configurado para crear un usuario específico. Para crear más usuarios, necesitas modificar las siguientes variables:

#### Variables a Modificar (líneas 31-36):

```powershell
$userData = @{
    username = "NUEVO_EMAIL@ejemplo.com"    # ← Cambiar email
    name = "Nombre"                          # ← Cambiar nombre
    lastName = "Apellido"                    # ← Cambiar apellido
    companyId = "000002"                     # ← Cambiar si es otra compañía
    allowedNodes = $allowedNodes             # ← Se obtiene automáticamente
} | ConvertTo-Json -Depth 10
```

#### Para Cambiar la Compañía (línea 13):

```powershell
$companyResponse = Invoke-RestMethod -Uri "http://104.248.53.141:7001/wes/api/acl-entities/v1/companies/000002" `
```

Cambia `000002` por el ID de la compañía deseada.

### Ejecutar el Script

```powershell
powershell -ExecutionPolicy Bypass -File create_user.ps1
```

---

## 💡 Ejemplos Prácticos

### Ejemplo 1: Crear Usuario para Compañía 000002

**Datos del Usuario**:
- Email: `maria.garcia@empresa.com`
- Nombre: `María`
- Apellido: `García`
- Compañía: `000002`

**Modificación en el script**:
```powershell
$userData = @{
    username = "maria.garcia@empresa.com"
    name = "María"
    lastName = "García"
    companyId = "000002"
    allowedNodes = $allowedNodes
} | ConvertTo-Json -Depth 10
```

### Ejemplo 2: Crear Usuario para Otra Compañía

**Datos del Usuario**:
- Email: `carlos.rodriguez@empresa.com`
- Nombre: `Carlos`
- Apellido: `Rodríguez`
- Compañía: `000001` (diferente)

**Modificaciones necesarias**:
1. Cambiar el endpoint de obtención de compañía (línea 13):
```powershell
$companyResponse = Invoke-RestMethod -Uri "http://104.248.53.141:7001/wes/api/acl-entities/v1/companies/000001" `
```

2. Cambiar el companyId en userData (línea 35):
```powershell
companyId = "000001"
```

### Ejemplo 3: Crear Script Parametrizado

Puedes crear una versión del script que acepte parámetros:

```powershell
param(
    [Parameter(Mandatory=$true)]
    [string]$Email,
    
    [Parameter(Mandatory=$true)]
    [string]$Nombre,
    
    [Parameter(Mandatory=$true)]
    [string]$Apellido,
    
    [Parameter(Mandatory=$true)]
    [string]$CompanyId
)

# ... resto del código usando las variables $Email, $Nombre, $Apellido, $CompanyId
```

**Uso**:
```powershell
.\create_user.ps1 -Email "nuevo@ejemplo.com" -Nombre "Nuevo" -Apellido "Usuario" -CompanyId "000002"
```

---

## 🔧 Solución de Problemas

### Error 406 (Not Acceptable)
**Causa**: Headers incorrectos o formato de body inválido.

**Solución**: Asegúrate de usar:
```powershell
-Headers @{
    "Accept" = "application/json"
    "Content-Type" = "application/json"
}
```

### Error 404 (Not Found)
**Causa**: 
- El `companyId` no existe
- El endpoint está mal escrito

**Solución**: 
- Verifica que el `companyId` sea correcto
- Usa `get_company.ps1` para verificar que la compañía existe

### Error 500 (Internal Server Error)
**Causa**: Error del servidor o datos inválidos.

**Solución**: 
- Verifica que todos los campos requeridos estén presentes
- Asegúrate de que los `nodeId` en `allowedNodes` sean válidos

### El Usuario se Crea pero Sin Nodos
**Causa**: El array `allowedNodes` está vacío o los nodos no existen.

**Solución**: 
- Verifica que la compañía tenga nodos
- Revisa que los `nodeId` extraídos sean correctos

---

## 📚 Endpoints Relacionados

### Obtener Usuario por Email
```
GET /users?email={email}
```

### Obtener Usuario por ID
```
GET /users/{userId}
```

### Eliminar Usuario
```
DELETE /configuration/users/{userId}
```

### Agregar Nodo a Usuario
```
PUT /configuration/users/{userId}/allowed/{nodeId}
```

### Eliminar Nodo de Usuario
```
DELETE /configuration/users/{userId}/nodes/{nodeId}
```

---

## ✅ Checklist para Crear un Usuario

- [ ] Conocer el `companyId` de la compañía
- [ ] Verificar que la compañía existe (usar `get_company.ps1`)
- [ ] Tener el email del usuario
- [ ] Tener el nombre completo del usuario
- [ ] Modificar el script con los datos correctos
- [ ] Ejecutar el script
- [ ] Verificar que la respuesta sea exitosa (Status 200)
- [ ] (Opcional) Verificar el usuario creado con `GET /users?email={email}`

---

## 📝 Notas Importantes

1. **Nodos Automáticos**: El script actual obtiene automáticamente TODOS los nodos de la compañía. Si necesitas asignar solo algunos nodos específicos, modifica el array `$allowedNodes`.

2. **Email como Username**: Generalmente el `username` es el email del usuario, pero puede ser cualquier identificador único.

3. **Validación**: La API no valida el formato del email, pero es recomendable usar un formato válido.

4. **Nodos Vacíos**: Si una compañía no tiene nodos, el array `allowedNodes` estará vacío y el usuario se creará sin acceso a nodos.

5. **Documentación Completa**: Para ver todos los endpoints disponibles, visita: http://104.248.53.141:7001/wes/api/acl-entities/v1/swagger-ui.html

---

## 🎓 Resumen Rápido

**Para crear un usuario necesitas:**

1. **Obtener nodos**: `GET /companies/{companyId}` → Extraer `nodeId` de cada nodo
2. **Crear usuario**: `POST /configuration/users` con:
   - `username` (email)
   - `name` (nombre)
   - `lastName` (apellido)
   - `companyId` (ID compañía)
   - `allowedNodes` (array de nodeIds)

**El script `create_user.ps1` hace ambos pasos automáticamente. Solo modifica los datos del usuario y ejecuta.**

---

*Última actualización: Basado en la documentación de la API WES v1*


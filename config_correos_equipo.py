"""
Configuración de correos del equipo WES.
Este archivo contiene las direcciones de correo del equipo para uso en envíos automáticos.
"""

# Correos del equipo WES
CORREOS_EQUIPO_WES = {
    # Dirección principal
    "anibal": {
        "nombre": "Anibal Aranda Alvarado",
        "email": "anibal.aoperaciones@wes.cl"
    },
    
    # Equipo técnico/desarrollo
    "diego": {
        "nombre": "Diego Ignacio Carrasco Vallejos",
        "email": "diegocarrasco@wes.cl"
    },
    "juan": {
        "nombre": "Juan Manuel Lopez Delgado",
        "email": "juanlopez@wes.cl"
    },
    "benjamin": {
        "nombre": "Benjamín Gumucio Labbe",
        "email": "benjamingumucio@wes.cl"
    },
}

# Lista de correos para CC (copia) - solo emails
CORREOS_CC_DEFAULT = [
    "diegocarrasco@wes.cl",
    "juanlopez@wes.cl",
    "benjamingumucio@wes.cl",
]

# Lista de correos para TO (destinatario principal) - solo emails
CORREOS_TO_DEFAULT = [
    "anibal.aoperaciones@wes.cl",
]

# Función helper para obtener solo los emails
def obtener_emails():
    """Retorna una lista con solo los emails del equipo."""
    return [info["email"] for info in CORREOS_EQUIPO_WES.values()]

# Función helper para obtener email por nombre clave
def obtener_email(rol):
    """Retorna el email de un rol específico."""
    return CORREOS_EQUIPO_WES.get(rol, {}).get("email")

# Función helper para obtener nombre completo por rol
def obtener_nombre(rol):
    """Retorna el nombre completo de un rol específico."""
    return CORREOS_EQUIPO_WES.get(rol, {}).get("nombre")

# Función helper para obtener todos los correos (solo emails)
def obtener_todos_los_correos():
    """Retorna una lista con todos los emails del equipo."""
    return obtener_emails()

# Función helper para obtener correos por rol
def obtener_correos_por_rol(*roles):
    """Retorna una lista de emails basada en los roles especificados."""
    return [obtener_email(rol) for rol in roles if obtener_email(rol)]


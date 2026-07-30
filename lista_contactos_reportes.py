"""
Lista de contactos para personalización de correos y reportes.
Cada contacto contiene toda la información necesaria para elaborar correos personalizados.
"""

# Lista de correos AUTORIZADOS para recibir reportes automáticos
# Cada entrada contiene:
# - email: Correo autorizado
# - puntos_monitoreo: Lista de IDs de puntos específicos O nombre de empresa (ej: "DERCO", "BUPA")
#   Si es "DERCO" o nombre de empresa, se generarán reportes de TODOS los puntos de esa empresa
# - periodo: Diccionario con "inicio" y "fin" (formato: "DD/MM/YYYY") O None para periodo automático
# - tipo_reporte: "individual", "agregado", "ambos"
# - empresa_id: ID de la empresa (opcional, se puede inferir del nombre)
CORREOS_AUTORIZADOS = [
    {
        "email": "jose.otarola@wildstream.cl",
        "puntos_monitoreo": ["Todas"],  # "Todas" permite solicitar reportes de cualquier empresa
        "periodo": None,  # None = periodo automático (últimos 30 días o según solicitud)
        "tipo_reporte": "ambos",  # individual, agregado, ambos
        "empresa_id": None,  # Se puede inferir del nombre o especificar
        "notas": "Puede solicitar reportes de todas las empresas"
    },
    {
        "email": "juanlopez@wes.cl",
        "puntos_monitoreo": ["Todas"],  # "Todas" permite solicitar reportes de cualquier empresa
        "periodo": None,  # None = periodo automático (últimos 30 días o según solicitud)
        "tipo_reporte": "ambos",  # individual, agregado, ambos
        "empresa_id": None,  # Se puede inferir del nombre o especificar
        "notas": "Juan Manuel Lopez Delgado - Puede solicitar reportes de todas las empresas"
    },
    {
        "email": "benjamingumucio@wes.cl",
        "puntos_monitoreo": ["Todas"],  # "Todas" permite solicitar reportes de cualquier empresa
        "periodo": None,  # None = periodo automático (últimos 30 días o según solicitud)
        "tipo_reporte": "ambos",  # individual, agregado, ambos
        "empresa_id": None,  # Se puede inferir del nombre o especificar
        "notas": "Benjamín Gumucio Labbe - Puede solicitar reportes de todas las empresas"
    },
    {
        "email": "diegocarrasco@wes.cl",
        "puntos_monitoreo": ["Todas"],  # "Todas" permite solicitar reportes de cualquier empresa
        "periodo": None,  # None = periodo automático (últimos 30 días o según solicitud)
        "tipo_reporte": "ambos",  # individual, agregado, ambos
        "empresa_id": None,  # Se puede inferir del nombre o especificar
        "notas": "Diego Ignacio Carrasco Vallejos - Puede solicitar reportes de todas las empresas"
    },
    {
        "email": "anibal.aoperaciones@wes.cl",
        "puntos_monitoreo": ["Todas"],  # "Todas" permite solicitar reportes de cualquier empresa
        "periodo": None,  # None = periodo automático (últimos 30 días o según solicitud)
        "tipo_reporte": "ambos",  # individual, agregado, ambos
        "empresa_id": None,  # Se puede inferir del nombre o especificar
        "notas": "Anibal Aranda Alvarado - Puede solicitar reportes de todas las empresas"
    },
    {
        "email": "agente.ia@wes.cl",
        "puntos_monitoreo": ["Todas"],  # "Todas" permite solicitar reportes de cualquier empresa
        "periodo": None,  # None = periodo automático (últimos 30 días o según solicitud)
        "tipo_reporte": "ambos",  # individual, agregado, ambos
        "empresa_id": None,  # Se puede inferir del nombre o especificar
        "notas": "Agente IA WES - Puede solicitar reportes de todas las empresas"
    },
    {
        "email": "joseotarola@wes.cl",
        "puntos_monitoreo": ["Todas"],  # "Todas" permite solicitar reportes de cualquier empresa
        "periodo": None,  # None = periodo automático (últimos 30 días o según solicitud)
        "tipo_reporte": "ambos",  # individual, agregado, ambos
        "empresa_id": None,  # Se puede inferir del nombre o especificar
        "notas": "Jose Otarola - Puede solicitar reportes de todas las empresas"
    },
    {
        "email": "felipecuevas.mancilla@gmail.com",
        "puntos_monitoreo": ["Todas"],  # "Todas" permite solicitar reportes de cualquier empresa
        "periodo": None,  # None = periodo automático (últimos 30 días o según solicitud)
        "tipo_reporte": "ambos",  # individual, agregado, ambos
        "empresa_id": None,  # Se puede inferir del nombre o especificar
        "notas": "Felipe Cuevas Mancilla - Puede solicitar reportes de todas las empresas"
    },
]

# Función para verificar si un correo está autorizado
def esta_autorizado(email):
    """
    Verifica si un correo está autorizado para recibir reportes.
    
    Args:
        email: Dirección de correo a verificar
    
    Returns:
        Diccionario con la configuración del correo autorizado o None si no está autorizado
    """
    if not CORREOS_AUTORIZADOS:
        return None  # Si la lista está vacía, nadie está autorizado
    
    for config in CORREOS_AUTORIZADOS:
        if config["email"].lower() == email.lower():
            return config
    
    return None

# Función para obtener configuración de un correo autorizado
def obtener_configuracion_autorizado(email):
    """
    Obtiene la configuración completa de un correo autorizado.
    
    Args:
        email: Dirección de correo
    
    Returns:
        Diccionario con la configuración o None
    """
    return esta_autorizado(email)

CONTACTOS_REPORTES = {
    "diego": {
        "nombre": "Diego Ignacio",
        "apellido": "Carrasco Vallejos",
        "nombre_completo": "Diego Ignacio Carrasco Vallejos",
        "email": "diegocarrasco@wes.cl",
        "tratamiento": "Estimado Diego",
        "despedida": "Quedo atento a tus comentarios.",
        "firma": "Tu Agente WES",
        "preferencias": {
            "formato_reporte": "completo",  # completo, resumen, ejecutivo
            "incluir_graficas": True,
            "incluir_detalle": True,
            "idioma": "español"
        },
        "empresas_interes": ["BUPA", "Parque Arauco"],  # Empresas de las que quiere reportes
        "notas": "Prefiere reportes detallados con todas las gráficas"
    },
    
    "juan": {
        "nombre": "Juan Manuel",
        "apellido": "Lopez Delgado",
        "nombre_completo": "Juan Manuel Lopez Delgado",
        "email": "juanlopez@wes.cl",
        "tratamiento": "Estimado Juan",
        "despedida": "Quedo a tu disposición para cualquier consulta.",
        "firma": "Tu Agente WES",
        "preferencias": {
            "formato_reporte": "ejecutivo",
            "incluir_graficas": True,
            "incluir_detalle": False,
            "idioma": "español"
        },
        "empresas_interes": ["BUPA", "Parque Arauco"],
        "notas": "Prefiere reportes ejecutivos con resumen"
    },
    
    "benjamin": {
        "nombre": "Benjamín",
        "apellido": "Gumucio Labbe",
        "nombre_completo": "Benjamín Gumucio Labbe",
        "email": "benjamingumucio@wes.cl",
        "tratamiento": "Estimado Benjamín",
        "despedida": "Estoy disponible para cualquier aclaración que necesites.",
        "firma": "Tu Agente WES",
        "preferencias": {
            "formato_reporte": "completo",
            "incluir_graficas": True,
            "incluir_detalle": True,
            "idioma": "español"
        },
        "empresas_interes": ["BUPA", "Parque Arauco"],
        "notas": "Interesado en análisis técnicos detallados"
    },
    
    "anibal": {
        "nombre": "Anibal",
        "apellido": "Aranda Alvarado",
        "nombre_completo": "Anibal Aranda Alvarado",
        "email": "anibal.aoperaciones@wes.cl",
        "tratamiento": "Estimado Aníbal",
        "despedida": "Quedo atento a tus indicaciones.",
        "firma": "Tu Agente WES",
        "preferencias": {
            "formato_reporte": "ejecutivo",
            "incluir_graficas": True,
            "incluir_detalle": False,
            "idioma": "español"
        },
        "empresas_interes": ["BUPA", "Parque Arauco", "Todas"],
        "notas": "Preferencia por reportes ejecutivos con enfoque estratégico"
    },
    
    "agente": {
        "nombre": "Agente",
        "apellido": "IA",
        "nombre_completo": "Agente IA",
        "email": "agente.ia@wes.cl",
        "tratamiento": "Estimado Jose",
        "despedida": "Quedo atento a tus comentarios.",
        "firma": "Tu Agente WES",
        "preferencias": {
            "formato_reporte": "completo",
            "incluir_graficas": True,
            "incluir_detalle": True,
            "idioma": "español"
        },
        "empresas_interes": ["BUPA", "Parque Arauco", "Todas"],
        "notas": "Propietario del sistema"
    },
    
    # Agregar más contactos según sea necesario
    # "nombre_contacto": {
    #     "nombre": "...",
    #     "apellido": "...",
    #     ...
    # }
}

# Función helper para obtener información de un contacto
def obtener_contacto(identificador):
    """
    Obtiene la información de un contacto por su identificador.
    
    Args:
        identificador: Puede ser el nombre clave (ej: "diego") o el email
    
    Returns:
        Diccionario con la información del contacto o None si no se encuentra
    """
    # Buscar por nombre clave
    if identificador.lower() in CONTACTOS_REPORTES:
        return CONTACTOS_REPORTES[identificador.lower()]
    
    # Buscar por email
    for key, contacto in CONTACTOS_REPORTES.items():
        if contacto["email"].lower() == identificador.lower():
            return contacto
    
    # Buscar por nombre completo
    for key, contacto in CONTACTOS_REPORTES.items():
        if contacto["nombre_completo"].lower() == identificador.lower():
            return contacto
    
    return None

# Función helper para obtener contacto por email
def obtener_contacto_por_email(email):
    """Obtiene un contacto por su dirección de correo."""
    return obtener_contacto(email)

# Función helper para listar todos los contactos
def listar_contactos():
    """Retorna una lista con todos los contactos."""
    return list(CONTACTOS_REPORTES.values())

# Función helper para obtener contactos interesados en una empresa
def obtener_contactos_por_empresa(empresa):
    """Retorna los contactos interesados en reportes de una empresa específica."""
    contactos = []
    for key, contacto in CONTACTOS_REPORTES.items():
        if empresa in contacto["empresas_interes"] or "Todas" in contacto["empresas_interes"]:
            contactos.append(contacto)
    return contactos


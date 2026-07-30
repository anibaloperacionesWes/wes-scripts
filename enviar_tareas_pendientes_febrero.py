"""
Script para enviar correo con tareas pendientes de febrero 2026.
"""

import sys
from pathlib import Path
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from docx import Document

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

# Configuración SMTP
SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_PASSWORD = "vxbynfpoehbweelj"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

# Destinatarios
DESTINATARIOS = [
    "anibal.aoperaciones@wes.cl",  # Aníbal
    "juanlopez@wes.cl",  # Juan
    "diegocarrasco@wes.cl",  # Diego
]

# Ruta del documento Word
WORD_PATH = Path("Tareas Pendientes febrero 2026.docx")


def leer_tareas_del_word(word_path: Path) -> str:
    """Lee el contenido del documento Word y extrae las tareas."""
    if not word_path.exists():
        raise FileNotFoundError(f"No se encontró el archivo: {word_path}")
    
    doc = Document(str(word_path))
    tareas_texto = []
    
    for para in doc.paragraphs:
        texto = para.text.strip()
        if texto:
            tareas_texto.append(texto)
    
    return "\n".join(tareas_texto)


def formatear_tareas_para_correo(tareas_texto: str) -> str:
    """Formatea las tareas en un formato profesional para el correo."""
    lineas = tareas_texto.split('\n')
    
    # Separar título y tareas
    titulo = lineas[0] if lineas else ""
    tareas = [linea for linea in lineas[1:] if linea.strip()]
    
    # Organizar tareas por sección
    cuerpo_html = f"""
    <div style="font-family: Arial, sans-serif; font-size: 11pt; line-height: 1.6; color: #333;">
        <p>Estimado/a,</p>
        
        <p>Me dirijo a usted para informarle sobre las tareas pendientes de operaciones WES para el mes de febrero 2026. 
        Estas tareas fueron elaboradas por <strong>José Otarola</strong> producto de que estará en período de vacaciones.</p>
        
        <p>A continuación, se detallan las tareas pendientes organizadas por cliente/área:</p>
        
        <div style="margin-top: 20px; margin-bottom: 20px;">
"""
    
    seccion_actual = None
    for tarea in tareas:
        # Detectar si es una sección (títulos en mayúsculas o sin punto al final)
        es_seccion = (
            tarea.isupper() or 
            (len(tarea) < 50 and not tarea.endswith('.')) or
            tarea in ["Parque Arauco - Curauma", "MAE", "MAM", "Kennedy", "BOM", "AEB", "MAQ", 
                     "Fundo Zapallar", "Las Condes", "La Reina", "CORMUP", "Puente Alto", 
                     "CDUC", "Renca", "Furgones"]
        )
        
        if es_seccion:
            if seccion_actual:
                cuerpo_html += "        </ul>\n"
            seccion_actual = tarea
            cuerpo_html += f"""
            <h3 style="color: #1f4788; font-size: 13pt; margin-top: 20px; margin-bottom: 10px; border-bottom: 2px solid #1f4788; padding-bottom: 5px;">
                {tarea}
            </h3>
            <ul style='margin-left: 20px; margin-bottom: 15px;'>
"""
        else:
            if seccion_actual:
                cuerpo_html += f"                <li style='margin-bottom: 8px;'>{tarea}</li>\n"
            else:
                # Si no hay sección, crear una lista general
                if "ul" not in cuerpo_html[-100:]:
                    cuerpo_html += "            <ul style='margin-left: 20px; margin-bottom: 15px;'>\n"
                cuerpo_html += f"                <li style='margin-bottom: 8px;'>{tarea}</li>\n"
    
    if seccion_actual:
        cuerpo_html += "            </ul>\n"
    
    cuerpo_html += """
        </div>
        
        <p>Quedo atento a cualquier consulta o requerimiento adicional.</p>
        
        <p>Saludos cordiales,<br>
        <strong>Sistema WES</strong></p>
    </div>
"""
    
    return cuerpo_html


def enviar_correo(tareas_html: str):
    """Envía el correo con las tareas pendientes."""
    print("=" * 70)
    print("  ENVÍO DE CORREO - TAREAS PENDIENTES FEBRERO 2026")
    print("=" * 70)
    print()
    
    try:
        # Crear mensaje
        msg = MIMEMultipart('alternative')
        msg["From"] = SMTP_USUARIO
        msg["To"] = ", ".join(DESTINATARIOS)
        msg["Subject"] = "Tareas pendientes operaciones WES"
        
        # Crear versión texto plano (simplificada)
        texto_plano = """
Estimado/a,

Me dirijo a usted para informarle sobre las tareas pendientes de operaciones WES para el mes de febrero 2026. 
Estas tareas fueron elaboradas por José Otarola producto de que estará en período de vacaciones.

Por favor, revise el contenido HTML del correo para ver el detalle completo de las tareas.

Saludos cordiales,
Sistema WES
"""
        
        # Adjuntar ambas versiones
        part1 = MIMEText(texto_plano, "plain", "utf-8")
        part2 = MIMEText(tareas_html, "html", "utf-8")
        
        msg.attach(part1)
        msg.attach(part2)
        
        # Enviar correo
        print(f"[1/2] Conectando al servidor SMTP...")
        with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PUERTO) as server:
            server.starttls()
            print(f"[2/2] Autenticando y enviando correo...")
            server.login(SMTP_USUARIO, SMTP_PASSWORD)
            server.send_message(msg, to_addrs=DESTINATARIOS)
        
        print()
        print("=" * 70)
        print("  CORREO ENVIADO EXITOSAMENTE")
        print("=" * 70)
        print(f"Destinatarios:")
        for dest in DESTINATARIOS:
            print(f"  - {dest}")
        print(f"Asunto: Tareas pendientes operaciones WES")
        print()
        
        return True
    except Exception as e:
        print(f"[ERROR] Error al enviar correo: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """Función principal."""
    print("=" * 70)
    print("  LECTURA Y ENVÍO DE TAREAS PENDIENTES")
    print("=" * 70)
    print()
    
    # Leer tareas del Word
    print("[1/3] Leyendo tareas del documento Word...")
    try:
        tareas_texto = leer_tareas_del_word(WORD_PATH)
        print(f"[OK] Documento leído: {WORD_PATH.name}")
        print()
    except Exception as e:
        print(f"[ERROR] No se pudo leer el documento: {e}")
        return 1
    
    # Formatear tareas para correo
    print("[2/3] Formateando tareas para correo...")
    tareas_html = formatear_tareas_para_correo(tareas_texto)
    print("[OK] Tareas formateadas")
    print()
    
    # Enviar correo
    print("[3/3] Enviando correo...")
    exito = enviar_correo(tareas_html)
    
    if not exito:
        return 1
    
    return 0


if __name__ == "__main__":
    sys.exit(main())

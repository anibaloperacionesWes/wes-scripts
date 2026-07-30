"""Script de prueba para enviar un reporte por correo."""

import sys
from pathlib import Path
from generar_reporte_word import enviar_reporte_por_correo, get_company_name, get_node_name, parse_date

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

def main():
    print("=" * 60)
    print("PRUEBA DE ENVÍO DE REPORTE POR CORREO")
    print("=" * 60)
    print()
    
    # Solicitar información al usuario
    print("Por favor, ingresa la siguiente información:")
    print()
    
    # Buscar un reporte reciente para la prueba
    reports_dir = Path("reports")
    reportes_recientes = []
    
    if reports_dir.exists():
        for empresa_dir in reports_dir.iterdir():
            if empresa_dir.is_dir():
                reporte_dir = empresa_dir / "REPORTE"
                if reporte_dir.exists():
                    for node_dir in reporte_dir.iterdir():
                        if node_dir.is_dir():
                            for archivo in node_dir.glob("*.docx"):
                                if "Reporte_" in archivo.name:
                                    reportes_recientes.append(archivo)
    
    if reportes_recientes:
        # Ordenar por fecha de modificación (más reciente primero)
        reportes_recientes.sort(key=lambda x: x.stat().st_mtime, reverse=True)
        reporte_prueba = reportes_recientes[0]
        print(f"Reporte encontrado para prueba: {reporte_prueba}")
        print()
    else:
        print("[ERROR] No se encontraron reportes. Genera un reporte primero.")
        return
    
    # Solicitar datos de correo
    print("Configuración del correo:")
    print()
    
    smtp_usuario = input("Correo del remitente (ej: tu.correo@empresa.com): ").strip()
    if not smtp_usuario:
        print("[ERROR] Se requiere el correo del remitente")
        return
    
    smtp_password = input("Contraseña o contraseña de aplicación: ").strip()
    if not smtp_password:
        print("[ERROR] Se requiere la contraseña")
        return
    
    destinatario = input("Correo del destinatario (ej: destinatario@empresa.com): ").strip()
    if not destinatario:
        print("[ERROR] Se requiere el correo del destinatario")
        return
    
    smtp_servidor = input("Servidor SMTP (Enter para usar smtp.gmail.com): ").strip() or "smtp.gmail.com"
    smtp_puerto = input("Puerto SMTP (Enter para usar 587): ").strip() or "587"
    
    try:
        smtp_puerto = int(smtp_puerto)
    except ValueError:
        smtp_puerto = 587
    
    print()
    print("=" * 60)
    print("ENVIANDO CORREO...")
    print("=" * 60)
    print()
    
    # Extraer información del nombre del archivo si es posible
    company_name = None
    node_name = None
    start_date = None
    end_date = None
    
    # Intentar extraer del path
    try:
        parts = reporte_prueba.name.replace("Reporte_", "").replace(".docx", "").split("_")
        if len(parts) >= 3:
            company_id = parts[0]
            node_id = parts[1] if "-" in parts[1] else None
            if company_id:
                company_name = get_company_name(company_id)
            if node_id:
                node_name = get_node_name(node_id)
    except Exception:
        pass
    
    # Enviar correo
    exito = enviar_reporte_por_correo(
        reporte_path=reporte_prueba,
        destinatario=destinatario,
        smtp_servidor=smtp_servidor,
        smtp_puerto=smtp_puerto,
        smtp_usuario=smtp_usuario,
        smtp_password=smtp_password,
        company_name=company_name,
        node_name=node_name,
        start_date=start_date,
        end_date=end_date,
    )
    
    print()
    if exito:
        print("=" * 60)
        print("[OK] PRUEBA EXITOSA - Correo enviado correctamente")
        print("=" * 60)
    else:
        print("=" * 60)
        print("[ERROR] PRUEBA FALLIDA - Revisa los errores anteriores")
        print("=" * 60)
        print()
        print("NOTAS:")
        print("- Si usas Gmail, necesitas una 'Contraseña de aplicación'")
        print("  (no tu contraseña normal). Obténla en:")
        print("  https://myaccount.google.com/apppasswords")
        print("- Verifica que el servidor SMTP y puerto sean correctos")
        print("- Algunas empresas usan servidores SMTP internos")

if __name__ == "__main__":
    main()



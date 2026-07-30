"""Script interactivo para probar el envío de correo con Google Workspace."""

import sys
from pathlib import Path
from generar_reporte_word import enviar_reporte_por_correo, get_company_name, get_node_name

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

def main():
    print("=" * 60)
    print("PRUEBA DE ENVÍO DE REPORTE POR CORREO")
    print("Google Workspace (agente.ia@wes.cl)")
    print("=" * 60)
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
                                if "Reporte_" in archivo.name and "Agregado" not in archivo.name:
                                    reportes_recientes.append(archivo)
    
    if not reportes_recientes:
        print("[ERROR] No se encontraron reportes individuales.")
        print("Genera un reporte primero con:")
        print("  python generar_reporte_word.py --company-id 000029 --node-id 000029-01 --start-date 01112025 --end-date 30112025")
        return
    
    # Ordenar por fecha de modificación (más reciente primero)
    reportes_recientes.sort(key=lambda x: x.stat().st_mtime, reverse=True)
    reporte_prueba = reportes_recientes[0]
    
    print(f"Reporte encontrado para prueba:")
    print(f"  {reporte_prueba}")
    print()
    
    # Información predefinida
    smtp_usuario = "agente.ia@wes.cl"
    destinatario = "agente.ia@wes.cl"
    smtp_servidor = "smtp.gmail.com"
    smtp_puerto = 587
    
    print("Configuración:")
    print(f"  Remitente: {smtp_usuario}")
    print(f"  Destinatario: {destinatario}")
    print(f"  Servidor SMTP: {smtp_servidor}")
    print(f"  Puerto: {smtp_puerto}")
    print()
    print("=" * 60)
    print("IMPORTANTE: Necesitas una CONTRASEÑA DE APLICACIÓN")
    print("=" * 60)
    print()
    print("No uses tu contraseña normal. Necesitas generar una")
    print("'Contraseña de aplicación' desde:")
    print("  https://myaccount.google.com/apppasswords")
    print()
    print("Pasos:")
    print("  1. Ve a: https://myaccount.google.com/apppasswords")
    print("  2. Selecciona 'Correo' y 'Otro (nombre personalizado)'")
    print("  3. Escribe 'WES Reportes'")
    print("  4. Copia la contraseña de 16 caracteres")
    print()
    print("=" * 60)
    print()
    
    # Solicitar contraseña de aplicación
    smtp_password = input("Ingresa tu CONTRASEÑA DE APLICACIÓN (16 caracteres): ").strip()
    
    if not smtp_password:
        print("[ERROR] Se requiere la contraseña de aplicación")
        return
    
    # Limpiar espacios de la contraseña
    smtp_password = smtp_password.replace(" ", "")
    
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
    print("=" * 60)
    if exito:
        print("[✓] PRUEBA EXITOSA - Correo enviado correctamente")
        print("=" * 60)
        print()
        print(f"Revisa tu bandeja de entrada: {destinatario}")
    else:
        print("[✗] PRUEBA FALLIDA")
        print("=" * 60)
        print()
        print("Posibles causas:")
        print("1. La contraseña de aplicación no es correcta")
        print("2. No has activado la verificación en 2 pasos")
        print("3. La contraseña de aplicación fue revocada")
        print()
        print("Solución:")
        print("1. Verifica que tengas verificación en 2 pasos activada")
        print("2. Genera una nueva contraseña de aplicación en:")
        print("   https://myaccount.google.com/apppasswords")
        print("3. Asegúrate de copiar la contraseña completa (16 caracteres)")

if __name__ == "__main__":
    main()



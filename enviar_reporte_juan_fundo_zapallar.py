"""
Script para enviar reporte de Fundo Zapallar de la última semana a Juan.
"""

import sys
from datetime import datetime, timedelta
from pathlib import Path
import argparse

# Importar funciones necesarias
from lista_contactos_reportes import obtener_contacto
from generar_reporte_word import generate_report, generate_aggregated_report, convertir_word_a_pdf
from monitorear_correos_y_generar_reportes import enviar_correo_personalizado

# Configurar codificación UTF-8
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except AttributeError:
        # Si stdout ya está redirigido, no hacer nada
        pass

def main():
    print("=" * 70)
    print("ENVIANDO REPORTE DE FUNDO ZAPALLAR A JUAN")
    print("=" * 70)
    print()
    
    # Obtener información de Juan
    contacto = obtener_contacto("juan")
    if not contacto:
        print("[ERROR] No se encontró el contacto 'juan'")
        return False
    
    print(f"[INFO] Contacto encontrado: {contacto['nombre_completo']}")
    print(f"[INFO] Email: {contacto['email']}")
    print()
    
    # Calcular fechas de la última semana
    end_date = datetime.now()
    start_date = end_date - timedelta(days=7)
    
    start_date_str = start_date.strftime('%d/%m/%Y')
    end_date_str = end_date.strftime('%d/%m/%Y')
    
    print(f"[INFO] Periodo: {start_date_str} - {end_date_str}")
    print()
    
    # ID de Fundo Zapallar
    empresa_id = "000027"
    empresa_nombre = "Fundo Zapallar"
    
    print(f"[INFO] Empresa: {empresa_nombre} (ID: {empresa_id})")
    print()
    
    # Obtener todos los nodos de Fundo Zapallar
    import requests
    try:
        url = f"http://104.248.53.141:7001/wes/api/acl-entities/v1/companies/{empresa_id}"
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            data = response.json()
            todos_los_nodos = data.get('nodes', [])
            node_ids = [node.get('nodeId') for node in todos_los_nodos]
            print(f"[OK] Se encontraron {len(node_ids)} nodo(s) para {empresa_nombre}")
            for node in todos_los_nodos:
                print(f"  - {node.get('nodeId')}: {node.get('name')}")
        else:
            print(f"[ERROR] No se pudieron obtener nodos: código {response.status_code}")
            return False
    except Exception as e:
        print(f"[ERROR] Error al obtener nodos: {e}")
        return False
    
    if not node_ids:
        print("[ERROR] No se encontraron nodos para procesar")
        return False
    
    print()
    print("[INFO] Generando reportes...")
    print()
    
    archivos_pdf = []
    
    # Generar reportes individuales
    print(f"[INFO] Generando {len(node_ids)} reporte(s) individual(es)...")
    for node_id in node_ids:
        try:
            print(f"  Generando reporte para {node_id}...")
            args = argparse.Namespace(
                company_id=empresa_id,
                node_id=node_id,
                start_date=start_date_str,
                end_date=end_date_str,
                output_dir="reports",
                enviar_correo=False,
                destinatario=None,
                smtp_servidor=None,
                smtp_puerto=None,
                smtp_usuario=None,
                smtp_password=None
            )
            reporte_path = generate_report(args)
            
            # Convertir a PDF
            pdf_path = convertir_word_a_pdf(reporte_path)
            if pdf_path:
                archivos_pdf.append(pdf_path)
                print(f"  [OK] Reporte generado: {pdf_path}")
        except Exception as e:
            print(f"  [ERROR] Error al generar reporte para {node_id}: {e}")
            import traceback
            traceback.print_exc()
    
    # Generar reporte agregado
    print()
    print("[INFO] Generando reporte agregado...")
    try:
        reporte_agregado = generate_aggregated_report(
            company_id=empresa_id,
            node_ids=node_ids,
            start_date=start_date_str,
            end_date=end_date_str
        )
        
        if reporte_agregado:
            # Convertir a PDF
            pdf_path = convertir_word_a_pdf(reporte_agregado)
            if pdf_path:
                archivos_pdf.append(pdf_path)
                print(f"  [OK] Reporte agregado generado: {pdf_path}")
    except Exception as e:
        print(f"  [ERROR] Error al generar reporte agregado: {e}")
        import traceback
        traceback.print_exc()
    
    # Enviar correo con reportes
    if archivos_pdf:
        print()
        print(f"[INFO] Enviando {len(archivos_pdf)} reporte(s) por correo a {contacto['email']}...")
        info_reporte = {
            'empresa': empresa_nombre,
            'periodo': f"{start_date_str} - {end_date_str}",
            'tipo': 'ambos',
            'puntos_monitoreo': [node.get('name') for node in todos_los_nodos]
        }
        
        if enviar_correo_personalizado(contacto, info_reporte, archivos_pdf=archivos_pdf):
            print()
            print("[OK] Reportes enviados correctamente a Juan")
            return True
        else:
            print()
            print("[ERROR] No se pudieron enviar los reportes")
            return False
    else:
        print()
        print("[ERROR] No se generaron reportes para enviar")
        return False

if __name__ == "__main__":
    try:
        success = main()
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\n[INFO] Operación cancelada por el usuario")
        sys.exit(1)
    except Exception as e:
        print(f"\n[ERROR] Error crítico: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

"""
Script para generar reportes de Providencia de diciembre y enviarlos por correo.
"""

import sys
from pathlib import Path
from datetime import datetime
import requests

# Agregar el directorio actual al path para importar módulos
sys.path.insert(0, str(Path(__file__).parent))

from monitorear_correos_y_generar_reportes import (
    obtener_empresa_id_por_nombre,
    obtener_todas_las_empresas,
    enviar_correo_personalizado
)
from generar_reporte_word import generate_report, generate_aggregated_report, convertir_word_a_pdf
from lista_contactos_reportes import obtener_contacto_por_email

# Configurar codificación UTF-8
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except AttributeError:
        # stdout ya está interceptado (por ejemplo, por TeeOutput)
        pass

def obtener_nodos_empresa(company_id):
    """Obtiene todos los nodos de una empresa."""
    try:
        url = f"http://104.248.53.141:7001/wes/api/acl-entities/v1/companies/{company_id}"
        response = requests.get(url, timeout=10)
        
        if response.status_code == 200:
            empresa_data = response.json()
            nodes = empresa_data.get('nodes', [])
            node_ids = [node.get('nodeId') for node in nodes if node.get('nodeId')]
            return node_ids, empresa_data.get('name', '')
        else:
            print(f"[ERROR] Error al obtener nodos: código {response.status_code}")
            return [], ""
    except Exception as e:
        print(f"[ERROR] Error al consultar API: {e}")
        return [], ""

def main():
    print("=" * 60)
    print("GENERACIÓN DE REPORTES - PROVIDENCIA - DICIEMBRE 2025")
    print("=" * 60)
    print()
    
    # 1. Buscar empresa Providencia
    print("[1/5] Buscando empresa Providencia...")
    empresas_dict = obtener_todas_las_empresas()
    empresa_id = obtener_empresa_id_por_nombre("Providencia", empresas_dict)
    
    if not empresa_id:
        print("[ERROR] No se encontró la empresa Providencia")
        return False
    
    print(f"  [OK] Empresa encontrada: ID {empresa_id}")
    
    # 2. Obtener todos los nodos
    print(f"\n[2/5] Obteniendo nodos de la empresa {empresa_id}...")
    node_ids, empresa_nombre = obtener_nodos_empresa(empresa_id)
    
    if not node_ids:
        print("[ERROR] No se encontraron nodos para la empresa")
        return False
    
    print(f"  [OK] Encontrados {len(node_ids)} nodo(s):")
    for node_id in node_ids:
        print(f"    - {node_id}")
    
    # 3. Fechas de diciembre 2025
    start_date = "2025-12-01"
    end_date = "2025-12-31"
    print(f"\n[3/5] Periodo: {start_date} a {end_date}")
    
    # 4. Generar reportes individuales
    print(f"\n[4/5] Generando reportes individuales...")
    archivos_pdf = []
    archivos_word = []
    
    for i, node_id in enumerate(node_ids, 1):
        print(f"  [{i}/{len(node_ids)}] Generando reporte para {node_id}...")
        try:
            import argparse
            args = argparse.Namespace(
                company_id=empresa_id,
                node_id=node_id,
                start_date=start_date,
                end_date=end_date,
                output_dir="reports",
                enviar_correo=False,
                destinatario=None,
                smtp_servidor=None,
                smtp_puerto=None,
                smtp_usuario=None,
                smtp_password=None
            )
            reporte_path = generate_report(args)
            
            if reporte_path and Path(reporte_path).exists():
                # Convertir a PDF
                pdf_path = convertir_word_a_pdf(reporte_path)
                if pdf_path:
                    archivos_pdf.append(pdf_path)
                    print(f"    [OK] Reporte PDF generado: {Path(pdf_path).name}")
                else:
                    print(f"    [ADVERTENCIA] No se pudo convertir a PDF, usando Word")
                    archivos_word.append(reporte_path)
            else:
                print(f"    [ERROR] No se generó el reporte para {node_id}")
        except Exception as e:
            print(f"    [ERROR] Error al generar reporte para {node_id}: {e}")
            import traceback
            traceback.print_exc()
    
    # 5. Generar reporte agregado (si hay más de un nodo)
    if len(node_ids) > 1:
        print(f"\n[5/5] Generando reporte agregado...")
        try:
            reporte_agregado = generate_aggregated_report(
                company_id=empresa_id,
                node_ids=node_ids,
                start_date=start_date,
                end_date=end_date
            )
            
            if reporte_agregado and Path(reporte_agregado).exists():
                # Convertir a PDF
                pdf_path = convertir_word_a_pdf(reporte_agregado)
                if pdf_path:
                    archivos_pdf.append(pdf_path)
                    print(f"  [OK] Reporte agregado PDF generado: {Path(pdf_path).name}")
                else:
                    print(f"  [ADVERTENCIA] No se pudo convertir a PDF, usando Word")
                    archivos_word.append(reporte_agregado)
            else:
                print(f"  [ERROR] No se generó el reporte agregado")
        except Exception as e:
            print(f"  [ERROR] Error al generar reporte agregado: {e}")
            import traceback
            traceback.print_exc()
    else:
        print(f"\n[5/5] Solo hay 1 nodo, no se genera reporte agregado")
    
    # 6. Enviar por correo
    print(f"\n[6/6] Enviando reportes por correo...")
    email_destino = "joseotarola@wes.cl"
    
    # Crear contacto temporal
    contacto = {
        "email": email_destino,
        "tratamiento": "Estimado Jose",
        "despedida": "Quedo atento a tus comentarios.",
        "nombre_completo": "Jose Otarola"
    }
    
    info_reporte = {
        'empresa': empresa_nombre or "Providencia",
        'periodo': f"{start_date} - {end_date}",
        'tipo_reporte': 'ambos' if len(node_ids) > 1 else 'individual',
        'formato': 'pdf'
    }
    
    # Enviar correo con PDFs
    if archivos_pdf:
        print(f"  Enviando {len(archivos_pdf)} reporte(s) PDF a {email_destino}...")
        if enviar_correo_personalizado(contacto, info_reporte, archivos_pdf=archivos_pdf):
            print(f"  [OK] Correo enviado correctamente")
        else:
            print(f"  [ERROR] No se pudo enviar el correo")
            return False
    elif archivos_word:
        print(f"  Enviando {len(archivos_word)} reporte(s) Word a {email_destino}...")
        if enviar_correo_personalizado(contacto, info_reporte, archivos_word=archivos_word):
            print(f"  [OK] Correo enviado correctamente")
        else:
            print(f"  [ERROR] No se pudo enviar el correo")
            return False
    else:
        print(f"  [ERROR] No hay reportes para enviar")
        return False
    
    print()
    print("=" * 60)
    print("PROCESO COMPLETADO EXITOSAMENTE")
    print("=" * 60)
    print(f"Reportes generados: {len(archivos_pdf) + len(archivos_word)}")
    print(f"Enviados a: {email_destino}")
    print()
    
    return True

if __name__ == "__main__":
    try:
        success = main()
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\n[INFO] Proceso cancelado por el usuario")
        sys.exit(1)
    except Exception as e:
        print(f"\n[ERROR] Error inesperado: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


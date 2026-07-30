"""Enviar el reporte de Fundo Zapallar - Matriz ESVAL (diciembre) a Diego."""

from pathlib import Path
from generar_reporte_word import enviar_reporte_por_correo, get_company_name, get_node_name

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_PASSWORD = "vxbynfpoehbweelj"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

COMPANY_ID = "000027"  # Fundo Zapallar
NODE_ID = "000027-01"  # Matriz ESVAL
DESTINATARIO = "diegocarrasco@wes.cl"

START_DATE = "01/12/2025"
END_DATE = "31/12/2025"


def main() -> None:
    # Buscar el reporte más reciente
    base_dir = Path("reports") / "Fundo_Zapallar" / "REPORTE"
    
    if not base_dir.exists():
        print("[ERROR] No se encontró la carpeta de reportes de Fundo Zapallar.")
        return
    
    # Buscar el reporte específico de Matriz ESVAL de diciembre
    patron = f"Reporte_{COMPANY_ID}_{NODE_ID}_20251201_20251231.docx"
    candidatos = list(base_dir.rglob(patron))
    
    if not candidatos:
        print(f"[ERROR] No se encontró el reporte: {patron}")
        print(f"Buscando en: {base_dir}")
        return
    
    # Tomar el más reciente si hay múltiples
    reporte_path = max(candidatos, key=lambda p: p.stat().st_mtime)
    
    print(f"Reporte encontrado: {reporte_path}")
    
    # Obtener nombres
    company_name = get_company_name(COMPANY_ID)
    node_name = get_node_name(NODE_ID)
    
    print(f"Empresa: {company_name}")
    print(f"Nodo: {node_name}")
    print(f"Enviando a: {DESTINATARIO}")
    
    # Enviar correo
    success = enviar_reporte_por_correo(
        reporte_path=reporte_path,
        destinatario=DESTINATARIO,
        smtp_servidor=SMTP_SERVIDOR,
        smtp_puerto=SMTP_PUERTO,
        smtp_usuario=SMTP_USUARIO,
        smtp_password=SMTP_PASSWORD,
        company_name=company_name,
        node_name=node_name,
        start_date=START_DATE,
        end_date=END_DATE,
    )
    
    if success:
        print(f"\n[OK] Reporte enviado exitosamente a {DESTINATARIO}")
    else:
        print(f"\n[ERROR] No se pudo enviar el reporte")


if __name__ == "__main__":
    main()













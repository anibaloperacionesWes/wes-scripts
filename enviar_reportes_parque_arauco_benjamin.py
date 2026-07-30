"""Enviar todos los reportes de Parque Arauco Kennedy existentes a Benjamín en PDF."""

from pathlib import Path

from generar_reporte_word import (
    enviar_reporte_por_correo,
    get_company_name,
    get_node_name,
)


SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_PASSWORD = "vxbynfpoehbweelj"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

DESTINATARIO = "benjamingumucio@wes.cl"
COMPANY_ID = "000025"


def main() -> None:
    base_dir = Path("reports") / "Parque_Arauco"
    report_dir = base_dir / "REPORTE"
    agregado_dir = base_dir / "ABREGADO"

    if not report_dir.exists():
        print("[ERROR] No se encontró la carpeta de reportes individuales de Parque Arauco.")
        return

    company_name = get_company_name(COMPANY_ID)

    # Enviar reportes individuales (uno por subcarpeta)
    docx_paths = sorted(report_dir.rglob("Reporte_*.docx"))
    print(f"Encontrados {len(docx_paths)} reportes individuales para enviar.")

    enviados_ok = 0
    for i, reporte_path in enumerate(docx_paths, start=1):
        # Inferir node_id desde el nombre del archivo (formato Reporte_company_node_start_end.docx)
        parts = reporte_path.stem.split("_")
        node_id = parts[2] if len(parts) >= 3 else None
        node_name = get_node_name(node_id) if node_id else None

        print(f"[{i}/{len(docx_paths)}] Enviando reporte de {node_name or 'desconocido'}...")

        exito = enviar_reporte_por_correo(
            reporte_path=reporte_path,
            destinatario=DESTINATARIO,
            smtp_servidor=SMTP_SERVIDOR,
            smtp_puerto=SMTP_PUERTO,
            smtp_usuario=SMTP_USUARIO,
            smtp_password=SMTP_PASSWORD,
            company_name=company_name,
            node_name=node_name,
            start_date="07-12-25",
            end_date="14-12-25",
        )

        if exito:
            enviados_ok += 1
            print("  [OK] Enviado correctamente")
        else:
            print("  [ERROR] Fallo al enviar este reporte")

    # Enviar reporte agregado si existe
    agregado_docx = None
    if agregado_dir.exists():
        candidatos = sorted(agregado_dir.rglob("Reporte_Agregado_*.docx"))
        if candidatos:
            agregado_docx = candidatos[-1]

    if agregado_docx:
        print(f"Enviando reporte agregado: {agregado_docx.name}")
        exito = enviar_reporte_por_correo(
            reporte_path=agregado_docx,
            destinatario=DESTINATARIO,
            smtp_servidor=SMTP_SERVIDOR,
            smtp_puerto=SMTP_PUERTO,
            smtp_usuario=SMTP_USUARIO,
            smtp_password=SMTP_PASSWORD,
            company_name=company_name,
            node_name=None,
            start_date="07-12-25",
            end_date="14-12-25",
        )
        if exito:
            print("[OK] Reporte agregado enviado correctamente")
        else:
            print("[ERROR] Fallo al enviar el reporte agregado")

    print()
    print("Resumen de envío:")
    print(f"  - Reportes individuales enviados OK: {enviados_ok}/{len(docx_paths)}")
    print(f"  - Reporte agregado enviado: {'Sí' if agregado_docx else 'No encontrado'}")


if __name__ == "__main__":
    main()















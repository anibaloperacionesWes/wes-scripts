"""Enviar solo el reporte agregado de Parque Arauco Kennedy como prueba."""

from pathlib import Path

from generar_reporte_word import enviar_reporte_por_correo, get_company_name

SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_PASSWORD = "vxbynfpoehbweelj"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

DESTINATARIO = "agente.ia@wes.cl"
COMPANY_ID = "000025"


def main() -> None:
    base_dir = Path("reports") / "Parque_Arauco" / "ABREGADO"

    if not base_dir.exists():
        print("[ERROR] No se encontró la carpeta de reportes agregados de Parque Arauco.")
        return

    # Tomar el último reporte agregado disponible
    candidatos = sorted(base_dir.rglob("Reporte_Agregado_*.docx"))
    if not candidatos:
        print("[ERROR] No se encontró ningún Reporte_Agregado_*.docx.")
        return

    reporte_path = candidatos[-1]
    company_name = get_company_name(COMPANY_ID)

    print(f"Enviando como prueba el reporte agregado: {reporte_path}")

    exito = enviar_reporte_por_correo(
        reporte_path=reporte_path,
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
        print("[OK] Correo de prueba enviado correctamente.")
    else:
        print("[ERROR] Falló el envío del correo de prueba.")


if __name__ == "__main__":
    main()















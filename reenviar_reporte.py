"""Script para reenviar el reporte agregado más reciente."""

import sys
from pathlib import Path
from generar_reporte_word import enviar_reporte_por_correo, get_company_name

# Configurar codificación UTF-8
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

# Configuración
DESTINATARIO = "silvanaaraya.roja@gmail.com"
SMTP_USUARIO = "agente.ia@wes.cl"
SMTP_PASSWORD = "vxbynfpoehbweelj"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PUERTO = 587

# Buscar el reporte agregado más reciente en toda la carpeta reports
reports_dir = Path("reports")
reportes_agregados = list(reports_dir.rglob("*.docx"))
# Filtrar solo los agregados
reportes_agregados = [r for r in reportes_agregados if "Agregado" in r.name or "AGREGADO" in str(r.parent)]

if not reportes_agregados:
    print("[ERROR] No se encontraron reportes agregados")
    print(f"Buscando en: {reports_dir.absolute()}")
    # Listar todos los .docx encontrados para debug
    todos = list(reports_dir.rglob("*.docx"))
    print(f"Total archivos .docx encontrados: {len(todos)}")
    if todos:
        print("Archivos encontrados:")
        for f in todos[:5]:
            print(f"  - {f}")
    sys.exit(1)

# Ordenar por fecha de modificación (más reciente primero)
reportes_agregados.sort(key=lambda x: x.stat().st_mtime, reverse=True)
reporte = reportes_agregados[0]

print("=" * 60)
print("REENVIANDO REPORTE AGREGADO")
print("=" * 60)
print(f"Reporte: {reporte.name}")
print(f"Destinatario: {DESTINATARIO}")
print()
print("Enviando correo...")
print()

exito = enviar_reporte_por_correo(
    reporte_path=reporte,
    destinatario=DESTINATARIO,
    smtp_servidor=SMTP_SERVIDOR,
    smtp_puerto=SMTP_PUERTO,
    smtp_usuario=SMTP_USUARIO,
    smtp_password=SMTP_PASSWORD,
    company_name="BUPA",
    node_name=None,
    start_date="01-11-25",
    end_date="30-11-25",
)

print()
if exito:
    print("=" * 60)
    print("[OK] CORREO ENVIADO EXITOSAMENTE")
    print("=" * 60)
    print(f"Destinatario: {DESTINATARIO}")
    print(f"Reporte: {reporte.name}")
else:
    print("=" * 60)
    print("[ERROR] FALLO EL ENVÍO DEL CORREO")
    print("=" * 60)


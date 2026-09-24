#!/usr/bin/env bash
# Sube Formulario.html al Apps Script del acta WES (requiere PC + clasp login).
set -euo pipefail
cd "$(dirname "$0")"
echo "Proyecto: $(python3 -c 'import json;print(json.load(open(".clasp.json"))["scriptId"])')"
npx --yes @google/clasp push -f
echo
echo "OK push. Ahora en Apps Script:"
echo "  Implementar → Administrar implementaciones → lápiz del /exec → Nueva versión → Implementar"
echo "Verificación en el form: etiqueta WES-build-21i bajo Tecnología."

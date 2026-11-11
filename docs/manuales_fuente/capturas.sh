#!/usr/bin/env bash
# Capturas del manual de usuario de la Consola Registro Info SAP, con datos inventados y sin red (ver README.md).
# La idea y el desarrollo de la aplicación, incluidas sus pruebas, son de Víctor Alejandro Martín Andrade,
# para el Departamento de Central de Compras de Princess Hotels en Canarias.
# Uso: docs/manuales_fuente/capturas.sh [carpeta_salida]     (por defecto docs/manuales_fuente/cap)
# Variables: PYTHON (intérprete con tkinter y los paquetes de compilacion/requisitos.txt), SOLO=prefijo, VOLCAR=1
# Requisitos: Xvfb (xvfb-run), ImageMagick (import) y la fuente Inter.
set -euo pipefail
AQUI="$(cd "$(dirname "$0")" && pwd)"
RAIZ="$(cd "$AQUI/../.." && pwd)"
SALIDA="${1:-$AQUI/cap}"
PY="${PYTHON:-python3}"
mkdir -p "$SALIDA"
timeout "${TIEMPO:-300}" xvfb-run -a -s "-screen 0 1400x900x24" "$PY" "$AQUI/capturas.py" "$RAIZ" "$SALIDA" 2>&1 | tee "$SALIDA/salida.log" | grep -E "✓|⚠|✗|FIN" || true
echo "Capturas en: $SALIDA"

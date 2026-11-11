#!/bin/bash
# Genera el manual de usuario de la Consola Registro Info SAP (docx + pdf) a partir de las capturas de ./cap,
# y el índice que usa «Consultar el manual» dentro de la app.
# La idea y el desarrollo de la aplicación, incluidas sus pruebas, son de Víctor Alejandro Martín Andrade,
# para el Departamento de Central de Compras de Princess Hotels en Canarias.
# Uso: ./generar.sh      (necesita Node con el paquete «docx», LibreOffice y poppler-utils)
set -e
cd "$(dirname "$0")"
NOMBRE=Manual_Usuario_ConsolaRegistroInfo
mkdir -p salida
node manual_usuario.js "salida/$NOMBRE.docx"
(cd salida && zip -d "$NOMBRE.docx" '*/' >/dev/null 2>&1 || true)
soffice --headless --convert-to pdf --outdir salida "salida/$NOMBRE.docx" >/dev/null
python3 paginar_indice.py salida/manual_indice.json "salida/$NOMBRE.pdf"
echo "Listo: salida/$NOMBRE.docx, .pdf y manual_indice.json"
echo "Revisarlos y copiar: el .docx y el .pdf a docs/, y el .pdf y manual_indice.json a assets/ (van dentro del .exe)."

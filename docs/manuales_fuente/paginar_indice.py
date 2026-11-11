# -*- coding: utf-8 -*-
"""Añade a manual_indice.json la página del PDF en la que empieza cada apartado, para que
«Consultar el manual» de la app diga «página N». Lo llama generar.sh.

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación son de
Víctor Alejandro Martín Andrade (VAMA).
Uso: python paginar_indice.py salida/manual_indice.json salida/Manual.pdf
"""
import json
import re
import subprocess
import sys
import unicodedata


def norm(t):
    t = unicodedata.normalize("NFKD", t.lower())
    t = "".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9 ]", " ", t)).strip()


def paginas(pdf):
    total = int(re.search(r"Pages:\s+(\d+)", subprocess.run(["pdfinfo", pdf], capture_output=True, text=True).stdout).group(1))
    return [norm(subprocess.run(["pdftotext", "-layout", "-f", str(i), "-l", str(i), pdf, "-"],
                                capture_output=True, text=True).stdout) for i in range(1, total + 1)]


def main(ruta_json, pdf):
    with open(ruta_json, encoding="utf-8") as f:
        datos = json.load(f)
    pags = paginas(pdf)
    desde = 3                      # portada e índice no cuentan
    sin = []
    for s in datos["secciones"]:
        t = norm(s["titulo"])
        pag = next((i + 1 for i in range(desde - 1, len(pags)) if t and t in pags[i]), None)
        if pag:
            s["pagina"] = pag
            desde = pag
        else:
            sin.append(s["titulo"])
    with open(ruta_json, "w", encoding="utf-8") as f:
        json.dump(datos, f, ensure_ascii=False, indent=1)
    print(f"Páginas puestas: {len(datos['secciones']) - len(sin)} de {len(datos['secciones'])}")
    if sin:
        print("Sin página:", sin)
        sys.exit(1)


if __name__ == "__main__":
    main(*sys.argv[1:3])

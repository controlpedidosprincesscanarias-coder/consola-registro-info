# -*- coding: utf-8 -*-
"""
generar_guia_sap.py — dibuja las imágenes de la guía «Descargar Productos.xlsx y Listas de productos.xlsx de SAP».

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).

Esquemas fieles a las pantallas que mandó Víctor (13 nov 2026): los dos iconos de las aplicaciones de SAP, la barra
de la tabla con el botón ⌄ de exportar (el último de la derecha) y la ventana «Exportar como». Se dibujan aquí, sin
capturas reales, y van en assets/ (dentro del .exe). Si se prefieren capturas reales, basta con dejar en assets/ unos
PNG con el mismo nombre.

Uso: python compilacion/generar_guia_sap.py
"""
import os
import sys

from PIL import Image, ImageDraw, ImageFont

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ASSETS = os.path.join(RAIZ, "assets")
E = 2                                   # escala (nítido en pantallas con zoom)
AZUL, AZUL_OSC, GRIS, GRIS_CL, TEXTO = (10, 110, 209), (0, 70, 160), (137, 145, 154), (234, 237, 240), (29, 45, 62)
NARANJA = (232, 89, 12)


def _fuente(tam, negrita=False):
    candidatas = (["segoeuib.ttf", "DejaVuSans-Bold.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"]
                  if negrita else
                  ["segoeui.ttf", "DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"])
    for c in candidatas:
        try:
            return ImageFont.truetype(c, tam * E)
        except OSError:
            continue
    return ImageFont.load_default()


def _marca(d, x, y, n):
    r = 13 * E
    d.ellipse((x - r - 2 * E, y - r - 2 * E, x + r + 2 * E, y + r + 2 * E), fill="white")
    d.ellipse((x - r, y - r, x + r, y + r), fill=NARANJA)
    d.text((x, y), str(n), font=_fuente(13, True), fill="white", anchor="mm")


def _recuadro(d, caja, grosor=3):
    x0, y0, x1, y1 = caja
    d.rounded_rectangle((x0 - 5 * E, y0 - 5 * E, x1 + 5 * E, y1 + 5 * E), radius=7 * E, outline=NARANJA,
                        width=grosor * E)


def _icono(d, x, y):
    """Icono de aplicación (cuadrado gris con un cubo), como los de la página de inicio de SAP."""
    s = 34 * E
    d.rounded_rectangle((x, y, x + s, y + s), radius=7 * E, fill=(98, 114, 132))
    cx, cy, r = x + s / 2, y + s / 2, 9 * E
    pts = [(cx, cy - r), (cx + r, cy - r / 2), (cx + r, cy + r / 2), (cx, cy + r), (cx - r, cy + r / 2),
           (cx - r, cy - r / 2)]
    d.polygon(pts, outline="white", width=2 * E)
    d.line((cx, cy, cx, cy + r), fill="white", width=2 * E)
    d.line((cx - r, cy - r / 2, cx, cy), fill="white", width=2 * E)
    d.line((cx + r, cy - r / 2, cx, cy), fill="white", width=2 * E)


def apps():
    img = Image.new("RGB", (620 * E, 110 * E), "white")
    d = ImageDraw.Draw(img)
    f = _fuente(13, True)
    _icono(d, 30 * E, 38 * E)
    d.text((80 * E, 55 * E), "Lista de productos", font=f, fill=TEXTO, anchor="lm")
    _icono(d, 320 * E, 38 * E)
    d.text((370 * E, 46 * E), "Gestionar datos maestros de", font=f, fill=TEXTO, anchor="lm")
    d.text((370 * E, 66 * E), "producto", font=f, fill=TEXTO, anchor="lm")
    fp = _fuente(10)
    d.text((80 * E, 80 * E), "→ Listas de productos.xlsx", font=fp, fill=AZUL_OSC, anchor="lm")
    d.text((370 * E, 88 * E), "→ Productos.xlsx", font=fp, fill=AZUL_OSC, anchor="lm")
    img.save(os.path.join(ASSETS, "guia_sap_apps.png"))


def barra():
    W, H = 760 * E, 124 * E
    img = Image.new("RGB", (W, H), (245, 246, 247))
    d = ImageDraw.Draw(img)
    d.rectangle((0, 0, W, 40 * E), fill="white")
    d.text((20 * E, 20 * E), "COMPRAS CANARIAS  ⌄", font=_fuente(13), fill=AZUL, anchor="lm")
    d.line((0, 40 * E, W, 40 * E), fill=(214, 219, 224), width=E)
    d.text((24 * E, 68 * E), "Productos (8.850)", font=_fuente(13), fill=TEXTO, anchor="lm")
    y0, y1 = 56 * E, 80 * E
    x = W - 30 * E
    # botones de la derecha: copiar (gris), ⚙, exportar y ⌄ (el último)
    botones = []
    for ancho in (28, 28, 28, 28):
        botones.append((x - ancho * E, y0, x, y1))
        x -= (ancho + 8) * E
    botones = botones[::-1]
    estilos = [(GRIS_CL, GRIS), ("white", AZUL), ("white", AZUL), ("white", AZUL)]
    for (bx0, by0, bx1, by1), (fondo, borde) in zip(botones, estilos):
        d.rounded_rectangle((bx0, by0, bx1, by1), radius=4 * E, fill=fondo, outline=borde, width=E)
    # iconos
    b = botones
    d.rectangle((b[0][0] + 9 * E, b[0][1] + 7 * E, b[0][2] - 9 * E, b[0][3] - 7 * E), outline=GRIS, width=E)
    d.ellipse((b[1][0] + 8 * E, b[1][1] + 6 * E, b[1][2] - 8 * E, b[1][3] - 6 * E), outline=AZUL, width=2 * E)
    d.rectangle((b[2][0] + 8 * E, b[2][1] + 7 * E, b[2][2] - 8 * E, b[2][3] - 7 * E), outline=AZUL, width=E)
    d.line((b[2][0] + 8 * E, (b[2][1] + b[2][3]) / 2, b[2][2] - 8 * E, (b[2][1] + b[2][3]) / 2), fill=AZUL, width=E)
    cx, cy = (b[3][0] + b[3][2]) / 2, (b[3][1] + b[3][3]) / 2
    d.line((cx - 6 * E, cy - 3 * E, cx, cy + 3 * E, cx + 6 * E, cy - 3 * E), fill=AZUL, width=2 * E)
    _recuadro(d, b[3])
    _marca(d, cx, y1 + 24 * E, 1)
    img.save(os.path.join(ASSETS, "guia_sap_barra.png"))


def exportar(nombre, archivo):
    W, H = 400 * E, 300 * E
    img = Image.new("RGB", (W, H), (120, 124, 128))
    d = ImageDraw.Draw(img)
    x0, y0, x1, y1 = 16 * E, 14 * E, W - 16 * E, H - 14 * E
    d.rounded_rectangle((x0, y0, x1, y1), radius=8 * E, fill="white")
    f, fb = _fuente(12), _fuente(14)
    d.text((x0 + 18 * E, y0 + 24 * E), "Exportar como", font=fb, fill=TEXTO, anchor="lm")
    d.line((x0, y0 + 46 * E, x1, y0 + 46 * E), fill=(214, 219, 224), width=E)
    d.text((x0 + 18 * E, y0 + 64 * E), "Nombre de fichero", font=f, fill=TEXTO, anchor="lm")
    caja_nombre = (x0 + 18 * E, y0 + 76 * E, x1 - 18 * E, y0 + 100 * E)
    d.rounded_rectangle(caja_nombre, radius=3 * E, outline=(137, 145, 154), width=E)
    d.text((caja_nombre[0] + 8 * E, (caja_nombre[1] + caja_nombre[3]) / 2), nombre, font=f, fill=TEXTO, anchor="lm")
    d.text((x0 + 18 * E, y0 + 116 * E), "Formato", font=f, fill=TEXTO, anchor="lm")
    caja_fmt = (x0 + 18 * E, y0 + 128 * E, x1 - 18 * E, y0 + 152 * E)
    d.rounded_rectangle(caja_fmt, radius=3 * E, outline=(137, 145, 154), width=E)
    d.text((caja_fmt[0] + 8 * E, (caja_fmt[1] + caja_fmt[3]) / 2), "Microsoft Excel (*.xlsx)", font=f, fill=TEXTO,
           anchor="lm")
    for i, (txt, marcado) in enumerate((("Incluir opciones de filtro", False), ("Dividir celdas con varios valores", True))):
        cy = y0 + 172 * E + i * 26 * E
        cb = (x0 + 22 * E, cy - 7 * E, x0 + 36 * E, cy + 7 * E)
        d.rounded_rectangle(cb, radius=2 * E, outline=(137, 145, 154), width=E)
        if marcado:
            d.line((cb[0] + 3 * E, cy, cb[0] + 6 * E, cy + 4 * E, cb[2] - 2 * E, cy - 5 * E), fill=AZUL_OSC, width=2 * E)
        d.text((x0 + 46 * E, cy), txt, font=f, fill=TEXTO, anchor="lm")
    d.line((x0, y1 - 44 * E, x1, y1 - 44 * E), fill=(214, 219, 224), width=E)
    btn = (x1 - 190 * E, y1 - 34 * E, x1 - 100 * E, y1 - 10 * E)
    d.rounded_rectangle(btn, radius=4 * E, fill=AZUL)
    d.text(((btn[0] + btn[2]) / 2, (btn[1] + btn[3]) / 2), "Exportar", font=_fuente(12, True), fill="white",
           anchor="mm")
    d.text((x1 - 50 * E, (btn[1] + btn[3]) / 2), "Cancelar", font=f, fill=AZUL, anchor="mm")
    _recuadro(d, caja_nombre, 2)
    _marca(d, caja_nombre[2] - 4 * E, caja_nombre[1] - 2 * E, 2)
    _recuadro(d, btn, 2)
    _marca(d, btn[0] - 20 * E, (btn[1] + btn[3]) / 2, 3)
    img.save(os.path.join(ASSETS, archivo))


if __name__ == "__main__":
    os.makedirs(ASSETS, exist_ok=True)
    apps()
    barra()
    exportar("Productos", "guia_sap_exportar_productos.png")
    exportar("Listas de productos", "guia_sap_exportar_listas.png")
    print("Imágenes de la guía en", ASSETS, file=sys.stderr)

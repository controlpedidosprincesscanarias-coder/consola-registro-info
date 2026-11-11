# -*- coding: utf-8 -*-
"""
generar_icono.py
Genera assets/logo.ico (icono del .exe, de las ventanas y de la barra de tareas) a partir de la
corona de compilacion/corona_icono.png: corona dorada sobre un cuadrado azul marino con borde dorado, para
distinguirla del icono del Organizador (la corona sola). Cada tamaño se dibuja aparte (a 8x y
reducido) para que se vea nítido también a 16 y 32 píxeles.

Uso: python compilacion/generar_icono.py

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación son de
Víctor Alejandro Martín Andrade (VAMA).
"""
import os

from PIL import Image, ImageDraw

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CORONA = os.path.join(RAIZ, "compilacion", "corona_icono.png")
SALIDA = os.path.join(RAIZ, "assets", "logo.ico")
AZUL = (11, 31, 58, 255)        # azul marino de la cabecera de la app
DORADO = (201, 168, 76, 255)    # dorado Princess
TAMANOS = (256, 128, 64, 48, 32, 24, 16)


def dibujar(lado, corona):
    f = 8
    L = lado * f
    im = Image.new("RGBA", (L, L), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    # En pequeño, el borde y el margen pesan más para que se sigan viendo.
    borde = max(1.2, lado * 0.05) * f
    margen = max(0.5, lado * 0.02) * f
    radio = lado * 0.2 * f
    d.rounded_rectangle((margen, margen, L - margen, L - margen), radius=radio, fill=AZUL,
                        outline=DORADO, width=round(borde))
    # Corona dorada centrada, más grande en los tamaños pequeños
    ocupa = 0.66 if lado >= 48 else 0.74
    alto = int(L * ocupa)
    ancho = int(corona.width * alto / corona.height)
    if ancho > L * ocupa:
        ancho = int(L * ocupa)
        alto = int(corona.height * ancho / corona.width)
    mascara = corona.resize((ancho, alto), Image.LANCZOS)
    color = Image.new("RGBA", (ancho, alto), DORADO)
    im.paste(color, ((L - ancho) // 2, (L - alto) // 2 + int(L * 0.01)), mascara)
    return im.resize((lado, lado), Image.LANCZOS)


def main():
    corona = Image.open(CORONA).convert("RGBA").getchannel("A")
    imagenes = [dibujar(t, corona) for t in TAMANOS]
    imagenes[0].save(SALIDA, format="ICO", sizes=[(t, t) for t in TAMANOS],
                     append_images=imagenes[1:])
    print(f"Icono generado: {SALIDA} ({', '.join(str(t) for t in TAMANOS)} px)")


if __name__ == "__main__":
    main()

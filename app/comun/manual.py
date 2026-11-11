# -*- coding: utf-8 -*-
"""
manual.py
Manual de usuario dentro de la app (nov 2026): abrir el PDF y «Consultar el manual».

- assets/Manual_Usuario_ConsolaRegistroInfo.pdf: el manual, dentro del .exe (sirve sin red y es
  siempre el de la versión instalada). abrir_manual() lo copia a una carpeta temporal y lo abre con
  el visor de PDF del equipo.
- assets/manual_indice.json: los apartados del manual (capítulo, título, página y texto), que genera
  docs/manuales_fuente/generar.sh junto con el PDF. buscar() devuelve los apartados que mejor encajan
  con lo que se escribe: sin Internet, sin claves y siempre con el texto del propio manual.

Sin tkinter: se prueba con tests/test_manual.py.
Central Compras Princess Canarias · La idea y el desarrollo de la aplicación son de
Víctor Alejandro Martín Andrade (VAMA).
"""
import json
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata

from app.comun.tarifa_constants import resource_path

NOMBRE_PDF = "Manual_Usuario_ConsolaRegistroInfo.pdf"
RUTA_PDF = "assets/" + NOMBRE_PDF
RUTA_INDICE = "assets/manual_indice.json"

# Palabras que no ayudan a encontrar un apartado
_VACIAS = set("""
a al algo alguna alguno algunos ante antes aqui asi aun cada como con contra cual cuales cuando de
del desde donde dos durante e el ella ellos en entre era es esa ese eso esta este esto estos fue ha
hace hacer hay hasta la las le les lo los mas me mi mis muy ni no nos o otra otro para pero poco por
porque puede pueden puedo que quiero se sea ser si sin sobre son su sus tambien te tengo tiene toda
todo todos tu un una uno unos y ya yo hago debo ver sale salen
""".split())


def normalizar(texto):
    """Minúsculas y sin acentos («Código» → «codigo»)."""
    t = unicodedata.normalize("NFKD", str(texto or "").lower())
    return "".join(c for c in t if not unicodedata.combining(c))


def raiz(palabra):
    """Raíz muy simple para que «tarifas», «tarifa» y «tarifario», o «creo», «crea» y «crear», se
    encuentren entre sí."""
    if len(palabra) > 5:
        return palabra[:5]
    if palabra.isdigit():
        return palabra
    if len(palabra) > 3 and palabra.endswith("s"):
        palabra = palabra[:-1]
    for fin in ("ar", "er", "ir", "a", "e", "o"):
        if palabra.endswith(fin) and len(palabra) - len(fin) >= 3:
            return palabra[:-len(fin)]
    return palabra


def palabras(texto):
    """Raíces de las palabras útiles del texto (también números: códigos, acciones…)."""
    return [raiz(p) for p in re.findall(r"[a-z0-9]+", normalizar(texto)) if p not in _VACIAS and (len(p) > 1 or p.isdigit())]


def cargar_indice(ruta=None):
    """Lista de apartados [{capitulo, titulo, pagina, texto}] o [] si no está el índice."""
    ruta = ruta or resource_path(RUTA_INDICE)
    try:
        with open(ruta, encoding="utf-8") as f:
            datos = json.load(f)
    except (OSError, ValueError):
        return []
    secciones = datos.get("secciones", datos) if isinstance(datos, dict) else datos
    return [s for s in secciones if isinstance(s, dict) and s.get("titulo")]


def _fragmento(texto, raices, largo=260):
    """La frase del apartado que más palabras de la consulta contiene, recortada."""
    frases = [f.strip() for f in re.split(r"(?<=[.:;!?])\s+|\n+", texto or "") if f.strip()]
    if not frases:
        return ""
    # Una frase o dos seguidas: la que recoja más palabras de la consulta (a igualdad, la más corta
    # y la primera).
    candidatos = [(i, f) for i, f in enumerate(frases)]
    candidatos += [(i, frases[i] + " " + frases[i + 1]) for i in range(len(frases) - 1)]
    mejor = max(candidatos, key=lambda c: (len(set(palabras(c[1])) & raices), -len(c[1]) // 120, -c[0]))[1]
    return mejor if len(mejor) <= largo else mejor[:largo].rsplit(" ", 1)[0] + "…"


def buscar(indice, consulta, maximo=5):
    """Los apartados que mejor responden a `consulta`, de más a menos: [{capitulo, titulo, pagina,
    fragmento, puntos}]. Cuenta más una palabra en el título que en el texto, y más las palabras raras
    (que salen en pocos apartados) que las que salen en todos."""
    raices = set(palabras(consulta))
    if not raices or not indice:
        return []
    docs = [(s, palabras(s.get("titulo", "")), palabras(s.get("capitulo", "")), palabras(s.get("texto", "")))
            for s in indice]
    n = len(docs)
    df = {r: sum(1 for _, t, c, x in docs if r in t or r in x or r in c) for r in raices}
    resultados = []
    for s, t, c, x in docs:
        puntos, encontradas = 0.0, 0
        for r in raices:
            if not df[r]:
                continue
            idf = math.log(1 + n / df[r])
            veces = 4 * t.count(r) + 1.5 * c.count(r) + min(x.count(r), 6)
            if veces:
                encontradas += 1
                puntos += idf * (1 + math.log(veces))
        if encontradas:
            puntos *= (encontradas / len(raices)) ** 2      # mejor si están todas las palabras
            resultados.append({"capitulo": s.get("capitulo", ""), "titulo": s.get("titulo", ""),
                               "pagina": s.get("pagina"), "fragmento": _fragmento(s.get("texto", ""), raices),
                               "puntos": round(puntos, 3)})
    resultados.sort(key=lambda r: -r["puntos"])
    return resultados[:maximo]


def ruta_pdf():
    return resource_path(RUTA_PDF)


def abrir_manual():
    """Copia el manual a la carpeta temporal y lo abre con el visor de PDF del equipo. Devuelve la
    ruta abierta. (Dentro del .exe el PDF está en una carpeta temporal de PyInstaller que se borra al
    cerrar la app; la copia se queda para poder tenerlo abierto o guardarlo.)"""
    origen = ruta_pdf()
    if not os.path.isfile(origen):
        raise FileNotFoundError("No se encuentra el manual dentro de la aplicación.")
    destino_dir = os.path.join(tempfile.gettempdir(), "ConsolaRegistroInfo")
    os.makedirs(destino_dir, exist_ok=True)
    destino = os.path.join(destino_dir, NOMBRE_PDF)
    try:
        shutil.copyfile(origen, destino)
    except OSError:
        if not os.path.isfile(destino):      # p. ej. la copia anterior sigue abierta en el visor
            raise
    if sys.platform.startswith("win"):
        os.startfile(destino)  # noqa: S606
    elif sys.platform == "darwin":
        subprocess.Popen(["open", destino])
    else:
        subprocess.Popen(["xdg-open", destino])
    return destino


def guardar_copia(destino):
    """Guarda una copia del manual donde elija el usuario."""
    shutil.copyfile(ruta_pdf(), destino)
    return destino

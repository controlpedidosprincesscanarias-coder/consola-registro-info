"""Versión automática de cada compilación de la Consola Registro Info SAP (TARIFA Finder).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).

Lo usa .github/workflows/compilar.yml (y se puede lanzar a mano):

    python compilacion/version.py [--releases etiquetas.txt] [--forzar 10.6.00] [--escribir] [--github-output FICHERO]

- La versión nueva es la última Release ya creada (también los borradores) +2 en el último número, siempre
  par y con dos cifras como mínimo (10.4.00 → 10.4.02 … 10.4.98 → 10.4.100).
- Salto de versión (petición de Víctor, 23 nov 2026: «la siguiente versión sumamos dos al del centro y 0 al
  final, v 10.4.00»): si APP_VERSION de tarifa_app.py es MAYOR que todas las Releases, se usa tal cual.
  Para el próximo salto basta con poner la nueva en APP_VERSION (p. ej. "v 10.6.00").
- Con --forzar se elige otra (por ejemplo, 10.6.00); tiene que ser mayor que todas y par.
- Con --escribir, la pone en APP_VERSION de tarifa_app.py (solo en la copia que se compila: el
  workflow no sube ese cambio, la versión de cada compilación queda en la etiqueta de su Release).
- Con --github-output, deja version («10.2.38»), etiqueta («v10.2.38») y app_version («v 10.2.38»).
"""
import argparse
import os
import re

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_RE_APP = re.compile(r'^(APP_VERSION\s*=\s*")v\s*([\d.]+)(".*)$', re.M)


def numeros(texto):
    return tuple(int(n) for n in re.findall(r"\d+", str(texto)))


def leer_app_version(src):
    m = _RE_APP.search(src)
    if not m:
        raise SystemExit("ERROR: no se encuentra APP_VERSION en tarifa_app.py")
    return m.group(2)


def versiones_releases(texto):
    """Etiquetas tipo v10.2.38 (una por línea; se ignora lo que no sea una versión de 3 números)."""
    return [m.group(1) for m in re.finditer(r"v?\s*(\d+\.\d+\.\d+)\b", texto or "")]


def formato(n):
    """(10, 4, 0) → «10.4.00»: el último número con dos cifras como mínimo, como el Organizador."""
    x, y, z = n
    return f"{x}.{y}.{z:02d}"


def siguiente(app, releases, forzar=None):
    rel = [numeros(v) for v in releases]
    tope = max([numeros(app)] + rel)
    if forzar:
        n = numeros(forzar)
        if len(n) != 3:
            raise SystemExit(f"ERROR: «{forzar}» no es una versión X.Y.Z")
        if n <= tope:
            raise SystemExit(f"ERROR: {forzar} no es mayor que la última ({formato(tope)})")
        if n[2] % 2:
            raise SystemExit(f"ERROR: {forzar} acaba en impar; las versiones acaban en número par")
        return formato(n)
    n_app = numeros(app)
    if len(n_app) == 3 and n_app[2] % 2 == 0 and (not rel or n_app > max(rel)):
        return formato(n_app)              # salto puesto a mano en APP_VERSION
    x, y, z = tope
    z = z + 2 if z % 2 == 0 else z + 1      # siempre al siguiente par
    return formato((x, y, z))


def escribir_app_version(src, version):
    return _RE_APP.sub(lambda m: f"{m.group(1)}v {version}{m.group(3)}", src, count=1)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--releases")
    ap.add_argument("--forzar")
    ap.add_argument("--escribir", action="store_true")
    ap.add_argument("--github-output")
    a = ap.parse_args()
    ruta = os.path.join(RAIZ, "tarifa_app.py")
    with open(ruta, encoding="utf-8") as f:
        src = f.read()
    rel = []
    if a.releases:
        with open(a.releases, encoding="utf-8-sig") as f:
            rel = versiones_releases(f.read())
    actual = leer_app_version(src)
    nueva = siguiente(actual, rel, (a.forzar or "").strip() or None)
    print(f"APP_VERSION en el código: {actual} · Releases: {len(rel)} · versión nueva: {nueva}")
    if a.escribir:
        with open(ruta, "w", encoding="utf-8", newline="") as f:
            f.write(escribir_app_version(src, nueva))
        print(f"APP_VERSION = \"v {nueva}\" escrito en tarifa_app.py")
    if a.github_output:
        with open(a.github_output, "a", encoding="utf-8") as f:
            f.write(f"version={nueva}\netiqueta=v{nueva}\napp_version=v {nueva}\n")


if __name__ == "__main__":
    main()

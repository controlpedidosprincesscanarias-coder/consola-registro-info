"""Versión automática de cada compilación de la Consola Registro Info SAP (TARIFA Finder).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).

Lo usa .github/workflows/compilar.yml (y se puede lanzar a mano):

    python compilacion/version.py [--releases etiquetas.txt] [--forzar 10.3.0] [--escribir] [--github-output FICHERO]

- La versión nueva es la mayor entre APP_VERSION (tarifa_app.py) y las Releases ya creadas, +2 en el
  último número (siempre par, como en el resto de aplicaciones: 10.2.36 → 10.2.38).
- Con --forzar se elige otra (por ejemplo, para pasar a 10.3.0); tiene que ser mayor que todas.
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


def siguiente(app, releases, forzar=None):
    existentes = [numeros(v) for v in [app] + list(releases)]
    tope = max(existentes)
    if forzar:
        n = numeros(forzar)
        if len(n) != 3:
            raise SystemExit(f"ERROR: «{forzar}» no es una versión X.Y.Z")
        if n <= tope:
            raise SystemExit(f"ERROR: {forzar} no es mayor que la última ({'.'.join(map(str, tope))})")
        if n[2] % 2:
            raise SystemExit(f"ERROR: {forzar} acaba en impar; las versiones acaban en número par")
        return ".".join(map(str, n))
    x, y, z = tope
    z = z + 2 if z % 2 == 0 else z + 1      # siempre al siguiente par
    return f"{x}.{y}.{z}"


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

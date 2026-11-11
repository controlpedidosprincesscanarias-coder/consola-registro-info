# -*- coding: utf-8 -*-
"""
actualizaciones.py
Actualización automática de TARIFA_Finder.exe desde las Releases de GitHub (nov 2026), igual en
espíritu que el Organizador Princess. Sin tkinter: se puede probar con tests/test_actualizaciones.py.

Cómo funciona (cada compañero tiene su propia copia del .exe en su PC):
  1. Al arrancar se consulta, sin iniciar sesión, la última Release PUBLICADA del repositorio
     (los borradores no salen en /releases/latest, así que no llegan a nadie hasta «Publish release»).
  2. Si su versión es mayor que APP_VERSION, la interfaz avisa (app/ui/ventana_actualizacion.py).
  3. «Actualizar ahora»: se descarga TARIFA_Finder.exe a la MISMA carpeta que el .exe actual
     («<nombre>.nuevo.exe»), se comprueba su SHA-256 con el .sha256 publicado (si no coincide, se
     borra y no se toca nada) y se sustituye: el .exe en uso se renombra a «<nombre>.anterior.exe»
     (Windows deja renombrar un .exe abierto), el nuevo ocupa su sitio, se abre y se cierra el viejo.
  4. En el arranque siguiente se borran los restos (.anterior.exe / .nuevo.exe).
tarifa_config.json (carpeta del proyecto y contraseña cifrada) está junto al .exe y no se toca.
Solo se actualiza el .exe compilado; ejecutando «python tarifa_app.py» nunca.

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación son de
Víctor Alejandro Martín Andrade (VAMA).
"""
import hashlib
import json
import os
import re
import ssl
import subprocess
import sys
import urllib.error
import urllib.request

REPO = "controlpedidosprincesscanarias-coder/consola-registro-info"
API_ULTIMA = f"https://api.github.com/repos/{REPO}/releases/latest"
API_ETIQUETA = f"https://api.github.com/repos/{REPO}/releases/tags/{{}}"   # (nov 2026) notas de una versión
URL_VERSIONES = f"https://github.com/{REPO}/releases"
NOMBRE_ASSET = "TARIFA_Finder.exe"
USER_AGENT = "ConsolaRegistroInfo"
TAMANO_MAXIMO = 300 * 1024 * 1024          # un .exe de más de 300 MB no es nuestro


def numeros(version):
    """«v 10.2.38» / «v10.2.38» / «10.2.38» → (10, 2, 38)."""
    return tuple(int(n) for n in re.findall(r"\d+", str(version or "")))


def es_mas_nueva(remota, local):
    r, l = numeros(remota), numeros(local)
    return bool(r) and r > l


def ssl_ctx():
    """Verificación de certificados con el almacén del sistema (en Windows, también el de un proxy
    corporativo con inspección SSL). Nunca sin verificar: de ahí se baja un .exe que se ejecuta."""
    ctx = ssl.create_default_context()
    ctx.load_default_certs()
    return ctx


def datos_release(release):
    """Del JSON de /releases/latest, lo que hace falta: versión, notas y URLs del .exe y su .sha256.
    Devuelve None si la Release no trae el .exe o es un borrador / prerelease."""
    if not isinstance(release, dict) or release.get("draft") or release.get("prerelease"):
        return None
    assets = {a.get("name"): a for a in release.get("assets") or [] if isinstance(a, dict)}
    exe = assets.get(NOMBRE_ASSET)
    if not exe or not exe.get("browser_download_url"):
        return None
    sha = assets.get(NOMBRE_ASSET + ".sha256")
    return {
        "version": ".".join(map(str, numeros(release.get("tag_name")))),
        "etiqueta": release.get("tag_name") or "",
        "titulo": release.get("name") or release.get("tag_name") or "",
        "notas": (release.get("body") or "").strip(),
        "url_exe": exe["browser_download_url"],
        "tamano": int(exe.get("size") or 0),
        "url_sha": sha.get("browser_download_url") if sha else None,
    }


def consultar_ultima(timeout=10):
    """La última Release publicada (datos_release) o None si aún no hay ninguna publicada (GitHub
    responde 404). Lanza la excepción de red en cualquier otro fallo."""
    req = urllib.request.Request(API_ULTIMA, headers={
        "Accept": "application/vnd.github+json", "User-Agent": USER_AGENT})
    try:
        with urllib.request.urlopen(req, timeout=timeout, context=ssl_ctx()) as resp:
            return datos_release(json.loads(resp.read().decode("utf-8")))
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None
        raise


def notas_de_version(version, timeout=10):
    """(nov 2026) Las notas de la Release de `version` («v 10.2.44» → etiqueta «v10.2.44»), para verlas al
    pulsar el número de versión. None si esa versión no tiene Release (404: p. ej. desde el código, o
    una versión que no se publicó). Lanza la excepción de red en cualquier otro fallo."""
    nums = numeros(version)
    if not nums:
        return None
    etiqueta = "v" + ".".join(map(str, nums))
    req = urllib.request.Request(API_ETIQUETA.format(etiqueta), headers={
        "Accept": "application/vnd.github+json", "User-Agent": USER_AGENT})
    try:
        with urllib.request.urlopen(req, timeout=timeout, context=ssl_ctx()) as resp:
            release = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None
        raise
    return (release.get("body") or "").strip() if isinstance(release, dict) else None


def extraer_hash(texto, nombre=NOMBRE_ASSET):
    """SHA-256 (64 hex) para `nombre` en un .sha256: «<hash>», «<hash>  <fichero>» o «<hash> *<fichero>»."""
    lineas = [l.strip() for l in (texto or "").splitlines() if l.strip()]
    for l in lineas:
        m = re.match(r"^([0-9a-fA-F]{64})(?:\s+\*?(.+))?$", l)
        if m and (not m.group(2) or os.path.basename(m.group(2).strip()) == nombre):
            return m.group(1).lower()
    return None


def rutas(ruta_exe):
    """(nuevo, anterior) junto al .exe actual: «X.nuevo.exe» y «X.anterior.exe»."""
    base, ext = os.path.splitext(ruta_exe)
    return base + ".nuevo" + (ext or ".exe"), base + ".anterior" + (ext or ".exe")


def puede_actualizarse():
    """Solo el .exe compilado y si se puede escribir en su carpeta."""
    if not getattr(sys, "frozen", False):
        return False
    return os.access(os.path.dirname(os.path.abspath(sys.executable)), os.W_OK)


def limpiar_restos(ruta_exe):
    """Borra los restos de una actualización anterior («.anterior.exe» y «.nuevo.exe»). Nunca lanza.
    Devuelve (borrados, pendientes): nombres de archivo borrados y los que no se han podido borrar
    todavía (p. ej. la versión vieja aún se está cerrando; la interfaz lo reintenta)."""
    borrados, pendientes = [], []
    for r in rutas(ruta_exe):
        if not os.path.isfile(r):
            continue
        try:
            os.remove(r)
            borrados.append(os.path.basename(r))
        except OSError:
            pendientes.append(os.path.basename(r))
    return borrados, pendientes


def _abrir(url, timeout):
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    return urllib.request.urlopen(req, timeout=timeout, context=ssl_ctx())


def descargar_y_verificar(info, ruta_exe, progreso=None, abrir=_abrir):
    """Descarga el .exe de la Release junto al actual y comprueba su SHA-256. Devuelve la ruta del
    «.nuevo.exe». Si algo falla (sin .sha256, hash distinto, demasiado grande) lo borra y lanza
    RuntimeError con un mensaje para el usuario. `progreso(bytes, total)` se llama por trozos."""
    if not info.get("url_sha"):
        raise RuntimeError("La versión publicada no trae su .sha256: no se puede comprobar y no se instala.")
    with abrir(info["url_sha"], 20) as r:
        esperado = extraer_hash(r.read(65536).decode("utf-8", "ignore"))
    if not esperado:
        raise RuntimeError("No se ha podido leer el SHA-256 publicado.")
    nuevo, _ = rutas(ruta_exe)
    sha = hashlib.sha256()
    total = info.get("tamano") or 0
    hecho = 0
    try:
        with abrir(info["url_exe"], 120) as r, open(nuevo, "wb") as f:
            total = total or int(r.headers.get("Content-Length") or 0)
            while True:
                trozo = r.read(256 * 1024)
                if not trozo:
                    break
                hecho += len(trozo)
                if hecho > TAMANO_MAXIMO:
                    raise RuntimeError("El archivo descargado es demasiado grande.")
                f.write(trozo)
                sha.update(trozo)
                if progreso:
                    progreso(hecho, total)
        if sha.hexdigest() != esperado:
            raise RuntimeError("El archivo descargado no coincide con el SHA-256 publicado. Se ha borrado; "
                               "vuelve a intentarlo más tarde.")
    except BaseException:
        try:
            os.remove(nuevo)
        except OSError:
            pass
        raise
    return nuevo


def sustituir_exe(ruta_nuevo, ruta_exe):
    """Pone el .exe nuevo en el sitio del actual (renombrando el actual a «.anterior.exe»).
    Si falla a mitad, deja el actual como estaba y lanza RuntimeError."""
    _, anterior = rutas(ruta_exe)
    try:
        if os.path.exists(anterior):
            os.remove(anterior)
    except OSError:
        pass
    try:
        os.replace(ruta_exe, anterior)
    except OSError as e:
        raise RuntimeError(f"No se pudo renombrar el programa actual: {e}")
    try:
        os.replace(ruta_nuevo, ruta_exe)
    except OSError as e:
        try:
            os.replace(anterior, ruta_exe)      # vuelta atrás: el actual sigue funcionando
        except OSError:
            pass
        raise RuntimeError(f"No se pudo colocar la versión nueva: {e}")


def entorno_limpio(entorno=None):
    """Variables de entorno para abrir otro .exe de PyInstaller desde este: sin las internas del
    proceso actual (_PYI_*, _MEIPASS2), que harían que el nuevo usara la carpeta temporal del viejo,
    y con PYINSTALLER_RESET_ENVIRONMENT=1, que PyInstaller (6.9+) entiende como «arranque limpio»."""
    env = {k: v for k, v in (os.environ if entorno is None else entorno).items()
           if not k.startswith("_PYI") and k != "_MEIPASS2"}
    env["PYINSTALLER_RESET_ENVIRONMENT"] = "1"
    return env


def lanzar(ruta_exe):
    """Abre el .exe (ya sustituido) como proceso independiente."""
    flags = 0
    if sys.platform.startswith("win"):
        flags = getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    subprocess.Popen([ruta_exe], cwd=os.path.dirname(ruta_exe), close_fds=True, creationflags=flags,
                     env=entorno_limpio())

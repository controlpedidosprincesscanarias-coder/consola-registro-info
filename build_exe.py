# -*- coding: utf-8 -*-
"""
build_exe.py — compila TARIFA_Finder.exe con PyInstaller (--onefile --windowed).
Lo usan el workflow de GitHub (.github/workflows/compilar.yml, lo normal) y COMPILAR_TARIFA.bat
(compilación local en Windows). Incluye:
 - tarifa_app.py y todo el paquete app/ (--collect-submodules app)
 - win32com/pywin32 (Excel), openpyxl, pillow (con PIL._tkinter_finder), pandas y xlrd
 - los assets de ASSETS_NECESARIOS (se para si falta alguno)

Uso:
    python build_exe.py
"""

import subprocess
import sys
import os

# ── Cerrar el ejecutable si está corriendo antes de compilar ─────────────────
def _cerrar_proceso_si_activo(exe_name: str) -> None:
    """
    Termina cualquier proceso con nombre <exe_name>.exe que esté en ejecución.
    Intenta primero con psutil (más limpio); si no está instalado, usa taskkill.
    """
    nombre_proceso = f"{exe_name}.exe"
    terminados = 0

    # Intento 1: psutil (disponible si está instalado)
    try:
        import psutil
        for proc in psutil.process_iter(["pid", "name"]):
            if proc.info["name"] and proc.info["name"].lower() == nombre_proceso.lower():
                try:
                    proc.terminate()
                    proc.wait(timeout=5)
                    terminados += 1
                except (psutil.NoSuchProcess, psutil.TimeoutExpired):
                    try:
                        proc.kill()
                        terminados += 1
                    except psutil.NoSuchProcess:
                        pass
        if terminados:
            print(f"  ⚠ {terminados} instancia(s) de {nombre_proceso} cerrada(s) antes de compilar.")
        return
    except ImportError:
        pass  # psutil no disponible → usar taskkill

    # Intento 2: taskkill nativo de Windows (siempre disponible)
    result = subprocess.run(
        ["taskkill", "/F", "/IM", nombre_proceso],
        capture_output=True, text=True
    )
    if result.returncode == 0:
        print(f"  ⚠ {nombre_proceso} cerrado automáticamente antes de compilar.")
    # returncode 128 = proceso no encontrado → no hace falta avisar
# ─────────────────────────────────────────────────────────────────────────────


# ── Leer versión desde tarifa_app.py sin importar el módulo completo ─────────
def _leer_version(base):
    import re
    ruta = os.path.join(base, "tarifa_app.py")
    try:
        with open(ruta, encoding="utf-8") as f:
            for linea in f:
                if linea.startswith("APP_VERSION"):
                    # Extrae solo el literal entre comillas: APP_VERSION = "v1.0"  # comentario
                    m = re.search(r'["\']([^"\']+)["\']', linea)
                    if m:
                        return m.group(1)
    except Exception:
        pass
    return "vX.X"
# ─────────────────────────────────────────────────────────────────────────────


def main():
    base = os.path.dirname(os.path.abspath(__file__))
    app_version = _leer_version(base)
    exe_name = f"TARIFA_Finder"

    script_main     = os.path.join(base, "tarifa_app.py")
    # (nov 2026, etapa 2) Los módulos viven en app/ y PyInstaller los incluye solo al seguir
    # las importaciones de tarifa_app.py; aquí solo se comprueba que estén.
    modulos = [os.path.join(base, *r.split("/")) for r in (
        "app/comun/tarifa_constants.py", "app/comun/tarifa_core.py", "app/comun/credenciales.py",
        "app/comun/sap_helpers.py", "app/ui/tarifa_dialogs.py", "app/rinfo/crear_registro_info.py",
        "app/tarifas/transformar_tarifa.py", "app/acciones/organizar_dali.py")]
    carpeta_assets = os.path.join(base, "assets")
    logo_ico    = os.path.join(carpeta_assets, "logo.ico")
    # (nov 2026) Todo lo de assets/ se usa y va dentro del .exe: los logos y las imágenes que
    # antes estaban incrustadas en base64 en el código, y el manual de usuario con su índice para
    # «Consultar el manual» (se regeneran con docs/manuales_fuente/generar.sh).
    ASSETS_NECESARIOS = ("logo.png", "logo.ico", "logo_dali_articulos.png", "logo_sap.png",
                         "marcar_borrado_sap.png", "Manual_Usuario_ConsolaRegistroInfo.pdf",
                         "manual_indice.json")

    # ── Verificación de archivos obligatorios ────────────
    for f in [script_main] + modulos:
        if not os.path.isfile(f):
            print(f"ERROR: No se encuentra {f}")
            sys.exit(1)

    # ── Verificación de assets (no bloquean, solo avisan) ─
    faltan = [a for a in ASSETS_NECESARIOS if not os.path.isfile(os.path.join(carpeta_assets, a))]
    if faltan:
        print(f"ERROR: faltan en assets/: {', '.join(faltan)}")
        sys.exit(1)
    tiene_ico = True

    print("=" * 60)
    print(f" COMPILANDO TARIFA FINDER ({app_version} - módulos completos) ")
    print("=" * 60)
    print("Incluyendo:")
    print("  ✔ tarifa_app.py")
    for f in modulos:
        print(f"  ✔ {os.path.relpath(f, base)}")
    print("  ✔ win32com (Excel)")
    print("  ✔ openpyxl")
    print("  ✔ pillow")
    print("  ✔ pywin32")
    print("  ✔ pandas  (Acción 1 — organizar ARTICULOS DALI)")
    print("  ✔ xlrd    (lectura de .xls DALI)")
    for a in ASSETS_NECESARIOS:
        print(f"  ✔ assets/{a}")
    print("=" * 60)

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--windowed",
        "--clean",
        "--name", exe_name,

        # ── Mitigación para win32com ─────────────────────
        "--collect-submodules", "win32com",
        "--hidden-import", "win32com",
        "--hidden-import", "win32com.client",
        "--hidden-import", "pywintypes",
        "--hidden-import", "pythoncom",

        # ── Dependencias openpyxl ────────────────────────
        "--hidden-import", "openpyxl",
        "--hidden-import", "openpyxl.styles",
        "--hidden-import", "openpyxl.styles.differential",
        "--hidden-import", "openpyxl.formatting.rule",
        "--hidden-import", "openpyxl.utils",

        # ── Dependencias Pillow ──────────────────────────
        "--hidden-import", "PIL",
        "--hidden-import", "PIL.Image",
        "--hidden-import", "PIL.ImageTk",
        # (nov 2026) Sin él, ImageTk falla dentro del .exe («No module named
        # 'PIL._tkinter_finder'») y la cabecera muestra una «P» en vez del logo.
        "--hidden-import", "PIL._tkinter_finder",

        # ── pandas + xlrd (Acción 1 — ARTICULOS DALI) ───
        "--hidden-import", "pandas",
        "--hidden-import", "pandas.core.frame",
        "--hidden-import", "pandas.core.series",
        "--hidden-import", "pandas.io.formats.excel",
        "--hidden-import", "xlrd",
        "--hidden-import", "xlrd.biffh",

        # ── Paquete app/ (todos sus submódulos, aunque alguno se importe dentro de una función)
        "--collect-submodules", "app",
    ]

    # ── Assets (todos obligatorios, comprobados arriba) ──
    for a in ASSETS_NECESARIOS:
        cmd += ["--add-data", f"{os.path.join(carpeta_assets, a)}{os.pathsep}assets"]
    cmd += ["--icon", logo_ico]  # icono del .exe y barra de título

    # ── Archivo principal ────────────────────────────────
    cmd.append(script_main)

    # ── Cerrar el ejecutable si está corriendo ───────────
    _cerrar_proceso_si_activo(exe_name)

    # ── Ejecución de PyInstaller ─────────────────────────
    result = subprocess.run(cmd, cwd=base)

    if result.returncode == 0:
        exe = os.path.join(base, "dist", f"{exe_name}.exe")
        print("\n✔ COMPILACIÓN COMPLETADA CON ÉXITO")
        print(f"Ejecutable generado: {exe}")
        if tiene_ico:
            print("Icono personalizado aplicado al .exe")
        print("Listo para distribuir en cualquier PC con Excel instalado.")
    else:
        print("\n✘ ERROR DURANTE LA COMPILACIÓN")
        sys.exit(1)


if __name__ == "__main__":
    main()
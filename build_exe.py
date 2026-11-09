# -*- coding: utf-8 -*-
"""
build_exe.py — Versión V7.1 (accion6_registro_sap eliminado, Acción 6 integrada en tarifa_app)
Compila TARIFA Finder con soporte completo para:
 - win32com (Excel)
 - transformar_tarifa.py (V9.7)
 - interfaz tarifa_app.py
 - dependencias openpyxl, pillow, pywin32
 - dependencias pandas + xlrd  (Acción 1 — organizar ARTICULOS DALI)
 - assets/logo.png y assets/logo.ico (logo personalizado)
 - crear_registro_info.py (botón CREAR REGISTRO INFO — módulo independiente)

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
    script_mod      = os.path.join(base, "transformar_tarifa.py")
    script_constants = os.path.join(base, "tarifa_constants.py")
    script_core     = os.path.join(base, "tarifa_core.py")
    script_dialogs  = os.path.join(base, "tarifa_dialogs.py")
    script_cri      = os.path.join(base, "crear_registro_info.py")
    logo_png    = os.path.join(base, "assets", "logo.png")
    logo_ico    = os.path.join(base, "assets", "logo.ico")
    logo_dali   = os.path.join(base, "assets", "logo_dali_articulos.png")

    # ── Verificación de archivos obligatorios ────────────
    for f in (script_main, script_mod, script_constants, script_core, script_dialogs, script_cri):
        if not os.path.isfile(f):
            print(f"ERROR: No se encuentra {f}")
            sys.exit(1)

    # ── Verificación de assets (no bloquean, solo avisan) ─
    tiene_png = os.path.isfile(logo_png)
    tiene_ico = os.path.isfile(logo_ico)
    tiene_logo_dali = os.path.isfile(logo_dali)

    print("=" * 60)
    print(f" COMPILANDO TARIFA FINDER ({app_version} - módulos completos) ")
    print("=" * 60)
    print("Incluyendo:")
    print("  ✔ tarifa_app.py")
    print("  ✔ tarifa_constants.py")
    print("  ✔ tarifa_core.py")
    print("  ✔ tarifa_dialogs.py")
    print("  ✔ crear_registro_info.py (botón CREAR REGISTRO INFO)")
    print("  ✔ transformar_tarifa.py (V9.7)")
    print("  ✔ win32com (Excel)")
    print("  ✔ openpyxl")
    print("  ✔ pillow")
    print("  ✔ pywin32")
    print("  ✔ pandas  (Acción 1 — organizar ARTICULOS DALI)")
    print("  ✔ xlrd    (lectura de .xls DALI)")
    print(f"  {'✔' if tiene_png else '✘'} assets/logo.png"
          + ("" if tiene_png else "  ← NO ENCONTRADO (se usará logo embebido)"))
    print(f"  {'✔' if tiene_ico else '✘'} assets/logo.ico"
          + ("" if tiene_ico else "  ← NO ENCONTRADO (icono genérico)"))
    print(f"  {'✔' if tiene_logo_dali else '✘'} assets/logo_dali_articulos.png"
          + ("" if tiene_logo_dali else "  ← NO ENCONTRADO (logo omitido en Acción 1)"))
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

        # ── pandas + xlrd (Acción 1 — ARTICULOS DALI) ───
        "--hidden-import", "pandas",
        "--hidden-import", "pandas.core.frame",
        "--hidden-import", "pandas.core.series",
        "--hidden-import", "pandas.io.formats.excel",
        "--hidden-import", "xlrd",
        "--hidden-import", "xlrd.biffh",

        # ── Módulo de transformación ─────────────────────
        "--add-data", f"{script_mod}{os.pathsep}.",
        "--add-data", f"{script_constants}{os.pathsep}.",
        "--add-data", f"{script_core}{os.pathsep}.",
        "--add-data", f"{script_dialogs}{os.pathsep}.",
        "--add-data", f"{script_cri}{os.pathsep}.",
    ]

    # ── Assets: solo se añaden si existen ───────────────
    if tiene_png:
        cmd += ["--add-data", f"{logo_png}{os.pathsep}assets"]
    if tiene_ico:
        cmd += ["--add-data", f"{logo_ico}{os.pathsep}assets"]
        cmd += ["--icon", logo_ico]  # icono del .exe y barra de título
    if tiene_logo_dali:
        cmd += ["--add-data", f"{logo_dali}{os.pathsep}assets"]

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
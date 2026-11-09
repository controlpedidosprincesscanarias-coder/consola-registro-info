# -*- coding: utf-8 -*-
import re, sys, os

accion = sys.argv[1] if len(sys.argv) > 1 else ""
ruta   = "tarifa_app.py"

if accion == "leer":
    try:
        with open(ruta, encoding="utf-8") as f:
            m = re.search(r'APP_VERSION\s*=\s*"([^"]+)"', f.read())
        print(m.group(1) if m else "NO_ENCONTRADA")
    except Exception as e:
        print("ERROR: " + str(e))

elif accion == "escribir":
    nueva = sys.argv[2] if len(sys.argv) > 2 else ""
    if not nueva:
        print("ERROR: version vacia")
        sys.exit(1)
    # Asegurar que siempre empieza por "v " aunque el usuario no lo escriba
    nueva = nueva.strip()
    if not nueva.lower().startswith("v"):
        nueva = "v " + nueva
    elif nueva.lower().startswith("v") and not nueva.lower().startswith("v "):
        nueva = "v " + nueva[1:]
    try:
        with open(ruta, encoding="utf-8") as f:
            c = f.read()
        n = re.sub(r'(APP_VERSION\s*=\s*)"[^"]*"', r'\g<1>"' + nueva + '"', c, count=1)
        if n == c:
            print("ERROR: APP_VERSION no encontrada")
            sys.exit(1)
        with open(ruta, "w", encoding="utf-8") as f:
            f.write(n)
        print("OK:" + nueva)
    except Exception as e:
        print("ERROR: " + str(e))
        sys.exit(1)

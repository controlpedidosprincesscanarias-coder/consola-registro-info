# -*- coding: utf-8 -*-
"""
capturas.py — capturas del manual de usuario de la Consola Registro Info SAP.

Arranca la app de verdad (tarifa_app.App) en una pantalla virtual (Xvfb, ver capturas.sh) con un
proyecto de EJEMPLO inventado en una carpeta temporal, sin red (ni Catálogo Asignaciones ni GitHub)
y abre una a una sus ventanas para fotografiarlas. Cada captura es la ventana recortada, al doble de
tamaño, con una barra de título como la de Windows y los números de referencia en naranja, como en
los manuales de Control de Pedidos y del Organizador Princess.

La idea y el desarrollo de la aplicación, incluidas sus pruebas, son de Víctor Alejandro Martín
Andrade, para el Departamento de Central de Compras de Princess Hotels en Canarias.
Uso: capturas.sh [carpeta_salida]   (SOLO=prefijo repite solo esas capturas; VOLCAR=1 guarda los widgets)
"""
import io
import json
import os
import subprocess
import sys
import tempfile
import tkinter as tk
import traceback
import unicodedata

from PIL import Image, ImageDraw, ImageFont

RAIZ = os.path.abspath(sys.argv[1])
SALIDA = os.path.abspath(sys.argv[2])
SOLO = os.environ.get("SOLO", "")
VOLCAR = os.environ.get("VOLCAR") == "1"
sys.path.insert(0, RAIZ)
os.chdir(RAIZ)

NARANJA = (232, 89, 12)
E = 2
FUENTE = "/usr/share/fonts/opentype/inter/Inter-Bold.otf"
FUENTE_TIT = "/usr/share/fonts/opentype/inter/Inter-Regular.otf"
LOG = []


def log(m):
    LOG.append(m)
    print(m, flush=True)


# ── Proyecto de ejemplo (todo inventado) ─────────────────────────────────────
DEMO = tempfile.mkdtemp(prefix="consola_manual_")
PROYECTO = os.path.join(DEMO, "PROYECTO")
TARIFAS = os.path.join(PROYECTO, "TARIFAS")
RECURSOS = os.path.join(PROYECTO, "RECURSOS")
for d in (TARIFAS, RECURSOS):
    os.makedirs(d, exist_ok=True)
for n in ("SUMINISTROS DEMO SL [INCOMPLETO] [11-11-2026].xlsx", "PAPELERIA EJEMPLO SL [INCOMPLETO] [10-11-2026].xlsx",
          "SUMINISTROS DEMO SL [11-11-2026].xlsx", "DISTRIBUCIONES MODELO SA [09-11-2026].xlsx",
          "RINFO DISTRIBUCIONES MODELO SA [09-11-2026].xlsx", "Purchasing Info Records.xlsx", "PROVEEDOR_DEMO.xls",
          "OTRO_PROVEEDOR.xls"):
    open(os.path.join(TARIFAS, n), "wb").close()
for n in ("ARTICULOS_DALI.xlsx", "Productos.xlsx", "Listas de productos.xlsx", "Proveedores.xlsx"):
    open(os.path.join(RECURSOS, n), "wb").close()
# Catálogo completo de ejemplo (desde la v10.2.50 es el único archivo de artículos; ya no hay LISTADO)
CATALOGO = os.path.join(RECURSOS, "catalogo-dali-completo.xlsx")


def _crear_catalogo_demo():
    import openpyxl
    from app.comun.sap_helpers import CATALOGO_COLUMNAS, CATALOGO_FILA_CABECERA
    wb = openpyxl.Workbook()
    ws = wb.active
    for i, (_k, titulo, _a) in enumerate(CATALOGO_COLUMNAS, start=1):
        ws.cell(row=CATALOGO_FILA_CABECERA, column=i, value=titulo)
    col = {k: i for i, (k, _t, _a) in enumerate(CATALOGO_COLUMNAS, start=1)}
    for n, (dali, desc, prov, fam, sap) in enumerate(ARTICULOS_DEMO, start=CATALOGO_FILA_CABECERA + 1):
        for k, v in (("clave", dali), ("descripcion", desc), ("unidad", "UN"), ("proveedor", prov), ("familia", fam),
                     ("cod_sap", sap), ("activo", "SI")):
            if v is not None:
                ws.cell(row=n, column=col[k], value=v)
    wb.save(CATALOGO)


ARTICULOS_DEMO = [
    (10001, "BOLIGRAFO AZUL CAJA 50 UD", "PAPELERIA EJEMPLO SL", "PAPELERIA", None),
    (10002, "CARPETA ARCHIVADOR A4", "PAPELERIA EJEMPLO SL", "PAPELERIA", None),
    (10003, "PAPEL A4 80 GR PAQUETE 500", "PAPELERIA EJEMPLO SL", "PAPELERIA", None),
    (10011, "BAYETA MICROFIBRA AZUL", "SUMINISTROS DEMO SL", "LIMPIEZA", None),
    (10012, "GUANTE NITRILO TALLA M CAJA 100", "SUMINISTROS DEMO SL", "LIMPIEZA", None),
    (10013, "BOLSA BASURA 85X105 ROLLO 10", "SUMINISTROS DEMO SL", "LIMPIEZA", None),
    (10021, "VASO PLASTICO 200 ML PAQUETE 100", "DISTRIBUCIONES MODELO SA", "MENAJE DESECHABLE", None),
    (10022, "SERVILLETA PAPEL 40X40 PAQUETE 50", "DISTRIBUCIONES MODELO SA", "MENAJE DESECHABLE", None),
    (10031, "ROTULADOR PERMANENTE NEGRO", "PAPELERIA EJEMPLO SL", "PAPELERIA", "10509001"),
]
CONFIG = os.path.join(DEMO, "tarifa_config.json")
with open(CONFIG, "w", encoding="utf-8") as f:
    json.dump({"carpeta_proyecto": PROYECTO}, f)

# ── Sin red y con la configuración de ejemplo ────────────────────────────────
import webbrowser  # noqa: E402
webbrowser.open = lambda *a, **k: True
from app.comun import tarifa_core  # noqa: E402
tarifa_core.get_config_path = lambda: CONFIG
from app.comun import actualizaciones  # noqa: E402
actualizaciones.consultar_ultima = lambda timeout=10: None
actualizaciones.puede_actualizarse = lambda: True       # la ventana de aviso, como la verá un compañero
RUTA_WIN = "G:\\CENTRAL COMPRAS\\PROYECTO"             # ruta de ejemplo que se enseña en las capturas
from app.comun import manual as _manual  # noqa: E402
_manual.abrir_manual = lambda: None

import tarifa_app  # noqa: E402
# El inicio de sesión de Catálogo Asignaciones (sin red) se fotografía aparte (c_login); aquí se entra sin él.
tarifa_app.App._iniciar_sesion_consola = lambda self: (self.deiconify(), self._tras_iniciar_sesion())
# La versión que se enseña es la del .exe que llevará este manual (en el código fuente, APP_VERSION es la antigua).
tarifa_app.APP_VERSION = os.environ.get("VERSION", "v 10.2.98")
_crear_catalogo_demo()


# ── Utilidades de captura (como las del Organizador) ─────────────────────────
def pantalla():
    datos = subprocess.run(["import", "-window", "root", "png:-"], capture_output=True, check=True).stdout
    return Image.open(io.BytesIO(datos)).convert("RGB")


def widgets(raiz):
    out = []

    def rec(w):
        for h in w.winfo_children():
            if isinstance(h, tk.Toplevel):
                continue
            out.append(h)
            rec(h)
    rec(raiz)
    return out


def texto(w):
    for o in ("text",):
        try:
            return str(w.cget(o))
        except tk.TclError:
            pass
    try:
        return w.get() if w.winfo_class() == "Entry" else ""
    except Exception:  # noqa: BLE001
        return ""


def buscar(raiz, t, clase=None, n=0):
    hits = [w for w in widgets(raiz) if (clase is None or w.winfo_class() == clase)
            and t.lower() in texto(w).lower() and w.winfo_ismapped()]
    if len(hits) <= n:
        log(f"   ⚠ sin widget «{t}» ({clase})")
        return None
    return hits[n]


def de_clase(raiz, clase, n=0):
    hits = [w for w in widgets(raiz) if w.winfo_class() == clase and w.winfo_ismapped()]
    return hits[n] if len(hits) > n else None


def caja_de(v, ws):
    ws = [w for w in ws if w is not None and w.winfo_ismapped()]
    if not ws:
        return None
    x0 = min(w.winfo_rootx() for w in ws) - v.winfo_rootx(); y0 = min(w.winfo_rooty() for w in ws) - v.winfo_rooty()
    x1 = max(w.winfo_rootx() + w.winfo_width() for w in ws) - v.winfo_rootx()
    y1 = max(w.winfo_rooty() + w.winfo_height() for w in ws) - v.winfo_rooty()
    return (x0, y0, x1 - x0, y1 - y0)


def volcar(v, nombre):
    if not VOLCAR:
        return
    with open(os.path.join(SALIDA, f"arbol_{nombre}.txt"), "w", encoding="utf-8") as f:
        for w in widgets(v):
            try:
                f.write(f"{w.winfo_class():12} {texto(w)[:50]!r:55} @{w.winfo_rootx() - v.winfo_rootx()},"
                        f"{w.winfo_rooty() - v.winfo_rooty()} {w.winfo_width()}x{w.winfo_height()}\n")
            except tk.TclError:
                pass


def _barra_titulo(img, titulo):
    W, h = img.width, 32 * E
    barra = Image.new("RGB", (W, h), (32, 32, 32))
    d = ImageDraw.Draw(barra)
    try:
        ico = Image.open(os.path.join(RAIZ, "assets", "logo.png")).convert("RGBA")
        ico.thumbnail((18 * E, 18 * E))
        barra.paste(ico, (10 * E, (h - ico.height) // 2), ico)
    except Exception:  # noqa: BLE001
        pass
    titulo = "".join(ch for ch in titulo if unicodedata.category(ch) not in ("So", "Mn", "Cn", "Co", "Cs")).strip()
    d.text((36 * E, h // 2), titulo, font=ImageFont.truetype(FUENTE_TIT, 12 * E), fill=(235, 235, 235), anchor="lm")
    y = h // 2
    for xc, forma in ((W - 23 * E, "x"), (W - 69 * E, "o"), (W - 115 * E, "-")):
        if forma == "x":
            d.line((xc - 5 * E, y - 5 * E, xc + 5 * E, y + 5 * E), fill=(235, 235, 235), width=E)
            d.line((xc - 5 * E, y + 5 * E, xc + 5 * E, y - 5 * E), fill=(235, 235, 235), width=E)
        elif forma == "o":
            d.rectangle((xc - 5 * E, y - 5 * E, xc + 5 * E, y + 5 * E), outline=(235, 235, 235), width=E)
        else:
            d.line((xc - 5 * E, y, xc + 5 * E, y), fill=(235, 235, 235), width=E)
    out = Image.new("RGB", (W, img.height + h))
    out.paste(barra, (0, 0)); out.paste(img, (0, h))
    return out, h


def _marca(d, caja, n, lado="tl", W=0, H=0):
    x0, y0, x1, y1 = caja
    d.rounded_rectangle((x0 - 4 * E, y0 - 4 * E, x1 + 4 * E, y1 + 4 * E), radius=8 * E, outline=NARANJA, width=int(2.5 * E))
    s = 26 * E
    bw, bh = x1 - x0, y1 - y0
    pos = {"l": (x0 - s - 8 * E, y0 + bh / 2 - s / 2), "r": (x1 + 8 * E, y0 + bh / 2 - s / 2),
           "t": (x0 + bw / 2 - s / 2, y0 - s - 8 * E), "b": (x0 + bw / 2 - s / 2, y1 + 8 * E),
           "tr": (x1 - s / 2, y0 - s / 2 - 4 * E), "ir": (x1 - s - 6 * E, y0 + bh / 2 - s / 2),
           "in": (x0 + 4 * E, y0 + bh / 2 - s / 2)}.get(lado, (x0 - s / 2 - 2 * E, y0 - s / 2 - 2 * E))
    x, y = max(2, min(W - s - 2, pos[0])), max(2, min(H - s - 2, pos[1]))
    d.ellipse((x - 2 * E, y - 2 * E, x + s + 2 * E, y + s + 2 * E), fill=(255, 255, 255))
    d.ellipse((x, y, x + s, y + s), fill=NARANJA)
    d.text((x + s / 2, y + s / 2), str(n), font=ImageFont.truetype(FUENTE, 14 * E), fill=(255, 255, 255), anchor="mm")


def foto(v, nombre, marcas=(), titulo=None):
    """Guarda la ventana `v` como <nombre>.png. marcas: [(widget | (x, y, w, h), n, lado)]."""
    if SOLO and not nombre.startswith(SOLO):
        return
    v.lift(); v.update_idletasks(); v.update()
    # El ratón, a una esquina: si queda encima de un botón, Tk lo pinta «activo» (texto negro sobre fondo oscuro)
    try:
        v.event_generate("<Motion>", warp=True, x=0, y=0)
        v.update()
    except Exception:  # noqa: BLE001
        pass
    volcar(v, nombre)
    vx, vy, W, H = v.winfo_rootx(), v.winfo_rooty(), v.winfo_width(), v.winfo_height()
    img = pantalla().crop((vx, vy, vx + W, vy + H)).resize((W * E, H * E), Image.LANCZOS)
    img, dy = _barra_titulo(img, titulo if titulo is not None else v.title())
    d = ImageDraw.Draw(img)
    for m in marcas:
        w, n = m[0], m[1]
        if w is None:
            log(f"   ⚠ {nombre}: falta el widget de la marca {n}")
            continue
        x, y, ww, hh = w if isinstance(w, tuple) else (w.winfo_rootx() - vx, w.winfo_rooty() - vy, w.winfo_width(), w.winfo_height())
        _marca(d, (x * E, y * E + dy, (x + ww) * E, (y + hh) * E + dy), n, m[2] if len(m) > 2 else "tl", img.width, img.height)
    img.save(os.path.join(SALIDA, nombre + ".png"))
    log(f"  ✓ {nombre}")


def ultima(app, antes):
    nuevas = [w for w in app.winfo_children() if isinstance(w, tk.Toplevel) and str(w) not in antes and w.winfo_exists()]
    return nuevas[-1] if nuevas else None


# ── Datos de ejemplo para las ventanas ───────────────────────────────────────
LOG_DEMO = [
    ("── INICIO DE TRANSFORMACIÓN ──────────────────────────────", None),
    ("Archivo fuente: TARIFAS\\TARIFA.xls", "INFO"),
    ("Formato detectado: A  (cab=col4, desc=col4, fmt=col6, precio=col7, dto=col8)", "OK"),
    ("  └ cod=900  nombre=SUMINISTROS DEMO SL  periodo=1er SEMESTRE 2027", "OK"),
    ("Paso 27b: Cruzando CODIGO DALI (col B) con el catálogo DALI-SAP...", "INFO"),
    ("  └ 124 coincidencias / 0 sin match", "OK"),
    ("  ⚠  3 código(s) DALI sin DESCRIPCION/UNIDAD SAP en el catálogo: 10001, 10002, 10003", "WARN"),
    ("Guardando: SUMINISTROS DEMO SL [INCOMPLETO] [11-11-2026].xlsx", "INFO"),
    ("  ⚠ Archivo marcado como [INCOMPLETO] — faltan 3 códigos SAP", "WARN"),
    ("Catálogo Asignaciones: comprobando qué ha cambiado...", "INFO"),
    ("  └ Catálogo Asignaciones actualizado: la web y el catálogo local coinciden.", "OK"),
    ("── FIN DE TRANSFORMACIÓN ─────────────────────────────────", None),
    ("Iniciando: Crear Registro Info…", "INFO"),
    ("Proveedor SAP: «SUMINISTROS DEMO SL» [cód. 1009999]", "OK"),
    ("Purchasing encontrado: Purchasing Info Records.xlsx (score 100%)", "OK"),
    ("  ✅ Sin incidencias detectadas. El archivo puede importarse en SAP.", "OK"),
    ("Archivo generado: RINFO SUMINISTROS DEMO SL [11-11-2026].xlsx (3 material(es) nuevo(s) añadido(s))", "OK"),
]


def pasos(app):
    """Lista de capturas: (nombre, abrir(app) -> ventana | None, marcas(v) -> lista)."""
    from app.ui.tarifa_dialogs import VentanaCrearProyecto, VentanaSeleccionXLS, VentanaSAPFaltantes, VentanaResolverSAPDuplicados
    from app.rinfo import crear_registro_info as cri
    from app.ui.ventana_actualizacion import VentanaActualizacion
    from app.ui.ventana_consulta_manual import VentanaConsultaManual
    from app.comun.tarifa_core import listar_xls
    from app.ui import ventana_login, ventana_publicar_catalogo
    from app.ui.ventana_asignar_sap import VentanaAsignarSap
    from app.ui.ventana_listados_sap import VentanaListadosSap, VentanaAntesAccion3
    from app.comun import listados_sap as _L                 # (19 nov 2026) ventana antes de la Acción 3
    from app.comun import asignar_sap
    from app.comun import estado_web                          # (18 nov 2026) aviso sin conexión
    from app.ui.ventana_sin_conexion import VentanaSinConexion
    tarifa = os.path.join(TARIFAS, "SUMINISTROS DEMO SL [11-11-2026].xlsx")
    win_tarifas = RUTA_WIN + "\\TARIFAS"
    B = buscar

    def caja(v, *ws):
        return caja_de(v, [w for w in ws if w is not None])

    def boton(v, t, n=0):
        return B(v, t, "Button", n=n)
    lista = [
        ("c_crear_proyecto", lambda: VentanaCrearProyecto(app), lambda v: [(boton(v, "Crear carpetas"), 1, "t")]),
        ("c_seleccion_xls", lambda: VentanaSeleccionXLS(app, listar_xls(TARIFAS), TARIFAS, app.consola), lambda v: [(de_clase(v, "Treeview"), 1, "ir"), (boton(v, "Usar este archivo"), 2, "t")]),
        ("c_incompletos_a4", lambda: app._elegir_inc_para_procesar([
            {"nombre_archivo": n, "nombre_display": os.path.splitext(n)[0], "ruta": os.path.join(TARIFAS, n), "fecha_mod": "11-11-2026 10:15"}
            for n in sorted(os.listdir(TARIFAS)) if "[INCOMPLETO]" in n]), lambda v: [(de_clase(v, "Treeview"), 1, "ir"), (boton(v, "Seleccionar"), 2, "t")]),
        ("c_incompletos_a6", lambda: app._crear_registro_sap(), lambda v: [(B(v, "Instrucciones de uso").master, 1, "ir"), (de_clase(v, "Treeview"), 2, "ir"), (boton(v, "Abrir en Excel"), 3, "t")]),
        ("c_sap_faltantes", lambda: VentanaSAPFaltantes(app, [(30, 10001), (49, 10002), (66, 10003)],
                                                         CATALOGO,
                                                         app.log_consola, {10001: "BOLIGRAFO AZUL CAJA 50 UD", 10002: "CARPETA ARCHIVADOR A4",
                                                                           10003: "PAPEL A4 80 GR PAQUETE 500"}), lambda v: [(caja(v, B(v, "CÓDIGO DALI"), B(v, "FORMATO")), 1, "b"), (B(v, "CÓDIGO SAP  (introducir)"), 2, "b"), (boton(v, "Guardar en el catálogo"), 3, "t")]),
        ("c_sap_duplicados", lambda: VentanaResolverSAPDuplicados(app, {"10509999": [(10002, True, 2010, "CARPETA ARCHIVADOR A4"),
                                                                                      (10004, False, None, "CARPETA ARCHIVADOR A4 AZUL")]}),
         lambda v: [(caja(v, B(v, "Borrar de DALI 010002"), B(v, "Borrar de DALI 010004")), 1, "t"), (boton(v, "Confirmar y continuar"), 2, "t"), (boton(v, "Cancelar todo"), 3, "t")]),
        ("c_login", lambda: ventana_login.pedir_inicio_sesion(app, "comprador@ejemplo.es", None,
                                                               lambda e, c: ("no_autorizado", "Usuario no autorizado.", None)),
         lambda v: [(de_clase(v, "Entry", 0), 1, "r"), (de_clase(v, "Entry", 1), 2, "r"), (boton(v, "Entrar"), 3, "t"),
                    (boton(v, "Salir"), 4, "t")]),
        ("c_publicar", lambda: ventana_publicar_catalogo.pedir_confirmacion(app, DIF_DEMO, SIMS_DEMO, "completar la tarifa (Acción 4)"),
         lambda v: [(B(v, "Datos DALI:"), 1, "r"), (B(v, "se cambiaron también en la web"), 2, "r"), (B(v, "que la web NO cambiará"), 3, "r"),
                    (de_clase(v, "Text"), 4, "ir"), (boton(v, "Actualizar la web"), 5, "t"), (boton(v, "Ahora no"), 6, "t")]),
        ("c_asignar_sap", lambda: VentanaAsignarSap(app, asignar_sap.pendientes(CATALOGO), PRODUCTOS_DEMO, lambda n, r: None),
         marcas_asignar_sap),
        ("c_listados_sap", lambda: VentanaListadosSap(app, ["10509099"], RECURSOS), marcas_listados_sap),
        ("c_antes_accion3", lambda: _antes_accion3(app, VentanaAntesAccion3, _L), marcas_antes_accion3),
        ("c_volcado_dali", lambda: _abrir_volcado_dali(app), marcas_volcado_dali),
        ("c_sin_conexion", lambda: VentanaSinConexion(app, "traer lo cambiado en Catálogo Asignaciones",
                                                      estado_web.clasificar(TimeoutError("timed out"))),
         lambda v: [(B(v, "tarda en responder"), 1, "r"), (B(v, "sigue trabajando con el catálogo Excel"), 2, "r"),
                    (boton(v, "Reintentar"), 3, "l"), (boton(v, "Seguir sin la web"), 4, "r")]),
        ("c_rinfo_tarifa", lambda: cri.VentanaSeleccionTarifa(app, cri._listar_tarifas_validas(TARIFAS), TARIFAS), lambda v: [(de_clase(v, "Treeview"), 1, "ir"), (boton(v, "Continuar"), 2, "t")]),
        ("c_rinfo_purchasing", lambda: cri.VentanaPurchasingEncontrado(app, tarifa, os.path.basename(tarifa), "SUMINISTROS DEMO SL",
                                                                      "SUMINISTROS DEMO SL", "1009999",
                                                                      {"ruta": os.path.join(TARIFAS, "Purchasing Info Records.xlsx"),
                                                                       "archivo": "Purchasing Info Records.xlsx", "score": 100,
                                                                       "proveedor": "SUMINISTROS DEMO SL"}, TARIFAS, app), lambda v: [(B(v, "SUMINISTROS DEMO SL  [1009999]"), 1, "r"), (B(v, "Purchasing Info Records.xlsx"), 2, "r"), (boton(v, "Continuar usando"), 3, "l"), (boton(v, "Comenzar desde cero"), 4, "l")]),
        ("c_rinfo_purchasing_aviso", lambda: _purchasing_con_aviso(app, cri, tarifa), lambda v: [
            (next((w.master for w in widgets(v) if "puede estar desactualizado" in texto(w)), None), 1, "l"),
            (boton(v, "Cómo descargar el Purchasing"), 2, "r")]),
        ("c_rinfo_sin_purchasing", lambda: cri.VentanaOpcionesSinPurchasing(app, tarifa, os.path.basename(tarifa), "SUMINISTROS DEMO SL",
                                                                          "SUMINISTROS DEMO SL", "1009999", TARIFAS, app), lambda v: [(boton(v, "Descargar el Purchasing"), 1, "l"), (boton(v, "Crear nuevo Registro Info"), 2, "l")]),
        ("c_rinfo_instrucciones", lambda: cri.VentanaInstruccionesSAP(app, "SUMINISTROS DEMO SL"), lambda v: [(B(v, "https://"), 1, "r")]),
        ("c_rinfo_telefono", lambda: cri.VentanaTelefonoRINFO(app, lambda *a: None), lambda v: [(de_clase(v, "Entry"), 1, "r"), (boton(v, "Asignar teléfono"), 2, "l"), (boton(v, "Sin teléfono"), 3, "l")]),
        ("c_rinfo_fechas", lambda: cri.VentanaFechasRINFO(app, lambda *a: None), lambda v: [(de_clase(v, "Entry", 0), 1, "r"), (de_clase(v, "Entry", 1), 2, "r"), (boton(v, "Confirmar fechas"), 3, "t")]),
        ("c_rinfo_centros", lambda: cri.VentanaCentrosRINFO(app, lambda *a: None), lambda v: [(caja(v, B(v, "Todos"), B(v, "Ninguno")), 1, "r"), (caja(v, B(v, "0110"), B(v, "0910")), 2, "r"), (boton(v, "Confirmar selección"), 3, "t")]),
        ("c_plantilla_rinfo", lambda: __import__("app.ui.ventana_plantilla_rinfo", fromlist=["x"]).VentanaPlantillaRinfo(
            app, TARIFAS, lambda r: None), lambda v: (v._lbl_zona_sub.config(text="(uno o varios a la vez)"),   # el texto de Windows (en Linux no hay arrastre)
                         [(v._zona, 1, "l"), (boton(v, "Examinar"), 2, "t")])[1]),
        ("c_rinfo_sin_tarifa", lambda: cri._ventana_sin_tarifa(app, [
            {"codigo": "10509001", "desc": "ROTULADOR PERMANENTE NEGRO", "fecha": "31/12/2026", "fila_h1": 12},
            {"codigo": "10509002", "desc": "SOBRE BLANCO 110X220 CAJA 500", "fecha": "31/12/2026", "fila_h1": 13}]), lambda v: [(de_clase(v, "Treeview"), 1, "ir"), (boton(v, "MANTENER"), 2, "t"), (boton(v, "DAR BAJA"), 3, "t")]),
        ("c_rinfo_reactivar", lambda: cri.VentanaMaterialesParaReactivar(app, {"10509003": "5300009991"},
                                                                         {"10509003": "GRAPAS 26/6 CAJA 1000"},
                                                                         {"10509003": ["0110", "0120"]}), lambda v: [(B(v, "10509003"), 1, "r"), (B(v, "Cómo reactivarlos"), 2, "r")]),
        ("c_rinfo_generado", lambda: cri.VentanaConfirmacionGenerado(app, "RINFO SUMINISTROS DEMO SL [11-11-2026].xlsx",
                                                                     win_tarifas + "\\RINFO SUMINISTROS DEMO SL [11-11-2026].xlsx",
                                                                     "SUMINISTROS DEMO SL", "1009999", 92, 3,
                                                                     "Purchasing Info Records.xlsx"),
         lambda v: [(B(v, "Resumen de materiales").master, 1, "ir"), (boton(v, "Sí, abrir ahora"), 2, "t")]),
        ("c_rinfo_creado", lambda: cri.VentanaConfirmacionRINFO(app, "RINFO PAPELERIA EJEMPLO SL [11-11-2026].xlsx",
                                                                win_tarifas + "\\RINFO PAPELERIA EJEMPLO SL [11-11-2026].xlsx",
                                                                "PAPELERIA EJEMPLO SL", "1008888", [], 101),
         lambda v: [(B(v, "Valores escritos").master, 1, "ir"), (boton(v, "Sí, abrir ahora"), 2, "t")]),
        ("c_consulta_manual", lambda: VentanaConsultaManual(app, "códigos SAP faltantes", app.log_consola),
         lambda v: [(de_clase(v, "Entry"), 1, "ir"), (de_clase(v, "Canvas"), 2, "ir"), (boton(v, "Abrir el manual"), 3, "tr"), (boton(v, "Guardar una copia"), 4, "tl")]),
        ("c_actualizacion", lambda: VentanaActualizacion(app, {"version": "10.2.42", "notas": "## What's Changed\n* Mejoras en la "
                                                         "consulta del manual by @x in #12\n* Correcciones menores by @x in #13",
                                                         "url_exe": "", "url_sha": ""}, tarifa_app.APP_VERSION, app.log_consola),
         lambda v: [(de_clase(v, "Text"), 1, "ir"), (boton(v, "Actualizar ahora"), 2, "t"), (boton(v, "Más tarde"), 3, "t")]),
        ("c_principal", None, marcas_principal),
    ]
    return lista


PRODUCTOS_DEMO = {"10509011": ("BAYETA MICROFIBRA AZUL 40X40", "Pieza", "PI", "", "", ""),
                  "10509012": ("GUANTE NITRILO M CAJA 100", "Caja", "CJ", "", "", "")}
DIF_DEMO = {
    "accion2": [{"codigo_dali": 10041, "campos": {"nombre_articulo": (None, "TIJERA OFICINA 21 CM")}, "nuevo": True,
                 "editado_en_web": False, "proveedor": "PAPELERIA EJEMPLO SL"},
                {"codigo_dali": 10002, "campos": {"nombre_articulo": ("CARPETA A4", "CARPETA ARCHIVADOR A4")}, "nuevo": False,
                 "editado_en_web": False, "proveedor": "PAPELERIA EJEMPLO SL"}],
    "accion3": [{"codigo_dali": 10031, "campos": {"descripcion_sap": ("ROTULADOR NEGRO", "ROTULADOR PERMANENTE NEGRO")},
                 "nuevo": False, "editado_en_web": True, "proveedor": "PAPELERIA EJEMPLO SL"}],
    "accion4": [{"codigo_dali": 10001, "campos": {"precio": (4.10, 4.35)}, "nuevo": False, "editado_en_web": False,
                 "proveedor": "PAPELERIA EJEMPLO SL"},
                {"codigo_dali": 10041, "campos": {"precio": (None, 2.90)}, "nuevo": True, "editado_en_web": False,
                 "proveedor": "PAPELERIA EJEMPLO SL"}]}
SIMS_DEMO = {"accion4": {"irregulares": [{"codigo_dali": 10003, "tipo": "codigo_proveedor_distinto",
                                          "detalle": "El código de proveedor guardado es «P-500» y la tarifa trae «P-501»: no se toca."}]}}


def _purchasing_con_aviso(app, cri, tarifa):
    """(23 nov 2026) «Purchasing encontrado» con el aviso de Purchasing desactualizado (datos inventados)."""
    import datetime as _dt
    from unittest import mock
    ejemplo = {"faltan": [("10509003", "5300009991"), ("10509004", "5300009992")],
               "rinfo_posteriores": [("RINFO SUMINISTROS DEMO SL [10-11-2026].xlsx", _dt.datetime(2026, 11, 10, 9, 15))]}
    with mock.patch.object(cri.CP, "comprobar", return_value=ejemplo):
        return cri.VentanaPurchasingEncontrado(
            app, tarifa, os.path.basename(tarifa), "SUMINISTROS DEMO SL", "SUMINISTROS DEMO SL", "1009999",
            {"ruta": os.path.join(TARIFAS, "Purchasing Info Records.xlsx"), "archivo": "Purchasing Info Records.xlsx",
             "score": 100, "proveedor": "SUMINISTROS DEMO SL"}, TARIFAS, app)


def marcas_asignar_sap(v):
    B = buscar
    filas = v._filas_visibles()
    cab = lambda col: v._rejilla.grid_slaves(row=0, column=col)[0]   # noqa: E731  (cabecera de la columna)
    v._vars[filas[0]].set("10509011")            # en Productos.xlsx
    v._vars[filas[1]].set("10509099")            # no está en Productos.xlsx: se avisa, pero se guarda
    v._vars[filas[2]].set("10509012")
    v._celdas_sap[filas[3]].focus_set()
    v.update()
    return [
        (caja_de(v, [B(v, "Buscar"), B(v, "Quitar filtros")]), 1, "tr"),
        (caja_de(v, [cab(0), cab(4)]), 2, "ir"),
        (caja_de(v, [cab(5)] + [v._celdas_sap[d] for d in filas[:4]]), 3, "b"),
        (caja_de(v, [cab(6)] + [v._estados[d] for d in filas[:3]]), 4, "b"),
        (v._btn_guardar, 5, "in"),
    ]


def _abrir_volcado_dali(app):
    """Final de la Acción 1 (13 nov 2026), ya contestado «Sí» para que se vea el botón (sin abrir Excel de verdad)."""
    from app.comun import volcado_dali
    from app.ui.ventana_volcado_dali import VentanaVolcadoDali
    lineas = ["✔  Proceso completado correctamente", "", "   Artículos transferidos:  4812", "",
              "✘  2 código(s) faltante(s) entre 000001 y 010045:",
              "   Estos códigos de ARTICULOS1 no están en el resultado.", "      010012  (1 ausente)",
              "      010030  (1 ausente)", "", "📁  Puede acceder al listado en:", "   " + RUTA_WIN + "\\RECURSOS", "",
              "🗑  Archivos eliminados automáticamente de RECURSOS/:", "   • ARTICULOS1.xls", "   • ARTICULOS2.xls"]
    v = VentanaVolcadoDali(app, "DALI — Proceso completado con errores", volcado_dali.lineas_con_tipo(lineas),
                           os.path.join(RECURSOS, "ARTICULOS_DALI.xlsx"), abrir=lambda _r: None)
    v._btn_si.invoke()
    app.wait_window(v)


def marcas_volcado_dali(v):
    B = buscar
    return [
        (B(v, "Proceso completado correctamente").master, 1, "ir"),
        (caja_de(v, [v._btn_si, v._btn_no]), 2, "r"),
        (v._btn_actualizar, 3, "t"),
        (v._btn_cerrar, 4, "t"),
    ]


def _antes_accion3(app, Ventana, L):
    """(v 10.2.84) «Acción 3 — Listados de SAP», con un diagnóstico inventado: Productos.xlsx de hace 9 días y dos
    códigos SAP que los listados de RECURSOS no tienen."""
    import time as _t
    ahora = _t.time()
    fechas = {L.PRODUCTOS: ahora - 9 * 86400, L.LISTAS: ahora - 3600}
    fuera = ["10509098", "10509099"]
    diag = {"estado": L.FALTAN, "pendientes": 2, "fuera": fuera,
            "fuera_por_listado": {L.PRODUCTOS: fuera, L.LISTAS: fuera[1:]},
            "antiguos": [L.PRODUCTOS], "no_estan": [], "preguntar": True}
    return Ventana(app, diag, L.lineas_diagnostico(diag, fechas, ahora), RECURSOS, leer=lambda *a: [])


def marcas_antes_accion3(v):
    B = buscar
    return [
        (B(v, "Productos.xlsx: descargado"), 1, "r"),
        (B(v, "código(s) SAP no están"), 2, "r"),
        (B(v, "Sí, abrir RECURSOS", "Button"), 3, "b"),
        (B(v, "No", "Button"), 4, "b"),
        (B(v, "1.  En SAP"), 5, "r"),
        (B(v, "Cancelar", "Button"), 6, "l"),
    ]


def marcas_listados_sap(v):
    B = buscar
    for w in widgets(v):                          # la ruta de ejemplo, como se ve en Windows
        if w.winfo_class() == "Label" and texto(w).strip() == RECURSOS:
            w.configure(text="   " + RUTA_WIN + "\\RECURSOS")
    v.update()
    return [
        (B(v, "Has asignado"), 1, "ir"),
        (B(v, "1.  En SAP"), 2, "r"),
        (B(v, "2.  Pulsa"), 3, "l"),
        (B(v, "3.  Nombre de fichero"), 4, "l"),
        (B(v, "Ya están en RECURSOS", "Button"), 5, "ir"),
        (B(v, "Más tarde", "Button"), 6, "ir"),
    ]


def marcas_principal(v):
    B = buscar
    v._carpeta_proyecto.set(RUTA_WIN)          # solo lo que se ve; las ventanas anteriores ya se capturaron
    v._estado_conexion_web().marcar_ok()       # (18 nov 2026) el indicador de la cabecera, en verde
    v._pintar_indicador_web()
    v.update(); v.update_idletasks()
    ent = de_clase(v, "Entry", 0)
    prog = next((w for w in widgets(v) if texto(w) == "TARIFA.xls no encontrada"), None)
    return [
        (ent.master if ent else None, 1, "ir"),
        (caja_de(v, [w for w in widgets(v) if texto(w).startswith(("TARIFAS", "RECURSOS")) or "RAÍZ" in texto(w)]), 2, "r"),
        (B(v, "Crear estructura PROYECTO"), 3, "ir"),
        (B(v, "Localizar").master.master if B(v, "Localizar") else None, 4, "tr"),
        (prog, 5, "r"),
        (caja_de(v, [B(v, "Organizar artículos DALI").master, B(v, "Generar plantilla Tarifa").master]), 6, "tr"),
        (B(v, "Consultar el manual").master, 7, "tr"),
        (v.consola, 8, "ir"),
        (caja_de(v, [B(v, "guardar log"), B(v, "limpiar")]), 9, "b"),
        (B(v, "ACCESOS RÁPIDOS").master, 10, "t"),
        (B(v, "v 10.", n=0), 11, "l"),
        (getattr(v, "_lbl_web", None), 12, "l"),
    ]


# ── (22 nov 2026) Capturas de Excel: la TARIFA y la plantilla de revisión de precios ─────────────────────
_COLORES_MARCA = [(255, 0, 255), (0, 255, 255), (255, 128, 0), (0, 200, 0), (0, 0, 255), (200, 0, 0)]


def _foto_excel(ruta, nombre, titulo, marcas, ajustar=None):
    """Guarda <nombre>.png con el libro `ruta` tal como se ve en Excel (renderizado con LibreOffice, con la
    cuadrícula). marcas: [(rango, lado)]. Para situarlas, se renderiza antes una copia con cada rango pintado de un
    color único y se busca ese color en la imagen (el relleno no cambia la geometría)."""
    import openpyxl
    from openpyxl.styles import PatternFill
    tmp = tempfile.mkdtemp(prefix="excel_manual_")

    def copia(destino, pintar):
        wb = openpyxl.load_workbook(ruta)
        ws = wb.active
        if ajustar:
            ajustar(ws)
        if pintar:
            for (rango, _lado), rgb in zip(marcas, _COLORES_MARCA):
                for fila in ws[rango]:
                    for c in fila:
                        c.fill = PatternFill(fill_type="solid", fgColor="%02X%02X%02X" % rgb)
        ws.print_area = f"A1:L{min(ws.max_row, 16)}"
        ws.page_setup.orientation = "landscape"
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 1
        ws.print_options.gridLines = True
        wb.save(destino)

    def renderizar(xlsx):
        subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, xlsx],
                       check=True, capture_output=True, timeout=120)
        pdf = os.path.splitext(xlsx)[0] + ".pdf"
        subprocess.run(["pdftoppm", "-r", "200", "-png", "-singlefile", pdf, pdf[:-4]], check=True, timeout=120)
        return Image.open(pdf[:-4] + ".png").convert("RGB")

    def caja_de_color(img, rgb):
        mask = Image.new("L", img.size)
        mask.putdata([255 if abs(p[0] - rgb[0]) < 12 and abs(p[1] - rgb[1]) < 12 and abs(p[2] - rgb[2]) < 12 else 0
                      for p in img.getdata()])
        return mask.getbbox()

    copia(os.path.join(tmp, "marcas.xlsx"), True)
    copia(os.path.join(tmp, "final.xlsx"), False)
    con_marcas, img = renderizar(os.path.join(tmp, "marcas.xlsx")), renderizar(os.path.join(tmp, "final.xlsx"))
    cajas = [caja_de_color(con_marcas, rgb) for rgb in _COLORES_MARCA[:len(marcas)]]
    recorte = Image.eval(img.convert("L"), lambda v: 255 if v < 250 else 0).getbbox()
    x0, y0 = max(0, recorte[0] - 30), max(0, recorte[1] - 30)
    img = img.crop((x0, y0, min(img.width, recorte[2] + 30), min(img.height, recorte[3] + 30)))
    img, dy = _barra_titulo(img, titulo)
    d = ImageDraw.Draw(img)
    for n, (caja, (_rango, lado)) in enumerate(zip(cajas, marcas), start=1):
        if caja is None:
            log(f"   ⚠ {nombre}: no se encontró la marca {n}")
            continue
        _marca(d, (caja[0] - x0, caja[1] - y0 + dy, caja[2] - x0, caja[3] - y0 + dy), n, lado, img.width, img.height)
    img.save(os.path.join(SALIDA, nombre + ".png"))
    log(f"  ✓ {nombre}")


def captura_tarifa_cabecera(nombre="c_tarifa_cabecera"):
    """La cabecera de una TARIFA final de EJEMPLO (la genera transformar() de verdad, con un .xls inventado)."""
    if SOLO and not nombre.startswith(SOLO):
        return
    from unittest import mock
    from app.tarifas import transformar_tarifa as TT

    class _Hoja:
        class _C:
            def __init__(self, v):
                self.value = v

        def __init__(self, celdas):
            self._c, self.nrows, self.ncols = celdas, max(r for r, _ in celdas), 10

        def cell(self, r, c):
            return self._C(self._c.get((r + 1, c + 1), ""))

    celdas = {(2, 4): 4321, (4, 4): "SUMINISTROS DEMO SL", (8, 4): "", (10, 4): "",
              (17, 3): "01/01/2027 - 31/12/2027"}
    for i, (desc, prec) in enumerate((("ACEITE DE OLIVA DEMO 1L", 4.35), ("VINAGRE DE EJEMPLO 1L", 1.20),
                                      ("SAL MARINA MODELO 1KG", 0.85), ("AZUCAR DEMO 1KG", 1.10))):
        celdas.update({(24 + i, 1): 90001 + i, (24 + i, 2): f"REF-{i + 1:03d}", (24 + i, 4): desc,
                       (24 + i, 6): "UD", (24 + i, 7): prec})
    tmp = tempfile.mkdtemp(prefix="tarifa_manual_")
    os.makedirs(os.path.join(tmp, "TARIFAS"))
    xls = os.path.join(tmp, "TARIFAS", "TARIFA.xls")
    open(xls, "wb").close()
    with mock.patch.object(TT, "_leer_xls", return_value=_Hoja(celdas)):
        ruta, _ = TT.transformar(xls)

    def rellenar(ws):
        ws["C4"] = "928 000 000"
        for i in range(4):                                # códigos SAP y Registros Info inventados (A y H)
            ws.cell(8 + i, 1).value = 10100001 + i        # (sin catálogo de ejemplo, el paso 27b no los pone)
            ws.cell(8 + i, 8).value = 5300000101 + i
        for c in ("F2", "F5", "H2", "H3", "H4"):          # GY, TA, MG, GC, MT
            ws[c] = "✔"

    _foto_excel(ruta, nombre, "SUMINISTROS DEMO SL [22-11-2026].xlsx - Excel", [("B4:C4", "r"), ("E1:H6", "r")],
                rellenar)


def captura_plantilla_tarifa(nombre="c_plantilla_tarifa"):
    """(22 nov 2026) La plantilla de revisión de precios de la Acción 6, generada de verdad (plantilla_rinfo +
    plantilla_tarifa) con datos inventados de un Purchasing: un proveedor DEMO con dos grupos de hoteles."""
    if SOLO and not nombre.startswith(SOLO):
        return
    import datetime as _dt
    from collections import OrderedDict
    from app.comun import plantilla_rinfo as PR
    from app.tarifas.plantilla_tarifa import escribir_plantilla
    ini, fin = _dt.date(2026, 1, 1), _dt.date(2026, 12, 31)
    todos = ["0110", "0120", "0210", "0310", "0410", "0510", "0610", "0710", "0810", "0910"]
    materiales = OrderedDict()
    for mat, ref, texto, precio in (("10100001", "REF-001", "ACEITE OLIVA DEMO 1L", 4.35),
                                    ("10100002", "REF-002", "VINAGRE EJEMPLO 1L", 1.20),
                                    ("10100003", "REF-003", "SAL MARINA MODELO 1KG", 0.85),
                                    ("10100004", "REF-004", "AZUCAR DEMO 1KG", 1.10)):
        materiales[mat] = {"texto": texto, "um": "UN", "cod_prov_mat": ref,
                           "precios": {c: (precio - (0.25 if c in ("0810", "0910") and mat == "10100001" else 0),
                                           ini, fin) for c in todos}}
    registros = {m: str(5300000101 + i) for i, m in enumerate(materiales)}   # Registros Info inventados
    datos = {"proveedores": OrderedDict([("1009999", {"codigo": "1009999", "nombre": "SUMINISTROS DEMO SL",
                                                      "telefono": "928000000", "materiales": materiales,
                                                      "registros": registros})])}
    catalogo = PR.mapa_catalogo([(90001 + i, t, "UD", m) for i, (m, t) in
                                 enumerate((("10100001", "ACEITE DE OLIVA DEMO 1L"),
                                            ("10100002", "VINAGRE DE EJEMPLO 1L"),
                                            ("10100003", "SAL MARINA MODELO 1KG"),
                                            ("10100004", "AZUCAR DEMO 1KG")))])
    tmp = tempfile.mkdtemp(prefix="plantilla_manual_")
    pls = PR.preparar_plantillas(datos, catalogo, fecha=_dt.date(2026, 11, 22))
    ruta = escribir_plantilla(pls[0], tmp)
    _foto_excel(ruta, nombre, f"{pls[0]['nombre_archivo']} - Excel",
                [("C3:D3", "ir"), ("B4:C4", "ir"), ("E1:H6", "r"), ("F8:F11", "b"), ("H8:H11", "b"), ("I7:J11", "b")])


def main():
    os.makedirs(SALIDA, exist_ok=True)
    for _captura in (captura_tarifa_cabecera, captura_plantilla_tarifa):
        try:
            _captura()
        except Exception:  # noqa: BLE001
            log(f"   ✗ {_captura.__name__}: {traceback.format_exc()}")
    app = tarifa_app.App()
    app.geometry("1280x820+0+0")
    pendientes = []

    def preparar():
        app.consola.limpiar() if hasattr(app.consola, "limpiar") else None
        for t, n in LOG_DEMO:
            if n is None and hasattr(app.consola, "separador"):
                app.consola.separador(t.strip("─ "))
            else:
                app.log_consola(t, n or "INFO")
        pendientes.extend(pasos(app))
        app.after(800, siguiente)

    def siguiente():
        if not pendientes:
            log("FIN")
            app.after(300, app.destroy)
            return
        nombre, abrir, marcas = pendientes.pop(0)
        if SOLO and not nombre.startswith(SOLO):
            app.after(10, siguiente)
            return
        if abrir is None:
            try:
                foto(app, nombre, marcas(app))
            except Exception:  # noqa: BLE001
                log(f"   ✗ {nombre}: {traceback.format_exc()}")
            app.after(300, siguiente)
            return
        antes = {str(w) for w in app.winfo_children() if isinstance(w, tk.Toplevel)}

        def capturar():
            v = ultima(app, antes)
            if v is None:
                log(f"   ✗ {nombre}: no se abrió ninguna ventana")
            else:
                try:
                    foto(v, nombre, marcas(v))
                except Exception:  # noqa: BLE001
                    log(f"   ✗ {nombre}: {traceback.format_exc()}")
                try:
                    v.grab_release()
                except tk.TclError:
                    pass
                v.destroy()
            app.after(400, siguiente)
        app.after(1500, capturar)
        try:
            abrir()            # puede bloquear (wait_window) hasta que capturar() cierre la ventana
        except Exception:  # noqa: BLE001
            log(f"   ✗ {nombre} al abrir: {traceback.format_exc()}")

    app.after(2500, preparar)
    app.mainloop()
    with open(os.path.join(SALIDA, "capturas.log"), "w", encoding="utf-8") as f:
        f.write("\n".join(LOG))


if __name__ == "__main__":
    main()

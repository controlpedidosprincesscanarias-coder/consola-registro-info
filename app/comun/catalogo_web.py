# -*- coding: utf-8 -*-
"""
catalogo_web.py — la Consola escribe en el catálogo de «Catálogo Asignaciones» (etapa 6, fase 1, v10.2.50).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).

Ver docs/propuestas/etapa6-consola-escribe-catalogo.md en el repositorio de Catálogo Asignaciones
(dali-sap-articulos-app) y su v1.20.63 (POST /admin/catalogo/lote, GET /admin/catalogo/cambios).

Cómo funciona (sin tkinter: se prueba en tests/test_catalogo_web.py):
1. La Consola guarda en el PC una COPIA del catálogo de la web (`Cache`, en %LOCALAPPDATA%) con la hora de la
   última sincronización, y solo pide lo que ha cambiado desde entonces (petición de Víctor: «que la comunicación
   con Supabase sea solo la necesaria»).
2. Después de las Acciones 2, 3 y 4 (y de resolver un SAP duplicado), el catálogo LOCAL
   (RECURSOS/catalogo-dali-completo.xlsx, que la Consola sigue escribiendo como siempre: fase 1, doble escritura)
   se compara con esa copia. `calcular_diferencias` devuelve SOLO lo distinto, en tres grupos:
     - «accion2»: datos DALI (incluidas las altas de artículos nuevos);
     - «accion3»: datos SAP;
     - «accion4»: precio y código de proveedor, agrupados por proveedor asignado (la web solo los aplica si el
       proveedor coincide, y si el código de proveedor guardado es otro, avisa y no lo toca: decisión de Víctor).
   Lo que no se haya podido enviar (sin conexión, «Ahora no»…) vuelve a salir la próxima vez, porque la diferencia
   sigue ahí: no hace falta guardar pendientes aparte.
3. Se envía primero con «simular» (la web dice qué cambiaría) y, si se confirma, se aplica, por lotes.

Nunca se borra nada en la web: un artículo que está en la web y no en el Excel local se deja como está.
"""
import http.cookiejar
import json
import math
import os
import urllib.error
import urllib.request
from datetime import datetime

from app.comun.sap_helpers import CATALOGO_FILA_CABECERA, mapa_columnas_catalogo

DALI_APP_API_BASE = "https://dali-proxy-api.centralcompras1-canarias.workers.dev"
USER_AGENT = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
              "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36")
FILAS_POR_LOTE = 400          # la web acepta hasta 1.000; 400 cabe de sobra en su límite de 1 MB

# Clave de cada columna del catálogo local (sap_helpers.CATALOGO_COLUMNAS) → nombre en la web.
CLAVE_WEB = {
    "clave": "codigo_dali", "cod_prov": "codigo_proveedor", "descripcion": "nombre_articulo",
    "unidad": "unidad_dali", "igic": "igic", "naturaleza": "naturaleza", "familia": "familia",
    "subfamilia": "subfamilia", "proveedor": "proveedor", "cod_sap": "codigo_sap",
    "desc_sap": "descripcion_sap", "unidad_sap": "unidad_sap", "texto_um": "texto_um_sap",
    "precio": "precio", "grp_prod": "grupo_productos_sap", "desc_grp_prod": "descripcion_grupo_sap",
    "ind_imp": "indicador_impuestos_sap", "categoria": "categoria_valoracion", "tipo_prod": "tipo_producto",
    "grp_compras": "grupo_compras_sap", "desc_grp_compras": "nombre_grupo_compras_sap", "activo": "activo",
}
GRUPOS = {
    "accion2": ("nombre_articulo", "unidad_dali", "igic", "naturaleza", "familia", "subfamilia", "proveedor", "activo"),
    "accion3": ("codigo_sap", "descripcion_sap", "unidad_sap", "texto_um_sap", "grupo_productos_sap",
                "descripcion_grupo_sap", "indicador_impuestos_sap", "categoria_valoracion", "tipo_producto",
                "grupo_compras_sap", "nombre_grupo_compras_sap"),
    "accion4": ("precio", "codigo_proveedor"),
}
# Solo se cargan desde la Consola (no se editan en la web): siempre manda lo de la Consola.
SOLO_DESDE_CONSOLA = ("nombre_articulo",)
NOMBRE_GRUPO = {"accion2": "Datos DALI", "accion3": "Datos SAP", "accion4": "Precio y código de proveedor"}


# ── Valores: misma normalización que la web (controllers/catalogo/lote.js) ─────────────────────────────
def texto(v):
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)                       # 10100416.0 (Excel) → "10100416"
    t = str(v).strip()
    return t or None


def precio(v):
    if v is None or str(v).strip() == "":
        return None
    try:
        n = float(str(v).strip().replace(",", "."))
        x = n * 100
        r = math.floor(x)
        return (r + 1 if x - r >= 0.5 else r) / 100     # igual que Math.round de la web (no redondeo bancario)
    except ValueError:
        return None


def activo(v):
    if isinstance(v, bool):
        return v
    return str(v or "").strip().upper() in ("SI", "SÍ", "S", "TRUE", "1")


def normalizar(campo, v):
    if campo == "precio":
        return precio(v)
    if campo == "activo":
        return activo(v)
    return texto(v)


def codigo_dali(v):
    try:
        n = int(float(str(v).strip()))
        return n if n > 0 else None
    except (ValueError, TypeError):
        return None


# ── Catálogo local (el Excel que regenera la Consola) ───────────────────────────────────────────────────
def leer_catalogo_local(ruta):
    """{codigo_dali: {campo_web: valor normalizado}} de RECURSOS/catalogo-dali-completo.xlsx.
    Las columnas se buscan por cabecera (fila 4), como hace el resto de la Consola."""
    import openpyxl
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    try:
        ws = wb.active
        filas = ws.iter_rows(min_row=CATALOGO_FILA_CABECERA, values_only=True)
        cabecera = next(filas, None) or ()
        mapa = mapa_columnas_catalogo(cabecera)          # clave local → nº de columna (base 1)
        indices = {CLAVE_WEB[k]: i - 1 for k, i in mapa.items() if k in CLAVE_WEB}
        if "codigo_dali" not in indices:
            raise ValueError("el catálogo local no tiene la columna «Código DALI» en la fila 4")
        catalogo = {}
        for fila in filas:
            if not fila:
                continue
            cod = codigo_dali(fila[indices["codigo_dali"]] if indices["codigo_dali"] < len(fila) else None)
            if cod is None:
                continue
            catalogo[cod] = {campo: normalizar(campo, fila[i] if i < len(fila) else None)
                             for campo, i in indices.items() if campo != "codigo_dali"}
        return catalogo
    finally:
        wb.close()


# ── Copia local de la web (se guarda en el PC) ─────────────────────────────────────────────────────────
def ruta_cache_por_defecto():
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    return os.path.join(base, "ConsolaRegistroInfo", "catalogo_web.json")


class Cache:
    """Copia del catálogo de la web: {hasta, articulos: {codigo_dali: {campo: valor}}}."""

    def __init__(self, ruta=None):
        self.ruta = ruta or ruta_cache_por_defecto()
        self.hasta = None
        self.articulos = {}

    def cargar(self):
        try:
            with open(self.ruta, encoding="utf-8") as f:
                datos = json.load(f)
            self.hasta = datos.get("hasta")
            self.articulos = {int(k): v for k, v in (datos.get("articulos") or {}).items()}
        except (OSError, ValueError, TypeError, AttributeError):
            self.hasta, self.articulos = None, {}
        return self

    def guardar(self):
        os.makedirs(os.path.dirname(self.ruta), exist_ok=True)
        tmp = self.ruta + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump({"hasta": self.hasta, "articulos": {str(k): v for k, v in self.articulos.items()}}, f,
                      ensure_ascii=False)
        os.replace(tmp, self.ruta)

    def aplicar_pagina(self, articulos):
        for a in articulos:
            cod = codigo_dali(a.get("codigo_dali"))
            if cod is None:
                continue
            self.articulos[cod] = {campo: normalizar(campo, a.get(campo))
                                   for campo in set().union(*GRUPOS.values())} | {"modificado_en": a.get("modificado_en")}

    def sincronizar(self, cliente, progreso=None):
        """Trae de la web lo cambiado desde la última vez (o todo, la primera). Devuelve cuántos artículos."""
        desde, pagina, total, hasta = self.hasta, 1, 0, None
        while True:
            datos = cliente.cambios(desde, pagina)
            hasta = hasta or datos.get("hasta")
            articulos = datos.get("articulos") or []
            self.aplicar_pagina(articulos)
            total += len(articulos)
            if progreso:
                progreso(total)
            if not datos.get("hay_mas"):
                break
            pagina += 1
        self.hasta = hasta
        self.guardar()
        return total


# ── Diferencias ───────────────────────────────────────────────────────────────────────────────────────
def calcular_diferencias(local, web, desde_sincronizacion=None):
    """Lo que difiere entre el catálogo local y la copia de la web, por grupo.

    Devuelve {grupo: [ {codigo_dali, campos: {campo: (web, local)}, nuevo, editado_en_web, proveedor} ]}.
    - Solo los campos distintos. Un artículo que no está en la web es «nuevo»: va entero en accion2 (alta) y,
      en accion3 y accion4, con los campos que tenga rellenos (se aplican después del alta, en ese orden).
    - accion3 y accion4 de un artículo nuevo y todo accion4: un valor VACÍO en local no borra el de la web
      (puede que esa tarifa o ese listado no se haya pasado nunca por esta Consola).
    - editado_en_web: el artículo cambió en la web (o lo envió otro compañero) después de
      `desde_sincronizacion` (la sincronización anterior).
    - El código DALI y su descripción solo se cargan desde la Consola (decisión de Víctor, 12 nov 2026):
      la descripción local siempre manda, aunque el artículo se haya tocado en la web.
    """
    resultado = {g: [] for g in GRUPOS}
    for cod in sorted(local):
        loc = local[cod]
        w = web.get(cod)
        editado = bool(w and desde_sincronizacion and (w.get("modificado_en") or "") > desde_sincronizacion)
        for grupo, campos in GRUPOS.items():
            distintos = {}
            for campo in campos:
                vl = loc.get(campo)
                vw = None if w is None else w.get(campo)
                if vl is None and (grupo == "accion4" or (w is None and grupo == "accion3")):
                    continue
                if w is None and grupo == "accion2":
                    distintos[campo] = (None, vl)
                elif vl != vw:
                    distintos[campo] = (vw, vl)
            if distintos:
                resultado[grupo].append({"codigo_dali": cod, "campos": distintos, "nuevo": w is None,
                                         "editado_en_web": editado, "proveedor": loc.get("proveedor")})
    return resultado


def filas_para_enviar(diferencias, grupo, incluir_editados=False):
    """Filas para POST /admin/catalogo/lote (solo los campos distintos). Para accion4, agrupadas por proveedor:
    {proveedor: [filas]}; para el resto, {None: [filas]}."""
    por_proveedor = {}
    for d in diferencias[grupo]:
        campos = d["campos"]
        if d["editado_en_web"] and not incluir_editados and not d["nuevo"]:
            campos = {c: v for c, v in campos.items() if c in SOLO_DESDE_CONSOLA}
            if not campos:
                continue
        fila = {"codigo_dali": d["codigo_dali"]}
        for campo, (_web, local) in campos.items():
            fila[campo] = ("SI" if local else "NO") if campo == "activo" else local
        clave = d["proveedor"] if grupo == "accion4" else None
        if grupo == "accion4" and not clave:
            continue                              # sin proveedor asignado no hay tarifa con la que cuadrar
        por_proveedor.setdefault(clave, []).append(fila)
    return por_proveedor


def en_lotes(filas, tamano=FILAS_POR_LOTE):
    for i in range(0, len(filas), tamano):
        yield filas[i:i + tamano]


# ── Conexión con la web ───────────────────────────────────────────────────────────────────────────────
class ErrorWeb(Exception):
    """Error con un mensaje para el usuario. `codigo` es el HTTP (o None si no hubo respuesta)."""

    def __init__(self, mensaje, codigo=None):
        super().__init__(mensaje)
        self.codigo = codigo


class ClienteCatalogo:
    """Sesión con Catálogo Asignaciones (cookie), igual que la descarga del catálogo de siempre."""

    def __init__(self, base=DALI_APP_API_BASE, abrir=None):
        self.base = base.rstrip("/")
        jar = http.cookiejar.CookieJar()
        opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(jar))
        opener.addheaders = [("User-Agent", USER_AGENT)]
        self._abrir = abrir or opener.open

    def _peticion(self, metodo, ruta, cuerpo=None, timeout=90):
        datos = json.dumps(cuerpo).encode("utf-8") if cuerpo is not None else None
        req = urllib.request.Request(self.base + ruta, data=datos, method=metodo,
                                     headers={"Content-Type": "application/json"} if datos else {})
        try:
            with self._abrir(req, timeout=timeout) as r:
                return json.loads(r.read().decode("utf-8") or "{}")
        except urllib.error.HTTPError as e:
            try:
                detalle = json.loads(e.read().decode("utf-8")).get("error", "")
            except Exception:  # noqa: BLE001
                detalle = ""
            if e.code == 403 and ruta.startswith("/admin/catalogo"):
                detalle = ("tu cuenta no tiene el permiso «Escribir desde la Consola Registro Info» en Catálogo "
                           "Asignaciones (pídeselo al administrador principal)")
            raise ErrorWeb(f"HTTP {e.code}: {detalle or e.reason}", e.code)
        except Exception as e:  # noqa: BLE001
            raise ErrorWeb(f"sin conexión con Catálogo Asignaciones ({e})")

    def login(self, email, password):
        datos = self._peticion("POST", "/auth/login", {"email": email, "password": password}, timeout=60)
        if datos.get("requiereVerificacion"):
            raise ErrorWeb("Catálogo Asignaciones pide un código de verificación: inicia sesión una vez desde el "
                           "navegador y vuelve a intentarlo.", 428)
        if not datos.get("ok"):
            raise ErrorWeb(datos.get("error") or "inicio de sesión rechazado", 401)
        return datos

    def cambios(self, desde=None, pagina=1):
        from urllib.parse import urlencode
        q = {"pagina": pagina}
        if desde:
            q["desde"] = desde
        return self._peticion("GET", "/admin/catalogo/cambios?" + urlencode(q))

    def lote(self, origen, filas, simular, proveedor_tarifa=None, archivo=None, sincronizado_desde=None):
        cuerpo = {"origen": origen, "simular": bool(simular), "filas": filas}
        if proveedor_tarifa:
            cuerpo["proveedor_tarifa"] = proveedor_tarifa
        if archivo:
            cuerpo["archivo"] = archivo
        if sincronizado_desde:
            cuerpo["sincronizado_desde"] = sincronizado_desde
        return self._peticion("POST", "/admin/catalogo/lote", cuerpo)


# ── Enviar un grupo entero (simular o aplicar), por lotes ─────────────────────────────────────────────
def enviar_grupo(cliente, grupo, por_proveedor, simular, archivo=None, sincronizado_desde=None):
    """Envía las filas de un grupo. Devuelve el resumen sumado y las listas de cambios, irregulares y avisos."""
    total = {"recibidas": 0, "nuevos": 0, "cambiados": 0, "sin_cambios": 0, "irregulares": 0, "editados_en_web": 0}
    cambios, irregulares, avisos = [], [], []
    for proveedor, filas in por_proveedor.items():
        for lote in en_lotes(filas):
            r = cliente.lote(grupo, lote, simular, proveedor_tarifa=proveedor if grupo == "accion4" else None,
                             archivo=archivo, sincronizado_desde=sincronizado_desde)
            for k in total:
                total[k] += int((r.get("resumen") or {}).get(k) or 0)
            cambios.extend(r.get("cambios") or [])
            irregulares.extend(r.get("irregulares") or [])
            avisos.extend(r.get("avisos") or [])
    return {"resumen": total, "cambios": cambios, "irregulares": irregulares, "avisos": avisos}


def ahora_texto():
    return datetime.now().strftime("%d/%m/%Y %H:%M")

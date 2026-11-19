# -*- coding: utf-8 -*-
"""
listados_sap.py — aviso «faltan los listados de SAP actualizados» tras asignar códigos SAP (13 nov 2026).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).

Petición de Víctor: cuando se asignan códigos SAP (Acción 8 o la ventana de códigos SAP faltantes de la Acción 4),
la Consola envía los códigos a Catálogo Asignaciones, pero los datos SAP de un material nuevo (descripción, unidad,
grupos, categoría de valoración) solo llegan con Productos.xlsx y Listas de productos.xlsx descargados de SAP
después de crearlo. La Consola tiene que avisarlo, guiar la descarga y, cuando estén en RECURSOS, lanzar sola la
Acción 3; si no se hace ahora, la web queda parcialmente actualizada hasta entonces.

Sin tkinter: la ventana está en app/ui/ventana_listados_sap.py. Pruebas en tests/test_listados_sap.py.
"""
import os

PRODUCTOS = "Productos.xlsx"
LISTAS = "Listas de productos.xlsx"

# Cómo se descarga cada listado (aplicación de SAP y nombre con el que se exporta)
LISTADOS = (
    {"archivo": PRODUCTOS, "app_sap": "Gestionar datos maestros de producto", "nombre": "Productos",
     "imagen": "assets/guia_sap_exportar_productos.png"},
    {"archivo": LISTAS, "app_sap": "Lista de productos", "nombre": "Listas de productos",
     "imagen": "assets/guia_sap_exportar_listas.png"},
)

PASOS = (
    "En SAP, abre la aplicación del listado (imagen de arriba).",
    "Pulsa «Ir» para cargar la tabla y, a la derecha de su barra, el botón ⌄ de exportar (el último, «Abrir menú») "
    "→ «Exportar como…».",
    "Deja el nombre de fichero que se indica y el formato «Microsoft Excel (*.xlsx)», y pulsa «Exportar».",
    "Guarda el archivo descargado en la carpeta RECURSOS del proyecto, sustituyendo al anterior.",
)

AVISO_PARCIAL = ("La web (Catálogo Asignaciones) queda PARCIALMENTE actualizada: ya tiene los códigos SAP, pero no "
                 "sus datos SAP (descripción, unidad, grupos, categoría de valoración) hasta que descargues "
                 "Productos.xlsx y Listas de productos.xlsx en RECURSOS y pulses la Acción 3.")


def rutas(carpeta_recursos):
    """{nombre de archivo: ruta} de los dos listados en RECURSOS."""
    return {L["archivo"]: os.path.join(carpeta_recursos, L["archivo"]) for L in LISTADOS}


def codigos_sin_datos(codigos_sap, tabla_productos):
    """Los códigos SAP (como texto, sin repetir y en orden) que no están en Productos.xlsx: los que se quedarán sin
    sus datos SAP hasta descargar los listados actualizados. Sin Productos.xlsx (tabla None o vacía), todos."""
    vistos, fuera = set(), []
    for c in codigos_sap:
        txt = str(c).strip()
        if not txt or txt in vistos:
            continue
        vistos.add(txt)
        if not tabla_productos or txt not in tabla_productos:
            fuera.append(txt)
    return fuera


def estado_archivos(carpeta_recursos, fechas_antes):
    """Para cada listado: (existe, actualizado) — actualizado = descargado de nuevo desde que salió el aviso.
    fechas_antes: fechas_listados() de cuando salió el aviso. (v10.2.76) Se compara con la fecha que tenía el propio
    archivo, no con la hora del PC (descargado_de_nuevo), igual que antes de la Acción 3."""
    return {nombre: (ahora is not None, descargado_de_nuevo(fechas_antes.get(nombre), ahora))
            for nombre, ahora in fechas_listados(carpeta_recursos).items()}


def sin_actualizar(estado):
    """Nombres de los listados que faltan o no se han descargado de nuevo."""
    return [n for n, (existe, actualizado) in estado.items() if not (existe and actualizado)]


# ── Antes de la Acción 3 (v10.2.74, petición de Víctor, 13 nov 2026) ───────────────────────────────────────────────
# «¿La aplicación podría saber si necesitas actualizar los dos Excel del paso 3? Normalmente se necesita porque se han
# insertado los códigos SAP pero no tiene el resto de información.» Antes de la Acción 3 se miran los artículos del
# catálogo con código SAP y sin sus datos SAP, y si esos códigos ya están en los listados de RECURSOS. Además, como no
# se puede saber si en SAP cambió algo de un material que ya tiene datos, se avisa si un listado tiene más de
# DIAS_ANTIGUEDAD días (7, decisión de Víctor del 19 nov 2026). La ventana está en app/ui/ventana_listados_sap.py
# (VentanaAntesAccion3).

DIAS_ANTIGUEDAD = 7

# Datos SAP que tiene que tener un artículo con código SAP (claves de catalogo_local.leer_catalogo):
# los de Productos.xlsx (descripción, unidad, grupo de productos) y la categoría de Listas de productos.xlsx.
CAMPOS_PRODUCTOS = ("desc_sap", "unidad_sap", "grp_prod")
CAMPO_LISTAS = "categoria"

# Resultado de la comprobación
NADA = "nada"                    # ningún artículo con código SAP sin sus datos
EN_RECURSOS = "en_recursos"      # faltan datos, pero los listados de RECURSOS ya los tienen
FALTAN = "faltan"                # algún código no está en los listados de RECURSOS: hay que descargarlos

PREGUNTA_ACCION3 = "¿Necesita descargar de nuevo Productos.xlsx y Listas de productos.xlsx?"
TEXTO_BOTON = ("Pulsa «Actualizar Catálogo Asignaciones» para completar los datos SAP del catálogo con los listados "
               "de RECURSOS y enviarlos a la web.")
TEXTO_DESCARGANDO = ("Se ha abierto la carpeta RECURSOS. Descarga los dos listados de SAP (guía de abajo), guárdalos "
                     "ahí sustituyendo a los anteriores y pulsa «Actualizar Catálogo Asignaciones».")


def _texto_codigo(v):
    """10104081, 10104081.0 o «10104081 » → «10104081»; vacío → «»."""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip() if v is not None else ""


def _vacio(v):
    return v is None or str(v).strip() == ""


def articulos_sin_datos_sap(filas_catalogo):
    """Los artículos del catálogo con código SAP y sin alguno de sus datos SAP.
    filas_catalogo: [(clave, {campo: valor})] como catalogo_local.leer_catalogo.
    Devuelve {código SAP: {"productos": bool, "listas": bool}} — qué listado tiene que traer lo que falta."""
    salida = {}
    for _clave, d in filas_catalogo:
        cod = _texto_codigo(d.get("cod_sap"))
        if not cod:
            continue
        falta_prod = any(_vacio(d.get(c)) for c in CAMPOS_PRODUCTOS)
        falta_listas = _vacio(d.get(CAMPO_LISTAS))
        if falta_prod or falta_listas:
            previo = salida.get(cod, {"productos": False, "listas": False})
            salida[cod] = {"productos": previo["productos"] or falta_prod,
                           "listas": previo["listas"] or falta_listas}
    return salida


def _ordenar(codigos):
    return sorted(codigos, key=lambda c: (len(c), c))


def codigos_fuera_por_listado(pendientes, tabla_productos, tabla_categoria):
    """{nombre del listado: códigos de `pendientes` (articulos_sin_datos_sap) que ese listado tendría que completar
    y no tiene}: a Productos.xlsx le tocan los que no tienen descripción, unidad o grupo; a Listas de productos, los
    que no tienen categoría de valoración. Sin un listado (tabla None), todos los que le tocan. Ordenados."""
    prod, listas = [], []
    for cod, falta in pendientes.items():
        if falta["productos"] and cod not in (tabla_productos or {}):
            prod.append(cod)
        if falta["listas"] and cod not in (tabla_categoria or {}):
            listas.append(cod)
    return {PRODUCTOS: _ordenar(prod), LISTAS: _ordenar(listas)}


def fechas_listados(carpeta_recursos):
    """{nombre: fecha de modificación (timestamp) o None si no está} de los dos listados."""
    fechas = {}
    for nombre, ruta in rutas(carpeta_recursos).items():
        try:
            fechas[nombre] = os.path.getmtime(ruta)
        except OSError:
            fechas[nombre] = None
    return fechas


def listados_antiguos(fechas, ahora, dias=DIAS_ANTIGUEDAD):
    """Nombres de los listados con más de `dias` días (los que no están no cuentan aquí: ya faltan)."""
    return [n for n, f in fechas.items() if f is not None and ahora - f > dias * 86400]


def diagnostico(filas_catalogo, tabla_productos, tabla_categoria, fechas, ahora, dias=DIAS_ANTIGUEDAD):
    """Comprobación antes de la Acción 3. Devuelve un dict:
       estado: NADA / EN_RECURSOS / FALTAN;
       pendientes: n.º de códigos SAP sin todos sus datos; fuera: los que los listados no tienen (lista);
       fuera_por_listado: {nombre: códigos que ese listado tendría que traer y no tiene};
       antiguos: listados con más de `dias` días; no_estan: listados que no están en RECURSOS;
       preguntar: si hay que preguntar «¿Necesita descargar…?» (faltan códigos, un listado no está o es antiguo)."""
    pendientes = articulos_sin_datos_sap(filas_catalogo)
    por_listado = codigos_fuera_por_listado(pendientes, tabla_productos, tabla_categoria)
    fuera = _ordenar(set(por_listado[PRODUCTOS]) | set(por_listado[LISTAS]))
    no_estan = [n for n, f in fechas.items() if f is None]
    antiguos = listados_antiguos(fechas, ahora, dias)
    if not pendientes:
        estado = NADA
    elif fuera:
        estado = FALTAN
    else:
        estado = EN_RECURSOS
    return {"estado": estado, "pendientes": len(pendientes), "fuera": fuera, "fuera_por_listado": por_listado,
            "antiguos": antiguos,
            "no_estan": no_estan, "preguntar": bool(fuera or no_estan or antiguos)}


def _fecha_texto(ts, ahora):
    from datetime import datetime
    if ts is None:
        return "no está en RECURSOS"
    dias = int((ahora - ts) // 86400)
    hace = "hoy" if dias <= 0 else "ayer" if dias == 1 else f"hace {dias} días"
    return f"descargado el {datetime.fromtimestamp(ts).strftime('%d/%m/%Y %H:%M')} ({hace})"


def lineas_diagnostico(diag, fechas, ahora, dias=DIAS_ANTIGUEDAD):
    """Texto de la ventana: [(texto, tipo)] con tipo ok / aviso / error / info."""
    lineas = []
    for nombre in (PRODUCTOS, LISTAS):
        ts = fechas.get(nombre)
        tipo = "error" if ts is None else "aviso" if nombre in diag["antiguos"] else "info"
        lineas.append((f"📄 {nombre}: {_fecha_texto(ts, ahora)}", tipo))
    lineas.append(("", "info"))
    n, fuera = diag["pendientes"], diag["fuera"]
    if diag["estado"] == NADA:
        lineas.append(("✔ Todos los artículos con código SAP tienen ya sus datos SAP: no falta nada por completar.",
                       "ok"))
    elif diag["estado"] == EN_RECURSOS:
        lineas.append((f"✔ Los listados de RECURSOS ya tienen los datos de los {n} artículo(s) con código SAP que "
                       "faltan por completar.", "ok"))
    else:
        muestra = ", ".join(fuera[:12]) + (" …" if len(fuera) > 12 else "")
        lineas.append((f"✘ {len(fuera)} código(s) SAP no están en los listados de RECURSOS, así que se quedarían "
                       f"sin descripción, unidad, grupos o categoría de valoración: {muestra}", "error"))
        if n > len(fuera):
            lineas.append((f"   (los listados sí tienen los otros {n - len(fuera)} que faltan por completar)",
                           "info"))
    if diag["antiguos"]:
        lineas.append((f"⚠ {' y '.join(diag['antiguos'])} tiene(n) más de {dias} días: si en SAP ha cambiado algo "
                       "(una descripción, una unidad…), conviene descargarlos de nuevo.", "aviso"))
    if diag["no_estan"]:
        lineas.append((f"✘ No está en RECURSOS: {', '.join(diag['no_estan'])}. Hay que descargarlo de SAP.",
                       "error"))
    return lineas


def descargado_de_nuevo(fecha_antes, fecha_ahora):
    """El listado ha cambiado desde que se contestó «Sí»: su fecha no es la que tenía entonces (o antes no estaba).
    Se compara con la fecha del propio archivo, no con la hora del PC: en la unidad de red (G:) la fecha la pone el
    reloj del servidor, que puede ir unos segundos por delante o por detrás."""
    return fecha_ahora is not None and fecha_ahora != fecha_antes


def estado_descarga(carpeta_recursos, fechas_antes, fuera_por_listado, contiene):
    """Para la línea de estado mientras se descargan los listados (cada 2 s).
    fechas_antes: fechas_listados() de cuando se contestó «Sí».
    fuera_por_listado: {nombre: códigos que ese listado tiene que traer} (de diagnostico).
    contiene: {nombre: códigos de esos que el archivo nuevo todavía NO tiene, o ausente si aún no se ha leído}.
    Devuelve [(nombre, listo, texto)]: listo si se ha descargado de nuevo y, si tenía que traer códigos, ya los tiene
    (no basta con que cambie su fecha)."""
    salida = []
    for nombre, (existe, actualizado) in estado_archivos(carpeta_recursos, fechas_antes).items():
        fuera = fuera_por_listado.get(nombre) or []
        if not existe:
            salida.append((nombre, False, "no está en RECURSOS"))
        elif not actualizado:
            salida.append((nombre, False, "sin descargar de nuevo"))
        elif not fuera:
            salida.append((nombre, True, "descargado de nuevo"))
        else:
            faltan = contiene.get(nombre)
            if faltan is None:
                salida.append((nombre, False, "descargado de nuevo, comprobando…"))
            elif faltan:
                salida.append((nombre, False, f"descargado de nuevo, pero le faltan {len(faltan)} código(s)"))
            else:
                salida.append((nombre, True, f"descargado de nuevo, con los {len(fuera)} código(s)"))
    return salida


def codigos_que_faltan_en(nombre, ruta, fuera):
    """Lee el listado `nombre` y devuelve los códigos de `fuera` que no tiene. Para Productos.xlsx basta con estar;
    para Listas de productos.xlsx, con tener Categoría valoración. Lanza la excepción de openpyxl si no se puede leer
    (por ejemplo, a medio copiar)."""
    from app.comun.sap_helpers import leer_listas_productos, leer_productos_xlsx
    if nombre == PRODUCTOS:
        tabla, _n = leer_productos_xlsx(ruta)
    else:
        tabla = leer_listas_productos(ruta)[0]
    return [c for c in fuera if c not in tabla]

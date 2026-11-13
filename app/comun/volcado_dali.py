# -*- coding: utf-8 -*-
"""
volcado_dali.py — final de la Acción 1: revisar ARTICULOS_DALI.xlsx y actualizar el catálogo y Catálogo Asignaciones
(13 nov 2026).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).

Petición de Víctor: al terminar la Acción 1, que la información se vuelque directamente en la web, preguntando
antes si se quiere revisar el Excel limpio (Sí / No); con «Sí» se abre el Excel y, en la misma ventana, aparece el
botón para actualizar Catálogo Asignaciones, oculto hasta responder. Con esto desaparece la Acción 2: lo que hacía
(pasar ARTICULOS_DALI.xlsx al catálogo completo, por código DALI) lo hace ese botón.

Sin tkinter: la ventana está en app/ui/ventana_volcado_dali.py. Pruebas en tests/test_volcado_dali.py.
"""
import os

from app.comun.sap_helpers import abierto_en_excel

ARTICULOS_DALI = "ARTICULOS_DALI.xlsx"
FILA_PRIMEROS_DATOS = 8          # ARTICULOS_DALI.xlsx: cabecera hasta la fila 7 (la generada por la Acción 1)

PREGUNTA = "¿Desea revisar ARTICULOS_DALI.xlsx en Excel antes de actualizar Catálogo Asignaciones?"
TEXTO_REVISANDO = ("Se ha abierto ARTICULOS_DALI.xlsx en Excel. Revisa y corrige lo que necesites, guarda y "
                   "cierra Excel, y pulsa «Actualizar Catálogo Asignaciones».")
TEXTO_SIN_REVISAR = "Pulsa «Actualizar Catálogo Asignaciones» para pasar los artículos al catálogo y a la web."
AVISO_MAS_TARDE = ("No se ha actualizado nada: ni catalogo-dali-completo.xlsx ni Catálogo Asignaciones.\n\n"
                   "ARTICULOS_DALI.xlsx se queda en RECURSOS. Cuando quieras, pulsa otra vez la Acción 1: "
                   "ofrecerá usar ese archivo para actualizar.")

# Estados del Excel al pulsar «Actualizar Catálogo Asignaciones»
ABIERTO = "abierto"              # sigue abierto en Excel: no se puede leer lo guardado con seguridad
NO_EXISTE = "no_existe"          # ya no está en RECURSOS (borrado o renombrado)
SIN_GUARDAR = "sin_guardar"      # se abrió para revisar, pero no se ha guardado ningún cambio
LISTO = "listo"


def fecha_modificacion(ruta):
    """Fecha de modificación del archivo (segundos) o None si no existe."""
    try:
        return os.path.getmtime(ruta)
    except OSError:
        return None


def estado_para_actualizar(ruta, revisado, mtime_al_abrir=None):
    """Cómo está ARTICULOS_DALI.xlsx al pulsar «Actualizar Catálogo Asignaciones».

    revisado: el usuario contestó «Sí» (se abrió en Excel). mtime_al_abrir: fecha de modificación cuando se abrió.
    SIN_GUARDAR solo es un aviso (puede que no hubiera nada que corregir); ABIERTO y NO_EXISTE impiden seguir."""
    if not os.path.isfile(ruta):
        return NO_EXISTE
    if abierto_en_excel(ruta):
        return ABIERTO
    if revisado and mtime_al_abrir is not None and fecha_modificacion(ruta) == mtime_al_abrir:
        return SIN_GUARDAR
    return LISTO


MENSAJES_ESTADO = {
    ABIERTO: ("ARTICULOS_DALI.xlsx sigue abierto en Excel.\n\n"
              "Guarda los cambios, cierra Excel y vuelve a pulsar «Actualizar Catálogo Asignaciones»."),
    NO_EXISTE: ("Ya no se encuentra ARTICULOS_DALI.xlsx en RECURSOS.\n\n"
                "Vuelve a ejecutar la Acción 1 con ARTICULOS1.xls y ARTICULOS2.xls."),
    SIN_GUARDAR: ("No se ha guardado ningún cambio en ARTICULOS_DALI.xlsx desde que se abrió.\n\n"
                  "¿Actualizar Catálogo Asignaciones con el archivo tal como está?"),
}


def contar_articulos(ruta):
    """Número de artículos (filas con código en la columna A desde la fila 8) de ARTICULOS_DALI.xlsx.
    Es el recuento que hacía la Acción 2 antes de pasar el archivo al catálogo."""
    import openpyxl
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    try:
        n = 0
        for fila in wb.active.iter_rows(min_row=FILA_PRIMEROS_DATOS, max_col=1, values_only=True):
            if fila and fila[0] not in (None, ""):
                n += 1
        return n
    finally:
        wb.close()


def fecha_encabezado(ruta):
    """Texto de la fecha de la cabecera (celda B3) de ARTICULOS_DALI.xlsx, como lo mostraba la Acción 2."""
    import openpyxl
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    try:
        b3 = wb.active["B3"].value
    finally:
        wb.close()
    return str(b3).strip() if b3 else "(fecha no disponible)"


def lineas_con_tipo(lineas):
    """Color de cada línea del resumen de la Acción 1 en la ventana: (texto, tipo).

    El bloque de códigos faltantes (desde la línea que empieza por «✘» hasta la siguiente cabecera «📁» o «🗑») va
    como "error"; «✔» como "ok"; las cabeceras «📁»/«🗑» y lo sangrado que las sigue, como "nota"; lo demás, "info"."""
    salida, bloque = [], "info"
    for texto in lineas:
        if texto.startswith("✘"):
            bloque = "error"
        elif texto.startswith(("📁", "🗑")):
            bloque = "nota"
        elif bloque == "nota" and not texto.startswith(" "):
            bloque = "info"
        if bloque == "info" and texto.startswith("✔"):
            salida.append((texto, "ok"))
        else:
            salida.append((texto, bloque))
    return salida

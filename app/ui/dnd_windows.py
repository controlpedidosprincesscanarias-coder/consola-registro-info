"""
app/ui/dnd_windows.py — arrastrar y soltar archivos del Explorador de Windows sobre una ventana de tkinter, sin
instalar nada: usa la API nativa del sistema (WM_DROPFILES) con ctypes. Si no es Windows o algo falla, no hace nada y
no rompe la ventana.

    hook_dropfiles(ventana, func)   # func(lista_de_rutas) se llama en el hilo de tkinter

(22 nov 2026) Copiado tal cual del Organizador Princess (app/ui/dnd_windows.py, v2.0 de su v4.18.02, en uso en el chat
desde entonces) para la Acción 6 «Generar plantilla Tarifa desde RINFO». Lo importante de esa versión:
  · usa SetWindowSubclass (la API documentada para esto, que se desengancha sola al destruir la ventana) y los
    callbacks no se liberan nunca;
  · solo se engancha la ventana de nivel superior (el marco de Tk y su área cliente), no cada widget;
  · dentro del callback nativo NO se toca tkinter: las rutas se dejan en una cola y un temporizador (after) las
    recoge en el hilo normal.

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación son de
Víctor Alejandro Martín Andrade (VAMA).
"""

import sys
import queue
import logging

_log = logging.getLogger("consola")
_WM_DROPFILES = 0x233
_VIVOS = []            # callbacks ctypes: nunca se liberan (la ventana puede seguir llamándolos hasta WM_NCDESTROY)
_ENGANCHADOS = set()   # hwnd con el gancho puesto ahora mismo
_IDENT = 0x4F50        # identificador del subclase («OP»)
_SONDEO_MS = 150


def _api():
    import ctypes
    from ctypes import wintypes
    shell32, comctl32 = ctypes.windll.shell32, ctypes.windll.comctl32
    LRESULT = ctypes.c_ssize_t
    SUBCLASSPROC = ctypes.WINFUNCTYPE(LRESULT, ctypes.c_void_p, ctypes.c_uint, ctypes.c_size_t,
                                      ctypes.c_ssize_t, ctypes.c_size_t, ctypes.c_size_t)
    comctl32.SetWindowSubclass.argtypes = [ctypes.c_void_p, SUBCLASSPROC, ctypes.c_size_t, ctypes.c_size_t]
    comctl32.SetWindowSubclass.restype = wintypes.BOOL
    comctl32.DefSubclassProc.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_size_t, ctypes.c_ssize_t]
    comctl32.DefSubclassProc.restype = LRESULT
    shell32.DragAcceptFiles.argtypes = [ctypes.c_void_p, wintypes.BOOL]
    shell32.DragQueryFileW.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_wchar_p, ctypes.c_uint]
    shell32.DragQueryFileW.restype = ctypes.c_uint
    shell32.DragFinish.argtypes = [ctypes.c_void_p]
    return ctypes, shell32, comctl32, SUBCLASSPROC


def _rutas_de_drop(ctypes, shell32, hdrop):
    n = shell32.DragQueryFileW(hdrop, 0xFFFFFFFF, None, 0)
    rutas = []
    for i in range(min(n, 50)):
        largo = shell32.DragQueryFileW(hdrop, i, None, 0)
        buf = ctypes.create_unicode_buffer(largo + 1)
        shell32.DragQueryFileW(hdrop, i, buf, largo + 1)
        rutas.append(buf.value)
    return rutas


def _crear_callback(api, cola):
    """Callback nativo del subclase. Solo lee rutas y las pone en la cola: sin tkinter."""
    ctypes, shell32, comctl32, SUBCLASSPROC = api

    def _proc(hwnd, msg, wp, lp, ident, ref):
        try:
            if msg == _WM_DROPFILES:
                rutas = []
                try:
                    rutas = _rutas_de_drop(ctypes, shell32, wp)
                except Exception as exc:   # noqa: BLE001
                    _log.warning("Arrastrar y soltar: no se pudo leer la lista de archivos: %s", exc)
                finally:
                    try:
                        shell32.DragFinish(wp)
                    except Exception:      # noqa: BLE001
                        pass
                if rutas:
                    cola.put(rutas)
                return 0
        except Exception:                  # noqa: BLE001 — nunca dejar escapar una excepción al sistema
            return 0
        return comctl32.DefSubclassProc(hwnd, msg, wp, lp)

    return SUBCLASSPROC(_proc)


def _hwnds_de_nivel_superior(ventana):
    """Marco de Tk (la ventana real de Windows) y su área cliente."""
    hwnds = []
    for obtener in (lambda: int(ventana.frame(), 16), lambda: int(ventana.winfo_id())):
        try:
            h = obtener()
            if h and h not in hwnds:
                hwnds.append(h)
        except Exception:                  # noqa: BLE001
            pass
    return hwnds


def hook_dropfiles(ventana, func):
    """Devuelve True si se pudo activar. Llamar cuando la ventana ya está construida."""
    if not sys.platform.startswith("win"):
        return False
    try:
        ventana.update_idletasks()
        api = _api()
        ctypes, shell32, comctl32, _ = api
        cola = queue.Queue()
        cb = _crear_callback(api, cola)
        _VIVOS.append(cb)                  # nunca se libera
        puestos = []
        for hwnd in _hwnds_de_nivel_superior(ventana):
            if hwnd in _ENGANCHADOS:
                continue
            if comctl32.SetWindowSubclass(hwnd, cb, _IDENT, 0):
                shell32.DragAcceptFiles(hwnd, True)
                _ENGANCHADOS.add(hwnd)
                puestos.append(hwnd)
        if not puestos:
            return False

        estado = {"ocupado": False}

        def _vigilar():
            try:
                if not ventana.winfo_exists():
                    return
                ventana.after(_SONDEO_MS, _vigilar)
            except Exception:              # noqa: BLE001
                return
            if estado["ocupado"]:          # ya hay un diálogo abierto por un arrastre anterior
                return
            estado["ocupado"] = True
            try:
                while True:
                    rutas = cola.get_nowait()
                    try:
                        func(rutas)
                    except Exception as exc:   # noqa: BLE001
                        _log.warning("Arrastrar y soltar: error al procesar los archivos: %s", exc)
            except queue.Empty:
                pass
            finally:
                estado["ocupado"] = False

        ventana.after(_SONDEO_MS, _vigilar)

        def _limpiar(event):
            if event.widget is ventana:
                for h in puestos:
                    _ENGANCHADOS.discard(h)
        ventana.bind("<Destroy>", _limpiar, add="+")
        return True
    except Exception as exc:               # noqa: BLE001
        _log.warning("Arrastrar y soltar no disponible: %s", exc)
        return False

# -*- coding: utf-8 -*-
"""
tarifa_core.py
Recursos compartidos: rutas, configuración persistente (tarifa_config.json, junto al .exe o
a tarifa_app.py), lógica de negocio de archivos y widgets reutilizables de UI. (Los logos están
en assets/ desde nov 2026; ya no hay imágenes en base64 en el código.)
Central Compras Princess Canarias
Autor: Victor Alejandro Martin Andrade (VAMA)
(c) VAMA 2026

Símbolos exportados:
    resource_path(relative_path)    — ruta compatible .py / .exe PyInstaller
    get_config_path / cargar_config / guardar_config
    TARIFA_NAME / FOLDER_TARIFAS / FOLDER_RECURSOS / FOLDER_PROYECTO
    crear_estructura_proyecto / buscar_tarifa
    listar_xls / renombrar_a_tarifa
    ConsolaLog                      — widget Text con colores y scroll
    CanvasProgressBar               — barra de progreso animada con Canvas
"""

import os
import shutil
import tkinter as tk
from tkinter import ttk
from datetime import datetime
import json
import sys

from app.comun.tarifa_constants import (
    resource_path, RAIZ_PROYECTO,
    C_DARK, C_MID, C_SIDEBAR, C_ACCENT, C_MUTED, C_LOG_BG,
    C_LOG_FG, C_LOG_INFO, C_LOG_WARN, C_LOG_ERR, C_LOG_OK,
)





# ---------------------------------------------------------
#  CONFIGURACION PERSISTENTE
# ---------------------------------------------------------
def get_config_path():
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = RAIZ_PROYECTO   # junto a tarifa_app.py, como siempre (este archivo está en app/comun/)
    return os.path.join(base, "tarifa_config.json")


def cargar_config():
    path = get_config_path()
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def guardar_config(config: dict):
    path = get_config_path()
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
    except Exception:
        pass


# ---------------------------------------------------------
#  LOGICA DE NEGOCIO
# ---------------------------------------------------------
TARIFA_NAME      = "TARIFA.xls"
FOLDER_TARIFAS   = "TARIFAS"
FOLDER_RECURSOS  = "RECURSOS"
FOLDER_PROYECTO  = "PROYECTO"


def crear_estructura_proyecto(escritorio):
    """Crea PROYECTO/TARIFAS y PROYECTO/RECURSOS en el escritorio."""
    base = os.path.join(escritorio, FOLDER_PROYECTO)
    tarifas  = os.path.join(base, FOLDER_TARIFAS)
    recursos = os.path.join(base, FOLDER_RECURSOS)
    os.makedirs(tarifas,  exist_ok=True)
    os.makedirs(recursos, exist_ok=True)
    return base


def buscar_tarifa(carpeta_proyecto):
    """Busca TARIFA.xls dentro de la subcarpeta TARIFAS del proyecto."""
    ruta = os.path.join(carpeta_proyecto, FOLDER_TARIFAS, TARIFA_NAME)
    return ruta if os.path.isfile(ruta) else None


def listar_xls(carpeta_tarifas):
    archivos = []
    try:
        for entry in os.scandir(carpeta_tarifas):
            if entry.is_file() and entry.name.lower().endswith(".xls"):
                ts = entry.stat().st_mtime
                fecha = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
                archivos.append({"nombre": entry.name, "ruta": entry.path, "fecha_mod": fecha})
    except PermissionError:
        pass
    archivos.sort(key=lambda x: x["nombre"].lower())
    return archivos


def renombrar_a_tarifa(ruta_origen, carpeta_tarifas):
    destino = os.path.join(carpeta_tarifas, TARIFA_NAME)
    if os.path.isfile(destino):
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup = os.path.join(carpeta_tarifas, f"TARIFA_backup_{ts}.xls")
        shutil.copy2(destino, backup)
    shutil.copy2(ruta_origen, destino)
    # Eliminar el archivo original si es distinto al destino
    if os.path.abspath(ruta_origen) != os.path.abspath(destino):
        try:
            os.remove(ruta_origen)
        except OSError:
            pass
    return destino


# ---------------------------------------------------------
#  BOILERPLATE COMPARTIDO DE VENTANAS Toplevel (punto 6 del roadmap
#  de limpieza, 2026-09-24) -- extraído de las clases VentanaXxxRINFO
#  de crear_registro_info.py y las VentanaXxx de tarifa_dialogs.py,
#  que repetían este mismo código byte a byte. Alcance deliberadamente
#  conservador: solo se extrajo lo que era IDÉNTICO en varias clases;
#  las geometrías con fórmulas de ancho/alto propias (una por caso de
#  uso) se dejaron tal cual, para no forzar una firma común que metiera
#  riesgo real de romper la disposición visual de alguna ventana.
# ---------------------------------------------------------
def cargar_icono_ventana(win):
    """
    Carga el icono estándar (assets/logo.ico) en una ventana Toplevel.
    Idéntico en 13 de las 15 ventanas del proyecto; se ignora en
    silencio si el icono no está disponible (mismo comportamiento que
    tenían todas las copias sueltas).
    """
    try:
        win.iconbitmap(resource_path("assets/logo.ico"))
    except Exception:
        pass


def poner_icono_aplicacion(root):
    """
    (nov 2026) Icono de la ventana principal y de la barra de tareas. Con solo
    iconbitmap(ruta), Windows podía enseñar en la barra de tareas la pluma de Tk:
    se pone el .ico como icono por defecto de todas las ventanas (default=) y,
    además, sus tamaños grandes con iconphoto(True, ...), que es el que usa la
    barra de tareas. Si algo falla se deja el de Tk, como antes.
    """
    ruta = resource_path("assets/logo.ico")
    try:
        root.iconbitmap(default=ruta)
    except Exception:
        pass
    try:
        from PIL import Image, ImageTk
        ico = Image.open(ruta)
        imagenes = []
        for lado in (256, 64, 48, 32, 24, 16):
            try:
                ico.size = (lado, lado)          # elige ese tamaño dentro del .ico
            except (ValueError, KeyError):
                continue
            imagenes.append(ImageTk.PhotoImage(ico.convert("RGBA"), master=root))
        if imagenes:
            root.iconphoto(True, *imagenes)
            root._iconos_app = imagenes          # que no los borre el recolector
    except Exception:
        pass


def centrar_sobre_padre(win, parent):
    """
    Centra una ventana Toplevel sobre el centro de su ventana padre
    (no sobre la pantalla completa). Usa el tamaño ya calculado de
    `win` (llamar después de construir la UI y de update_idletasks()
    si hace falta que winfo_width()/height() sean fiables).

    Idéntico en 4 de las ventanas de tarifa_dialogs.py (Selección XLS,
    Abrir tarifa edición, Crear proyecto, Selección tarifa RI).
    """
    win.update_idletasks()
    pw = parent.winfo_rootx() + parent.winfo_width()  // 2
    ph = parent.winfo_rooty() + parent.winfo_height() // 2
    w, h = win.winfo_width(), win.winfo_height()
    win.geometry(f"+{pw - w // 2}+{ph - h // 2}")


def ajustar_geometria_pantalla(win, ancho, margen_alto=20):
    """
    Centra una ventana Toplevel en la pantalla completa, con un ancho
    fijo (`ancho`) y alto ajustado al contenido real más un margen.

    Idéntico salvo el ancho en 3 ventanas de crear_registro_info.py
    (Teléfono, Fechas y Centros del flujo "crear RINFO desde cero").
    """
    win.update_idletasks()
    w = ancho
    h = win.winfo_reqheight() + margen_alto
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    x = (sw - w) // 2
    y = max(20, (sh - h) // 2)
    win.geometry(f"{w}x{h}+{x}+{y}")


# ---------------------------------------------------------
#  WIDGET: CONSOLA DE LOG
# ---------------------------------------------------------
class ConsolaLog(tk.Frame):

    COLORES = {
        "INFO":    C_LOG_INFO,
        "OK":      C_LOG_OK,
        "WARN":    C_LOG_WARN,
        "ERROR":   C_LOG_ERR,
        "SISTEMA": C_ACCENT,
        ">>":      C_LOG_FG,
    }

    def __init__(self, parent, **kwargs):
        super().__init__(parent, bg=C_LOG_BG, **kwargs)
        self._build()

    def _build(self):
        # Detectar mejor fuente monoespaciada disponible en el sistema
        try:
            import tkinter.font as _tkf
            _fams = _tkf.families()
            _mono = next((f for f in ["Cascadia Code", "Cascadia Mono",
                                       "JetBrains Mono", "Consolas", "Courier New"]
                          if f in _fams), "Courier New")
        except Exception:
            _mono = "Consolas"
        _FONT_LOG = (_mono, 8)

        # ── Cabecera estilo macOS ─────────────────────────────────────
        hdr = tk.Frame(self, bg=C_DARK, pady=5)
        hdr.pack(fill="x")

        # Tres puntos de color (close/minimise/maximise)
        _dot_cnv = tk.Canvas(hdr, bg=C_DARK, width=56, height=14,
                              highlightthickness=0)
        _dot_cnv.pack(side="left", padx=(10, 0))
        for _i, _col in enumerate(["#FF5F57", "#FEBC2E", "#28C840"]):
            _cx = 8 + _i * 18
            _dot_cnv.create_oval(_cx - 5, 2, _cx + 5, 12,
                                  fill=_col, outline="")

        tk.Label(hdr, text="VENTANA INMEDIATA  —  Flujo de operaciones",
                 bg=C_DARK, fg=C_ACCENT,
                 font=("Segoe UI", 8, "bold")).pack(side="left", padx=12)

        btn_limpiar = tk.Label(hdr, text="⟳ limpiar",
                                bg=C_DARK, fg=C_MUTED,
                                font=("Segoe UI", 8), cursor="hand2")
        btn_limpiar.pack(side="right", padx=10)
        btn_limpiar.bind("<Button-1>", lambda e: self.limpiar())
        btn_limpiar.bind("<Enter>",
                          lambda e: btn_limpiar.config(fg=C_ACCENT))
        btn_limpiar.bind("<Leave>",
                          lambda e: btn_limpiar.config(fg=C_MUTED))

        btn_guardar = tk.Label(hdr, text="💾 guardar log",
                                bg=C_DARK, fg=C_MUTED,
                                font=("Segoe UI", 8), cursor="hand2")
        btn_guardar.pack(side="right", padx=(0, 4))
        btn_guardar.bind("<Button-1>", lambda e: self.guardar_log())
        btn_guardar.bind("<Enter>",
                          lambda e: btn_guardar.config(fg=C_ACCENT))
        btn_guardar.bind("<Leave>",
                          lambda e: btn_guardar.config(fg=C_MUTED))

        tk.Frame(self, bg=C_ACCENT, height=1).pack(fill="x")

        # ── Área de texto ─────────────────────────────────────────────
        frm = tk.Frame(self, bg=C_LOG_BG)
        frm.pack(fill="both", expand=True)

        _sty = ttk.Style()
        _sty.configure("Thin.Vertical.TScrollbar",
                        background=C_MID, troughcolor=C_LOG_BG,
                        arrowcolor=C_LOG_BG, bordercolor=C_LOG_BG,
                        arrowsize=0, relief="flat", width=6)
        sb = ttk.Scrollbar(frm, orient="vertical",
                            style="Thin.Vertical.TScrollbar")
        sb.pack(side="right", fill="y")

        self.txt = tk.Text(frm, bg=C_LOG_BG, fg=C_LOG_FG,
                           font=_FONT_LOG, wrap="none",
                           state="disabled", cursor="arrow",
                           yscrollcommand=sb.set,
                           insertbackground=C_LOG_FG,
                           relief="flat", padx=8, pady=4,
                           selectbackground=C_MID, bd=0)
        self.txt.pack(side="left", fill="both", expand=True)
        sb.config(command=self.txt.yview)

        _sty.configure("Thin.Horizontal.TScrollbar",
                        background=C_MID, troughcolor=C_LOG_BG,
                        arrowcolor=C_LOG_BG, bordercolor=C_LOG_BG,
                        arrowsize=0, relief="flat", width=4)
        sbx = ttk.Scrollbar(self, orient="horizontal",
                             style="Thin.Horizontal.TScrollbar")
        sbx.pack(fill="x")
        self.txt.config(xscrollcommand=sbx.set)
        sbx.config(command=self.txt.xview)

        for nivel, color in self.COLORES.items():
            self.txt.tag_configure(nivel, foreground=color)
        self.txt.tag_configure("ts",  foreground="#2A4A6A")
        self.txt.tag_configure("sep", foreground="#1E3A5A")

    def log(self, mensaje, nivel=">>", ts=None):
        """
        ts: timestamp ya formado ("%H:%M:%S"), opcional. Si se omite, se
        genera aquí con datetime.now() -- pero eso captura el instante en
        que este método se EJECUTA, no en el que se LLAMÓ. Desde que
        App.log_consola() encola las llamadas desde hilos de fondo con
        self.after(0, ...) (2026-09-24, fix de thread-safety), ambos
        instantes pueden diferir si hay cola pendiente -- App.log_consola()
        ya captura y pasa el ts correcto por este motivo; el valor por
        defecto queda solo para otros llamadores directos.
        """
        if ts is None:
            ts = datetime.now().strftime("%H:%M:%S")
        self.txt.config(state="normal")
        self.txt.insert("end", f" {ts} ", "ts")
        self.txt.insert("end", f"[{nivel:<7}]", nivel)
        self.txt.insert("end", f"  {mensaje}\n", nivel)
        self.txt.config(state="disabled")
        self.txt.see("end")

    def log_batch(self, lineas_niveles, ts=None):
        """
        Inserta una lista de (mensaje, nivel) en un ÚNICO ciclo state=normal/disabled.
        Evita miles de repaints individuales cuando hay logging masivo.
        lineas_niveles: iterable de (str, str)
        ts: timestamp ya formado, opcional -- mismo motivo que en log().
        """
        if not lineas_niveles:
            return
        if ts is None:
            ts = datetime.now().strftime("%H:%M:%S")
        self.txt.config(state="normal")
        for mensaje, nivel in lineas_niveles:
            self.txt.insert("end", f" {ts} ", "ts")
            self.txt.insert("end", f"[{nivel:<7}]", nivel)
            self.txt.insert("end", f"  {mensaje}\n", nivel)
        self.txt.config(state="disabled")
        self.txt.see("end")

    def separador(self, texto=""):
        if not texto:
            linea = "─" * 58
        else:
            pad = max(0, 54 - len(texto))
            linea = f"── {texto} " + "─" * pad
        self.txt.config(state="normal")
        self.txt.insert("end", "\n " + linea + "\n", "sep")
        self.txt.config(state="disabled")
        self.txt.see("end")

    def limpiar(self):
        self.txt.config(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.config(state="disabled")

    def guardar_log(self):
        """Exporta el contenido actual de la consola a un fichero .txt elegido por el usuario."""
        from tkinter import filedialog, messagebox
        contenido = self.txt.get("1.0", "end").strip()
        if not contenido:
            messagebox.showinfo("Log vacío", "No hay entradas en el log para guardar.")
            return
        ts_nombre = datetime.now().strftime("%Y%m%d_%H%M%S")
        ruta = filedialog.asksaveasfilename(
            title="Guardar log",
            initialfile=f"log_{ts_nombre}.txt",
            defaultextension=".txt",
            filetypes=[("Texto plano", "*.txt"), ("Todos los archivos", "*.*")],
        )
        if not ruta:
            return
        try:
            with open(ruta, "w", encoding="utf-8") as f:
                f.write(contenido)
            messagebox.showinfo("Log guardado", f"Log exportado correctamente a:\n{ruta}")
        except Exception as exc:
            messagebox.showerror("Error al guardar", f"No se pudo guardar el log:\n{exc}")


# ---------------------------------------------------------
#  WIDGET: BARRA DE PROGRESO — Canvas custom (animación suave)
# ---------------------------------------------------------
class CanvasProgressBar(tk.Frame):
    """Barra de progreso dibujada con Canvas. Animación suave en modo indeterminado."""
    _TRACK_BG    = "#0D1F35"
    _FILL_CLR    = C_ACCENT
    _DONE_CLR    = "#27AE60"
    _BLOCK_RATIO = 0.32
    _STEP_PX     = 8
    _FPS_MS      = 25

    def __init__(self, parent, **kwargs):
        super().__init__(parent, bg=C_SIDEBAR, **kwargs)
        self._anim_pos = 0
        self._anim_job = None
        self._mode     = "determinate"
        self._value    = 0
        self._maximum  = 100
        self._build()

    def _build(self):
        self._lbl_estado = tk.Label(self, text="Listo",
                                     bg=C_SIDEBAR, fg=C_MUTED,
                                     font=("Segoe UI", 8), anchor="w")
        self._lbl_estado.pack(fill="x")
        frm = tk.Frame(self, bg=C_SIDEBAR)
        frm.pack(fill="x", pady=(2, 0))
        self._canvas = tk.Canvas(frm, height=8, bg=C_SIDEBAR,
                                  highlightthickness=0, bd=0)
        self._canvas.pack(side="left", fill="x", expand=True)
        self._canvas.bind("<Configure>", lambda e: self._redraw())
        self._lbl_pct = tk.Label(frm, text="", bg=C_SIDEBAR,
                                  fg=C_ACCENT, font=("Segoe UI", 8, "bold"),
                                  width=5, anchor="e")
        self._lbl_pct.pack(side="right", padx=(4, 0))

    def _redraw(self):
        c = self._canvas
        c.delete("all")
        w, h = c.winfo_width(), c.winfo_height()
        if w < 4:
            return
        c.create_rectangle(0, 1, w, h - 1, fill=self._TRACK_BG, outline="")
        if self._mode == "determinate":
            ratio = max(0.0, min(1.0, self._value / self._maximum)) if self._maximum else 0
            fw = int(w * ratio)
            if fw > 0:
                c.create_rectangle(0, 1, fw, h - 1,
                                   fill=self._FILL_CLR, outline="")
        elif self._mode == "indeterminate":
            bw  = max(int(w * self._BLOCK_RATIO), 30)
            pos = int(self._anim_pos) % (w + bw) - bw
            x0  = max(0, pos)
            x1  = min(w, pos + bw)
            if x1 > x0:
                c.create_rectangle(x0, 1, x1, h - 1,
                                   fill=self._FILL_CLR, outline="")
        elif self._mode == "complete":
            c.create_rectangle(0, 1, w, h - 1, fill=self._DONE_CLR, outline="")

    def _animate(self):
        if self._mode != "indeterminate":
            return
        w  = self._canvas.winfo_width() or 200
        bw = max(int(w * self._BLOCK_RATIO), 30)
        self._anim_pos = (self._anim_pos + self._STEP_PX) % (w + bw)
        self._redraw()
        self._anim_job = self.after(self._FPS_MS, self._animate)

    def _stop(self):
        if self._anim_job:
            self.after_cancel(self._anim_job)
            self._anim_job = None

    # ── API pública (misma que BarraProgreso) ─────────────────────
    def set(self, valor, maximo=100, estado=""):
        self._stop()
        self._mode    = "determinate"
        self._value   = valor
        self._maximum = maximo
        pct = int(valor / maximo * 100) if maximo else 0
        self._lbl_pct.config(text=f"{pct}%", fg=C_ACCENT)
        if estado:
            self._lbl_estado.config(text=estado)
        self._redraw()

    def indeterminate(self, estado="Procesando..."):
        self._stop()
        self._mode     = "indeterminate"
        self._anim_pos = 0
        self._lbl_estado.config(text=estado)
        self._lbl_pct.config(text="·  ·  ·", fg=C_MUTED)
        self._animate()

    def reset(self, estado="Listo"):
        self._stop()
        self._mode  = "determinate"
        self._value = 0
        self._lbl_pct.config(text="", fg=C_ACCENT)
        self._lbl_estado.config(text=estado)
        self._redraw()

    def completado(self, estado="Completado"):
        self._stop()
        self._mode = "complete"
        self._lbl_pct.config(text="✔", fg=self._DONE_CLR)
        self._lbl_estado.config(text=estado, fg=self._DONE_CLR)
        self._redraw()
        # Restaurar color texto estado tras un momento
        self.after(3000, lambda: self._lbl_estado.config(fg=C_MUTED))


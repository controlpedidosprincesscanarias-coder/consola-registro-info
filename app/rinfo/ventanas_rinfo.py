# -*- coding: utf-8 -*-
"""
ventanas_rinfo.py
Ventanas del motor RINFO que no dependen de otras: plantilla ausente, materiales sin tarifa,
materiales para reactivar, confirmaciones, instrucciones SAP, teléfono, fechas y centros.

Separado de crear_registro_info.py en la etapa 4 (nov 2026) sin cambiar el código.
Central Compras Princess Canarias · La idea y el desarrollo de la aplicación son de
Víctor Alejandro Martín Andrade (VAMA).
"""

import os
import re
import subprocess
import webbrowser
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from app.comun.tarifa_constants import (
    resource_path,
    C_MID, C_BG, C_TEXT, C_TEXT_LIGHT, C_HOVER, C_ACCENT2, C_ERROR,
    FONT_TITLE, FONT_SUB, FONT_LABEL,
    FONT_BOLD, FONT_SMALL,
)
from app.comun.tarifa_core import cargar_icono_ventana, ajustar_geometria_pantalla
from app.comun.selector_fecha import boton_calendario  # (13 nov 2026) calendario de las fechas de validez
from app.comun import fechas_tarifa as FT  # (13 nov 2026) fechas propuestas según el periodo de la tarifa
# normalizar_codigo_sap vive en sap_helpers.py (sin tkinter) para que
# también pueda importarla test_crear_registro_info.py sin arrastrar
# la UI. Ver ese fichero para la implementación y el porqué.
from app.rinfo.comun_rinfo import (  # noqa: F401
    URL_SAP,
    _CENTROS_LIST,
    _fila_valor,
)


# ─────────────────────────────────────────────────────────────────────────────
#  VENTANA INFORMATIVA: TEMPLATE SAP AUSENTE
# ─────────────────────────────────────────────────────────────────────────────

class VentanaTemplateAusente(tk.Toplevel):
    """
    Informa al usuario de que no se encontró el template SAP e indica
    exactamente qué archivo debe colocar y dónde.
    La ventana padre permanece visible para que pueda reintentar.
    """

    def __init__(self, parent, ruta_recursos: str, nombre_template: str):
        super().__init__(parent)
        self.title("Plantilla SAP no encontrada")
        self.configure(bg=C_BG)
        self.resizable(False, False)
        cargar_icono_ventana(self)
        self._ruta_recursos   = ruta_recursos
        self._nombre_template = nombre_template
        self._construir_ui()
        self._ajustar_geometria()
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.grab_set()
        self.transient(parent)

    def _ajustar_geometria(self):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w  = min(580, sw - 40)
        h  = self.winfo_reqheight() + 20
        x  = (sw - w) // 2
        y  = max(20, (sh - h) // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _construir_ui(self):
        hdr = tk.Frame(self, bg="#8B1A1A", pady=12, padx=16)
        hdr.pack(fill="x")
        tk.Label(hdr, text="\u274c  Plantilla SAP no encontrada",
                 bg="#8B1A1A", fg="white", font=FONT_TITLE).pack(anchor="w")
        tk.Label(hdr,
                 text="No se puede crear el Registro Info sin el archivo original de SAP.",
                 bg="#8B1A1A", fg="#F4AAAA", font=FONT_SUB).pack(anchor="w", pady=(4, 0))

        cuerpo = tk.Frame(self, bg=C_BG, padx=22, pady=16)
        cuerpo.pack(fill="both", expand=True)

        tk.Label(cuerpo,
                 text=("Para crear un Registro Info es imprescindible disponer de un\n"
                       "archivo \'Purchasing Info Records\' original descargado de SAP.\n"
                       "Este archivo se usa como plantilla y garantiza que el RINFO\n"
                       "mantiene exactamente la estructura y formato que espera SAP."),
                 bg=C_BG, fg=C_TEXT, font=FONT_LABEL,
                 justify="left").pack(anchor="w", pady=(0, 12))

        tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=(0, 10))

        tk.Label(cuerpo, text="Archivo necesario:",
                 bg=C_BG, fg=C_TEXT, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(cuerpo, text=f"  \U0001f4c4  {self._nombre_template}",
                 bg="#FFF3F3", fg="#8B1A1A",
                 font=("Segoe UI", 10, "bold"),
                 padx=10, pady=7, relief="flat").pack(fill="x", pady=(4, 10))

        tk.Label(cuerpo, text="Debe colocarlo en la carpeta:",
                 bg=C_BG, fg=C_TEXT, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(cuerpo, text=f"  \U0001f4c1  {self._ruta_recursos}",
                 bg="#F5F5F5", fg="#333",
                 font=FONT_LABEL, padx=10, pady=6, relief="flat",
                 wraplength=520, justify="left").pack(fill="x", pady=(4, 10))

        tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=(0, 10))

        tk.Label(cuerpo, text="¿Cómo obtener el archivo?",
                 bg=C_BG, fg=C_TEXT, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(cuerpo,
                 text=("  1.  Use el botón \'Descargar desde SAP\' para obtener\n"
                       "       el Purchasing Info Records original.\n"
                       "  2.  Copie el archivo descargado a la carpeta RECURSOS indicada.\n"
                       "  3.  Vuelva a pulsar \'Crear nuevo Registro Info\'."),
                 bg=C_BG, fg="#555", font=FONT_LABEL,
                 justify="left").pack(anchor="w", pady=(4, 12))

        tk.Button(cuerpo, text="Entendido — Cerrar",
                  command=self.destroy,
                  bg=C_MID, fg=C_TEXT_LIGHT,
                  font=FONT_BOLD, padx=14, pady=10,
                  cursor="hand2", relief="flat",
                  activebackground=C_HOVER,
                  activeforeground=C_ACCENT2).pack(fill="x")

# ─────────────────────────────────────────────────────────────────────────────
#  VENTANA: Materiales heredados sin tarifa — decisión de baja
# ─────────────────────────────────────────────────────────────────────────────

def _ventana_sin_tarifa(parent, pendientes: list) -> list:
    """
    Muestra una ventana modal tkinter para que el usuario decida qué hacer con
    los materiales heredados del Purchasing que NO tienen correspondencia en la
    tarifa activa y cuya fecha de Fin validez (Hoja 3 col J) está aún vigente.

    Parámetros
    ----------
    parent     : ventana padre tkinter (para posicionamiento y grab_set).
    pendientes : lista de dicts con claves:
                    'codigo'  (str)  — código SAP del material
                    'desc'    (str)  — texto breve (Hoja 1 col F)
                    'fecha'   (str)  — Fin validez formateado "dd/mm/yyyy"
                    'fila_h1' (int)  — fila en Hoja 1 (para escritura posterior)

    Devuelve
    --------
    Lista de dicts {'codigo': str, 'fila_h1': int, 'baja': bool}
    con la decisión del usuario para cada material:
        · baja=True   → el usuario marcó el ítem para dar de baja (col AJ = "X")
        · baja=False  → el usuario decidió mantenerlo (col AJ sin cambios)

    Si el usuario cierra la ventana o pulsa CANCELAR, todos los ítems
    devuelven baja=False (comportamiento conservador: no se toca nada).

    Notas de implementación
    -----------------------
    · ttk.Treeview con columna de checkbox simulada mediante texto ("☑"/"☐")
      ya que Treeview nativo no soporta checkboxes reales.
    · Botón [Info] por fila muestra un messagebox con el detalle del ítem.
    · [MANTENER] → todos los ítems sin check quedan baja=False y los
      marcados quedan baja=True; en definitiva es equivalente a [DAR BAJA]
      salvo que el nombre resulta más descriptivo para el flujo.
    · [DAR BAJA]   → solo los ítems con checkbox activo → baja=True.
    · [CANCELAR]   → todos baja=False.
    """

    n = len(pendientes)

    # ── Resultado compartido entre callbacks ──────────────────────────────────
    # key = índice en la lista 'pendientes', value = bool (True = dar de baja)
    checks = {i: False for i in range(n)}
    resultado = []   # se rellena al cerrar con decisión

    # ── Ventana toplevel ──────────────────────────────────────────────────────
    dlg = tk.Toplevel(parent)
    dlg.title(f"Materiales sin tarifa — Confirmar ({n} ítem{'s' if n != 1 else ''})")
    dlg.configure(bg=C_BG)
    dlg.resizable(True, True)
    try:
        dlg.iconbitmap(resource_path("assets/logo.ico"))
    except Exception:
        pass

    # Modal: bloquea interacción con el resto de la app
    dlg.grab_set()
    dlg.focus_set()

    # ── Cabecera ──────────────────────────────────────────────────────────────
    hdr = tk.Frame(dlg, bg="#8B4513", pady=10, padx=16)   # marrón oscuro = alerta
    hdr.pack(fill="x")
    tk.Label(
        hdr,
        text=f"⚠  {n} material{'es' if n != 1 else ''} sin tarifa activa",
        bg="#8B4513", fg="white", font=FONT_TITLE,
    ).pack(anchor="w")
    tk.Label(
        hdr,
        text="Estos materiales existen en el Purchasing pero no están en la tarifa. "
             "Marque los que desea dar de baja (col AJ = 'X').",
        bg="#8B4513", fg="#FFD5A0", font=FONT_SMALL,
        wraplength=620, justify="left",
    ).pack(anchor="w", pady=(4, 0))

    # ── Zona central: Treeview + scrollbar ───────────────────────────────────
    frm_tree = tk.Frame(dlg, bg=C_BG, padx=14, pady=10)
    frm_tree.pack(fill="both", expand=True)

    # Columnas: check (ancho fijo), codigo, desc, fecha, info
    cols = ("check", "codigo", "desc", "fecha", "info")
    tree = ttk.Treeview(
        frm_tree,
        columns=cols,
        show="headings",
        selectmode="none",
        height=min(n, 12),
    )

    # Configurar cabeceras
    tree.heading("check",  text="Baja",        anchor="center")
    tree.heading("codigo", text="Código SAP",  anchor="w")
    tree.heading("desc",   text="Descripción", anchor="w")
    tree.heading("fecha",  text="Fin validez", anchor="center")
    tree.heading("info",   text="",            anchor="center")

    # Anchuras de columna
    tree.column("check",  width=55,  minwidth=45,  stretch=False, anchor="center")
    tree.column("codigo", width=110, minwidth=80,  stretch=False, anchor="w")
    tree.column("desc",   width=280, minwidth=120, stretch=True,  anchor="w")
    tree.column("fecha",  width=110, minwidth=90,  stretch=False, anchor="center")
    tree.column("info",   width=60,  minwidth=50,  stretch=False, anchor="center")

    # Insertar filas
    item_ids = []   # iid por índice
    for i, mat in enumerate(pendientes):
        iid = tree.insert(
            "", "end",
            values=("☐", mat["codigo"], mat["desc"], mat["fecha"], "ℹ Info"),
            tags=("fila_par" if i % 2 == 0 else "fila_impar",),
        )
        item_ids.append(iid)

    # Colores alternos
    tree.tag_configure("fila_par",   background="#F5F8FF")
    tree.tag_configure("fila_impar", background="#FFFFFF")
    tree.tag_configure("marcado",    foreground="#C0392B", font=("Segoe UI", 9, "bold"))

    # Scrollbar vertical
    sb = ttk.Scrollbar(frm_tree, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=sb.set)
    tree.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")

    # ── Toggle checkbox al hacer clic en la columna "check" ──────────────────
    def _toggle_check(event):
        """Alterna el checkbox del ítem clicado (columna 'check' o doble clic)."""
        region = tree.identify_region(event.x, event.y)
        if region not in ("cell", "tree"):
            return
        col_id = tree.identify_column(event.x)
        iid    = tree.identify_row(event.y)
        if not iid:
            return
        # Permitir toggle en cualquier columna excepto "info"
        if col_id == "#5":   # columna info → gestionado por _click_info
            return
        idx = item_ids.index(iid)
        checks[idx] = not checks[idx]
        vals = list(tree.item(iid, "values"))
        vals[0] = "☑" if checks[idx] else "☐"
        tree.item(iid, values=vals,
                  tags=("marcado",) if checks[idx] else
                       ("fila_par" if idx % 2 == 0 else "fila_impar",))

    tree.bind("<Button-1>", _toggle_check)

    # ── Botón ℹ Info por fila ─────────────────────────────────────────────────
    def _click_info(event):
        """Muestra messagebox con el detalle del ítem cuando se pulsa la col info."""
        col_id = tree.identify_column(event.x)
        iid    = tree.identify_row(event.y)
        if not iid or col_id != "#5":
            return
        idx = item_ids.index(iid)
        mat = pendientes[idx]
        messagebox.showinfo(
            f"Detalle material — {mat['codigo']}",
            f"Código SAP:   {mat['codigo']}\n"
            f"Descripción:  {mat['desc'] or '(sin descripción)'}\n"
            f"Fin validez:  {mat['fecha']}\n"
            f"Fila Hoja 1:  {mat['fila_h1']}\n\n"
            "Este material existe en el Purchasing pero no está en la tarifa activa.\n"
            "Si lo marca para baja, se escribirá 'X' en la columna AJ (Hoja 1).",
            parent=dlg,
        )

    tree.bind("<ButtonRelease-1>", _click_info)

    # ── Botones [MANTENER] / [DAR BAJA] / [CANCELAR] ─────────────────────────
    frm_pie = tk.Frame(dlg, bg=C_BG, padx=16, pady=12)
    frm_pie.pack(side="bottom", fill="x")

    tk.Frame(frm_pie, bg="#D0D8E8", height=1).pack(fill="x", pady=(0, 10))

    # Texto informativo sobre la selección
    lbl_info = tk.Label(
        frm_pie,
        text="Marque los ítems a dar de baja y pulse DAR BAJA, "
             "o pulse MANTENER para conservar todos sin cambios.",
        bg=C_BG, fg=C_TEXT, font=FONT_SMALL,
        wraplength=600, justify="left",
    )
    lbl_info.pack(anchor="w", pady=(0, 8))

    frm_botones = tk.Frame(frm_pie, bg=C_BG)
    frm_botones.pack(fill="x")

    def _accion_mantener():
        """Conservar todos sin cambios (ningún check = baja)."""
        for i in range(n):
            checks[i] = False
        for d in pendientes:
            resultado.append({
                "codigo":  d["codigo"],
                "fila_h1": d["fila_h1"],
                "baja":    False,
            })
        dlg.destroy()

    def _accion_dar_baja():
        """Aplicar la baja solo a los ítems marcados con checkbox."""
        for i, d in enumerate(pendientes):
            resultado.append({
                "codigo":  d["codigo"],
                "fila_h1": d["fila_h1"],
                "baja":    checks[i],
            })
        dlg.destroy()

    def _accion_cancelar():
        """Cancelar: ningún ítem recibe baja (comportamiento conservador)."""
        for d in pendientes:
            resultado.append({
                "codigo":  d["codigo"],
                "fila_h1": d["fila_h1"],
                "baja":    False,
            })
        dlg.destroy()

    # Botón MANTENER (verde neutro)
    tk.Button(
        frm_botones,
        text="✔  MANTENER",
        command=_accion_mantener,
        bg="#1A7A40", fg="white",
        font=FONT_BOLD, padx=14, pady=9,
        cursor="hand2", relief="flat",
        activebackground="#228B4A",
    ).pack(side="left", padx=(0, 6))

    # Botón DAR BAJA (rojo)
    tk.Button(
        frm_botones,
        text="✖  DAR BAJA",
        command=_accion_dar_baja,
        bg="#C0392B", fg="white",
        font=FONT_BOLD, padx=14, pady=9,
        cursor="hand2", relief="flat",
        activebackground="#A93226",
    ).pack(side="left", padx=(0, 6))

    # Botón CANCELAR (gris)
    tk.Button(
        frm_botones,
        text="Cancelar",
        command=_accion_cancelar,
        font=FONT_LABEL, padx=14, pady=7,
        cursor="hand2", relief="flat",
        bg="#D8DEE8",
        activebackground="#C0C8D8",
    ).pack(side="left")

    # ── Ajustar geometría ─────────────────────────────────────────────────────
    dlg.update_idletasks()
    sw = dlg.winfo_screenwidth()
    sh = dlg.winfo_screenheight()
    w  = min(700, sw - 60)
    h  = max(380, min(dlg.winfo_reqheight() + 60, sh - 80))
    x  = (sw - w) // 2
    y  = max(20, (sh - h) // 2)
    dlg.geometry(f"{w}x{h}+{x}+{y}")
    dlg.minsize(560, 340)

    # ── Esperar a que el usuario cierre la ventana ────────────────────────────
    # wait_window() bloquea hasta que dlg.destroy() sea llamado.
    parent.wait_window(dlg)

    # Si el usuario cerró la ventana con la X (sin pulsar botón), resultado=[]
    if not resultado:
        for d in pendientes:
            resultado.append({
                "codigo":  d["codigo"],
                "fila_h1": d["fila_h1"],
                "baja":    False,
            })

    return resultado

class VentanaMaterialesParaReactivar(tk.Toplevel):
    """
    Aviso final (2026-09-30, caso real MELOCOTON/FRAN Y CHEMI): se muestra
    cuando la tarifa incluye materiales cuyo Purchasing traía el registro
    info marcado para borrado (ver _avisar_materiales_reactivados_bloqueados).

    Confirmado en real, dos veces, que ninguna carga por Excel puede
    reactivarlos (SAP compara por proveedor+material+organización de
    compras, no por el número de la fila, y rechaza tanto conservar el
    número real como renumerar como registro nuevo). La única vía que
    funcionó fue la transacción SAP "Marcar reg.info de compra para
    borrar", desmarcando sus dos casillas de borrado -- esta ventana
    muestra esa pantalla como referencia visual, junto con la lista de
    materiales/registros concretos que hay que reactivar así.
    """

    def __init__(self, parent, materiales: dict, descripciones: dict = None,
                 centros: dict = None):
        """
        materiales: {codigo_sap: registro_info_antiguo} -- ver
        self._materiales_registro_a_renumerar en GeneradorDesdeExistente.
        descripciones: {codigo_sap: texto_breve} opcional, para mostrar el
        nombre del material junto a su código.
        centros: {codigo_sap: [centro1, centro2, ...]} opcional -- un mismo
        registro/material puede tener el problema en unos centros y no en
        otros, así que se muestran explícitamente (ver
        self._materiales_centros_borrado en GeneradorDesdeExistente).
        """
        super().__init__(parent)
        self.title("Materiales a reactivar manualmente en SAP")
        self.configure(bg=C_BG)
        self.resizable(True, True)
        cargar_icono_ventana(self)

        self._materiales    = materiales
        self._descripciones = descripciones or {}
        self._centros       = centros or {}

        self._construir_ui()
        self._ajustar_geometria()
        self.grab_set()

    def _ajustar_geometria(self):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w  = min(920, sw - 40)
        h  = min(max(760, self.winfo_reqheight() + 40), sh - 60)
        x  = (sw - w) // 2
        y  = max(20, (sh - h) // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.minsize(800, 680)

    def _construir_ui(self):
        hdr = tk.Frame(self, bg="#B8860B", pady=12, padx=16)
        hdr.pack(fill="x")
        tk.Label(hdr, text="⚠  Materiales a reactivar manualmente en SAP",
                 bg="#B8860B", fg="white", font=FONT_TITLE).pack(anchor="w")
        tk.Label(hdr,
                 text="Su registro info seguía marcado para borrado en el "
                      "Purchasing -- este RINFO NO puede reactivarlos por sí solo.",
                 bg="#B8860B", fg="#FFF3D0", font=FONT_SUB,
                 wraplength=850, justify="left").pack(anchor="w", pady=(4, 0))

        frm_pie = tk.Frame(self, bg=C_BG, padx=22, pady=12)
        frm_pie.pack(side="bottom", fill="x")
        tk.Frame(frm_pie, bg="#D0D8E8", height=1).pack(fill="x", pady=(0, 10))
        tk.Button(frm_pie, text="Entendido, cerrar",
                  command=self._cerrar,
                  bg="#B8860B", fg="white",
                  font=FONT_BOLD, padx=14, pady=11,
                  cursor="hand2", relief="flat",
                  activebackground="#9A7009").pack(fill="x")

        # ── Zona con scroll (lista + instrucciones + imagen) ──────────────
        contenedor = tk.Frame(self, bg=C_BG)
        contenedor.pack(fill="both", expand=True, padx=22, pady=(14, 4))
        canvas  = tk.Canvas(contenedor, bg=C_BG, highlightthickness=0)
        scrolly = ttk.Scrollbar(contenedor, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrolly.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrolly.pack(side="right", fill="y")
        cuerpo = tk.Frame(canvas, bg=C_BG)
        canvas.create_window((0, 0), window=cuerpo, anchor="nw")
        cuerpo.bind("<Configure>",
                    lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        tk.Label(cuerpo,
                 text=f"Se han detectado {len(self._materiales)} material(es) "
                      f"con tarifa activa cuyo registro info seguía marcado "
                      f"para borrado en el Purchasing:",
                 bg=C_BG, fg=C_TEXT, font=("Segoe UI", 9, "bold"),
                 wraplength=850, justify="left").pack(anchor="w", pady=(0, 8))

        for cod_sap, registro_viejo in self._materiales.items():
            desc = self._descripciones.get(cod_sap, "")
            centros_mat = self._centros.get(cod_sap, [])
            centros_txt = (f"   centro(s): {', '.join(centros_mat)}"
                           if centros_mat else "")
            texto = f"  📦  {cod_sap}" + (f" — {desc}" if desc else "") \
                    + f"   (registro {registro_viejo})" + centros_txt
            tk.Label(cuerpo, text=texto,
                     bg="#FFF8E1", fg="#7B5800",
                     font=FONT_LABEL, padx=10, pady=6, relief="flat",
                     wraplength=850, justify="left", anchor="w").pack(
                fill="x", pady=(0, 4))

        tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=(12, 10))

        tk.Label(cuerpo, text="Cómo reactivarlos en SAP:",
                 bg=C_BG, fg=C_TEXT, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        pasos = (
            "1.  En SAP, busca la transacción \"Marcar reg.info de compra "
            "para borrar\".\n"
            "2.  Introduce el Registro info, Proveedor, Material y Centro "
            "del material afectado (uno por cada centro, si tiene más de "
            "uno marcado).\n"
            "3.  Desmarca las dos casillas de \"Peticiones de borrado\": "
            "\"Regtro-Info completo\" e \"Ind.borr.dat.org.co.\".\n"
            "4.  Ejecuta el cambio y confirma en \"Gestión de registros "
            "info de compras\" que el material ya aparece para ese centro."
        )
        tk.Label(cuerpo, text=pasos,
                 bg="#F0F5FB", fg="#285FA8",
                 font=FONT_LABEL, padx=10, pady=8, relief="flat",
                 wraplength=850, justify="left").pack(fill="x", pady=(4, 12))

        tk.Label(cuerpo, text="Referencia — pantalla de acceso de esa transacción:",
                 bg=C_BG, fg=C_TEXT, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        self._img_ref = self._cargar_imagen_referencia()
        if self._img_ref is not None:
            tk.Label(cuerpo, image=self._img_ref, bg=C_BG,
                     relief="solid", borderwidth=1).pack(pady=(6, 4))

    def _cargar_imagen_referencia(self):
        try:
            from PIL import Image, ImageTk
            img = Image.open(resource_path("assets/marcar_borrado_sap.png"))   # (nov 2026) antes, base64 incrustado aquí
            ancho_destino = 760
            alto_destino  = int(img.height * (ancho_destino / img.width))
            img = img.resize((ancho_destino, alto_destino), Image.LANCZOS)
            return ImageTk.PhotoImage(img)
        except Exception:
            return None

    def _cerrar(self):
        parent = self.master
        self.destroy()
        try:
            if parent and parent.winfo_exists():
                parent.deiconify()
        except Exception:
            pass

class VentanaConfirmacionGenerado(tk.Toplevel):
    """
    Confirmación final tras generar el nuevo archivo desde un Purchasing existente.
    Informa del nº de materiales ya existentes, los nuevos insertados y ofrece
    abrir el archivo generado.
    """

    def __init__(self, parent, nombre_nuevo: str, ruta_nuevo: str,
                 nombre_sap: str, codigo_sap: str,
                 n_existentes: int, n_insertados: int,
                 nombre_purchasing_base: str):
        super().__init__(parent)
        self.title("Archivo generado correctamente")
        self.configure(bg=C_BG)
        self.resizable(True, True)
        self._ruta_nuevo = ruta_nuevo
        cargar_icono_ventana(self)
        self._construir_ui(nombre_nuevo, ruta_nuevo, nombre_sap, codigo_sap,
                            n_existentes, n_insertados, nombre_purchasing_base)
        self._ajustar_geometria()
        self.grab_set()

    def _ajustar_geometria(self):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w  = min(680, sw - 40)
        h  = max(540, min(self.winfo_reqheight() + 40, sh - 60))
        x  = (sw - w) // 2
        y  = max(20, (sh - h) // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.minsize(580, 540)

    def _construir_ui(self, nombre_nuevo, ruta_nuevo, nombre_sap, codigo_sap,
                       n_existentes, n_insertados, nombre_purchasing_base):
        carpeta = os.path.dirname(ruta_nuevo)

        hdr = tk.Frame(self, bg="#1A7A40", pady=12, padx=16)
        hdr.pack(fill="x")
        tk.Label(hdr, text="✔  Archivo generado correctamente",
                 bg="#1A7A40", fg="white", font=FONT_TITLE).pack(anchor="w")
        tk.Label(hdr,
                 text="Los materiales nuevos han sido añadidos al Purchasing base.",
                 bg="#1A7A40", fg="#A8E6C0", font=FONT_SUB).pack(anchor="w", pady=(4, 0))

        frm_pie = tk.Frame(self, bg=C_BG, padx=22, pady=12)
        frm_pie.pack(side="bottom", fill="x")
        tk.Frame(frm_pie, bg="#D0D8E8", height=1).pack(fill="x", pady=(0, 10))
        tk.Label(frm_pie, text="¿Desea abrir el archivo ahora?",
                 bg=C_BG, fg=C_TEXT,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 8))
        tk.Button(frm_pie, text="✔  Sí, abrir ahora",
                  command=self._abrir_y_cerrar,
                  bg="#1A7A40", fg="white",
                  font=FONT_BOLD, padx=14, pady=11,
                  cursor="hand2", relief="flat",
                  activebackground="#228B4A").pack(fill="x", pady=(0, 6))
        tk.Button(frm_pie, text="No, cerrar",
                  command=self._cerrar,
                  font=FONT_LABEL, padx=14, pady=8,
                  cursor="hand2", relief="flat",
                  bg="#D8DEE8",
                  activebackground="#C0C8D8").pack(fill="x")

        cuerpo = tk.Frame(self, bg=C_BG, padx=22, pady=16)
        cuerpo.pack(fill="both", expand=True)

        tk.Label(cuerpo, text="✔  Archivo generado:",
                 bg=C_BG, fg=C_TEXT, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(cuerpo, text=f"  📄  {nombre_nuevo}",
                 bg="#EEF7EE", fg="#1B5E20",
                 font=("Segoe UI", 10, "bold"),
                 padx=10, pady=8, relief="flat").pack(fill="x", pady=(4, 10))

        tk.Label(cuerpo, text="Guardado en:",
                 bg=C_BG, fg=C_TEXT, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(cuerpo, text=f"  📁  {carpeta}",
                 bg="#F5F5F5", fg="#333",
                 font=FONT_LABEL, padx=10, pady=6, relief="flat",
                 wraplength=600, justify="left").pack(fill="x", pady=(4, 10))

        tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=(0, 10))

        tk.Label(cuerpo, text="Archivo base (Purchasing):",
                 bg=C_BG, fg=C_TEXT, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(cuerpo, text=f"  📋  {nombre_purchasing_base}",
                 bg="#F0F5FB", fg="#285FA8",
                 font=FONT_LABEL, padx=10, pady=6, relief="flat",
                 wraplength=600, justify="left").pack(fill="x", pady=(4, 10))

        tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=(0, 10))

        tk.Label(cuerpo, text="Resumen de materiales:",
                 bg=C_BG, fg=C_TEXT, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        frm_vals = tk.Frame(cuerpo, bg="#F0F5FB", padx=12, pady=8)
        frm_vals.pack(fill="x", pady=(4, 10))
        _fila_valor(frm_vals, "Ya existentes en Purchasing:", str(n_existentes))
        _fila_valor(frm_vals, "Nuevos insertados (tarifa):",  str(n_insertados))
        _fila_valor(frm_vals, "D/C — Proveedor SAP:",
                    f"{nombre_sap}  [{codigo_sap}]" if codigo_sap else nombre_sap)

        if n_insertados == 0:
            tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=(0, 8))
            tk.Label(cuerpo,
                     text="ℹ  Todos los materiales de la tarifa ya estaban en el Purchasing.",
                     bg="#FFF8E1", fg="#7B5800",
                     font=("Segoe UI", 9, "italic"),
                     padx=10, pady=6, relief="flat").pack(fill="x")

    def _abrir_y_cerrar(self):
        try:
            os.startfile(self._ruta_nuevo)
        except AttributeError:
            subprocess.call(["xdg-open", self._ruta_nuevo])
        except Exception:
            pass
        self._cerrar()

    def _cerrar(self):
        # Restaurar la ventana padre antes de destruir esta ventana.
        # Sin deiconify, el padre queda en withdraw() y la app parece cerrarse.
        parent = self.master
        self.destroy()
        try:
            if parent and parent.winfo_exists():
                parent.deiconify()
        except Exception:
            pass

# ─────────────────────────────────────────────────────────────────────────────
#  PASO 4a — VENTANA DE INSTRUCCIONES SAP (7 pasos)
# ─────────────────────────────────────────────────────────────────────────────

class VentanaInstruccionesSAP(tk.Toplevel):

    PASOS = [
        ("1", "Acceder a la aplicación TARIFAS en SAP",
         "En el menú de favoritos de SAP S/4HANA, localice el grupo  TARIFAS\n"
         "y pulse en:  «Modificaciones en masa de registros info de compras»."),
        ("2", "Pantalla de búsqueda",
         "Se abrirá la pantalla  «Estándar COMPRAS».\n"
         "Localice el campo  Proveedor:  en la parte superior izquierda."),
        ("3", "Abrir el selector de Proveedor",
         "Pulse el icono de cuadraditos (⊞) a la derecha del campo  Proveedor:\n"
         "para abrir el buscador de proveedores."),
        ("4", "Buscar y seleccionar el proveedor",
         ""),   # Se rellena dinámicamente
        ("5", "Ejecutar la búsqueda",
         "De vuelta en la pantalla principal, pulse el botón  Ir  (azul)\n"
         "para cargar todos los Registros Info del proveedor."),
        ("6", "Seleccionar todos y descargar",
         "Una vez cargados los registros:\n"
         "  1. Marque la casilla de cabecera para  SELECCIONAR TODOS.\n"
         "  2. Pulse el botón  Descargar  (parte superior derecha)."),
        ("7", "Guardar y volver a ejecutar",
         "Se descargará el archivo:  📄  Purchasing Info Records.xlsx\n"
         "Guárdelo en la carpeta  TARIFAS  del proyecto y vuelva\n"
         "a pulsar el botón  «CREAR REGISTRO INFO»  para continuar."),
    ]

    def __init__(self, parent, nombre_proveedor: str):
        super().__init__(parent)
        self.title("Cómo descargar el Purchasing Info Records desde SAP")
        self.resizable(True, True)
        self.configure(bg=C_BG)
        self._nombre_proveedor = nombre_proveedor
        cargar_icono_ventana(self)
        self._construir_ui()
        self._ajustar_geometria()
        self.grab_set()

    def _ajustar_geometria(self):
        self.update_idletasks()
        sw  = self.winfo_screenwidth()
        sh  = self.winfo_screenheight()
        h_disponible = sh - 48 - 30
        w = min(880, sw - 40)
        h = min(int(h_disponible * 0.88), h_disponible)
        x = (sw - w) // 2
        y = 20
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.minsize(min(w, 700), min(h, 540))

    def _construir_ui(self):
        hdr = tk.Frame(self, bg=C_MID, pady=12, padx=16)
        hdr.pack(fill="x")
        tk.Label(hdr, text="📥  Descargar Purchasing Info Records desde SAP",
                 bg=C_MID, fg="white", font=FONT_TITLE).pack(anchor="w")
        tk.Label(hdr,
                 text="Siga los pasos indicados a continuación en SAP S/4HANA Cloud:",
                 bg=C_MID, fg="#BDD5F0", font=FONT_SUB).pack(anchor="w", pady=(4, 0))

        frm_link = tk.Frame(self, bg="#EBF4FF", padx=16, pady=10)
        frm_link.pack(fill="x")
        tk.Label(frm_link,
                 text="🔗  Acceso directo a SAP S/4HANA — Modificaciones en masa de registros info:",
                 bg="#EBF4FF", fg="#1A3A6A",
                 font=("Segoe UI", 9, "bold")).pack(anchor="w")
        lnk = tk.Label(frm_link, text=URL_SAP,
                        bg="#EBF4FF", fg="#1565C0",
                        font=("Segoe UI", 9, "underline"),
                        cursor="hand2")
        lnk.pack(anchor="w", pady=(2, 0))
        lnk.bind("<Button-1>", lambda e: webbrowser.open(URL_SAP))

        frame_steps = tk.Frame(self, bg=C_BG)
        frame_steps.pack(fill="both", expand=True, padx=12, pady=(6, 0))

        pasos_renderizados = list(self.PASOS)
        pasos_renderizados[3] = (
            "4", "Buscar y seleccionar el proveedor",
            f"En el buscador que se abre:\n"
            f"  1. Escriba el nombre:  «{self._nombre_proveedor}»\n"
            f"  2. Pulse la lupa (🔍) o el botón  Ir.\n"
            f"  3. Marque la casilla junto al Acreedor (✔) para seleccionar todos.\n"
            f"  4. Pulse  OK.",
        )

        for idx, (num, titulo, desc) in enumerate(pasos_renderizados):
            frame_steps.rowconfigure(idx, weight=1)
            frame_steps.columnconfigure(0, weight=1)

            frm_fila = tk.Frame(frame_steps, bg=C_BG)
            frm_fila.grid(row=idx, column=0, sticky="nsew", padx=4, pady=1)

            tk.Label(frm_fila,
                     text=f" {num} ",
                     bg=C_MID, fg="white",
                     font=("Segoe UI", 10, "bold"),
                     padx=8, pady=4,
                     relief="flat").pack(side="left", anchor="n", padx=(0, 8))

            frm_txt = tk.Frame(frm_fila, bg="#EEF2F8", padx=10, pady=4)
            frm_txt.pack(side="left", fill="both", expand=True)
            tk.Label(frm_txt, text=titulo,
                     bg="#EEF2F8", fg="#1A2E4A",
                     font=("Segoe UI", 9, "bold"),
                     anchor="w", justify="left").pack(anchor="w")
            tk.Label(frm_txt, text=desc,
                     bg="#EEF2F8", fg=C_TEXT,
                     font=FONT_LABEL,
                     anchor="w", justify="left").pack(anchor="w", pady=(1, 0))

        frm_cerrar = tk.Frame(self, bg=C_BG, pady=6, padx=14)
        frm_cerrar.pack(fill="x")
        tk.Button(frm_cerrar, text="Cerrar",
                  command=self.destroy,
                  font=FONT_LABEL, padx=14, pady=5,
                  cursor="hand2", relief="flat",
                  bg="#D8DEE8",
                  activebackground="#C0C8D8").pack(side="right")
        self.grab_set()

# ─────────────────────────────────────────────────────────────────────────────
#  PASO 4b.0 — VENTANA: ASIGNACIÓN OPCIONAL DE TELÉFONO DE PROVEEDOR
# ─────────────────────────────────────────────────────────────────────────────

class VentanaTelefonoRINFO(tk.Toplevel):
    """
    Pregunta al usuario si desea asignar un número de teléfono al Registro Info
    antes de iniciar la creación.

    · Si confirma con un número válido (8-10 dígitos, espacios opcionales) → llama
      al callback con el número limpio como str (sin espacios).
    · Si deja vacío o pulsa «Sin teléfono / Continuar» → llama al callback con None.
    · Si cierra la ventana o cancela → llama al callback con False (abortar).

    Los espacios que el usuario introduzca se eliminan automáticamente antes de
    validar y de escribir el valor en la columna AD del RINFO.
    """

    _RE_TELEFONO = re.compile(r'^\d{8,10}$')

    def __init__(self, parent, on_confirm):
        """
        Parámetros
        ----------
        parent      : ventana padre (Toplevel o Tk).
        on_confirm  : callable(resultado) donde resultado es:
                        str   → número validado sin espacios (p.ej. «612345678»)
                        None  → sin teléfono, continuar igualmente
                        False → cancelado / cerrar sin crear RINFO
        """
        super().__init__(parent)
        self.title("Número de teléfono del proveedor")
        self.configure(bg=C_BG)
        self.resizable(False, False)
        self._on_confirm = on_confirm

        cargar_icono_ventana(self)

        self._construir_ui()
        ajustar_geometria_pantalla(self, 480)
        self.protocol("WM_DELETE_WINDOW", self._cancelar)
        self.grab_set()

    # ── Geometría ─────────────────────────────────────────────────────────────

    # ── UI ────────────────────────────────────────────────────────────────────
    def _construir_ui(self):
        # Cabecera
        hdr = tk.Frame(self, bg=C_MID, pady=10, padx=16)
        hdr.pack(fill="x")
        tk.Label(hdr, text="📞  Teléfono del proveedor (opcional)",
                 bg=C_MID, fg="white", font=FONT_TITLE).pack(anchor="w")
        tk.Label(hdr,
                 text="Este dato se escribirá en la columna AD del Registro Info.",
                 bg=C_MID, fg="#BDD5F0", font=FONT_SUB).pack(anchor="w", pady=(3, 0))

        # Cuerpo
        cuerpo = tk.Frame(self, bg=C_BG, padx=22, pady=16)
        cuerpo.pack(fill="both", expand=True)

        tk.Label(cuerpo,
                 text="¿Desea asignar un número de teléfono al Registro Info?",
                 bg=C_BG, fg=C_TEXT,
                 font=("Segoe UI", 10, "bold"),
                 wraplength=420, justify="left").pack(anchor="w", pady=(0, 4))

        tk.Label(cuerpo,
                 text="Formato: 000000000  (8-10 dígitos; los espacios se eliminan automáticamente)\n"
                      "Si deja vacío o pulsa «Sin teléfono», continuaré sin asignar teléfono.",
                 bg=C_BG, fg="#666",
                 font=FONT_LABEL,
                 justify="left").pack(anchor="w", pady=(0, 10))

        # Campo de entrada
        frm_entrada = tk.Frame(cuerpo, bg=C_BG)
        frm_entrada.pack(fill="x", pady=(0, 4))

        tk.Label(frm_entrada, text="Teléfono:",
                 bg=C_BG, fg=C_TEXT,
                 font=FONT_BOLD, width=10, anchor="w").pack(side="left")

        self._var_tel = tk.StringVar()
        self._entry = tk.Entry(frm_entrada,
                               textvariable=self._var_tel,
                               font=("Segoe UI", 11),
                               width=16, relief="solid", bd=1)
        self._entry.pack(side="left", padx=(4, 0))
        self._entry.bind("<Return>", lambda e: self._confirmar())

        # Etiqueta de error (oculta inicialmente)
        self._lbl_error = tk.Label(cuerpo, text="",
                                   bg=C_BG, fg=C_ERROR,
                                   font=("Segoe UI", 8, "italic"))
        self._lbl_error.pack(anchor="w", pady=(0, 8))

        tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=(0, 12))

        # Botones
        tk.Button(cuerpo,
                  text="✔  Asignar teléfono y continuar",
                  command=self._confirmar,
                  bg="#1A7A40", fg="white",
                  font=FONT_BOLD, padx=12, pady=10,
                  cursor="hand2", relief="flat",
                  activebackground="#228B4A").pack(fill="x", pady=(0, 6))

        tk.Button(cuerpo,
                  text="Sin teléfono — continuar igualmente",
                  command=self._sin_telefono,
                  bg=C_MID, fg=C_TEXT_LIGHT,
                  font=FONT_BOLD, padx=12, pady=9,
                  cursor="hand2", relief="flat",
                  activebackground=C_HOVER,
                  activeforeground=C_ACCENT2).pack(fill="x", pady=(0, 6))

        tk.Button(cuerpo, text="Cancelar",
                  command=self._cancelar,
                  font=FONT_LABEL, padx=12, pady=7,
                  cursor="hand2", relief="flat",
                  bg="#D8DEE8", activebackground="#C0C8D8").pack(fill="x")

        self._entry.focus_set()

    # ── Acciones ──────────────────────────────────────────────────────────────
    def _confirmar(self):
        valor = self._var_tel.get().strip()
        if not valor:
            # Vacío → sin teléfono
            self._sin_telefono()
            return
        # Eliminar espacios antes de validar (acepta "000 000 000" o "000000000")
        valor_limpio = valor.replace(" ", "")
        if not self._RE_TELEFONO.match(valor_limpio):
            self._lbl_error.config(
                text="⚠  Formato incorrecto. Se esperan entre 8 y 10 dígitos (los espacios se ignoran)")
            self._entry.focus_set()
            return
        self._lbl_error.config(text="")
        self.destroy()
        self._on_confirm(valor_limpio)

    def _sin_telefono(self):
        self.destroy()
        self._on_confirm(None)

    def _cancelar(self):
        self.destroy()
        self._on_confirm(False)

# ─────────────────────────────────────────────────────────────────────────────
#  PASO 4b.1 — VENTANA: FECHAS DE VALIDEZ DEL REGISTRO INFO
# ─────────────────────────────────────────────────────────────────────────────

class VentanaFechasRINFO(tk.Toplevel):
    """
    Solicita al usuario las fechas de inicio y fin de validez de la tarifa
    (Registro Info) antes de iniciar la creación del RINFO.

    Ambas fechas son OBLIGATORIAS.

    Formatos aceptados (entrada):
        dd-mm-yyyy  |  dd.mm.yyyy  |  dd,mm,yyyy  |  ddmmyyyy
    El valor que se escribe en el archivo es siempre "dd.mm.yyyy".

    Callbacks:
        on_confirm((fecha_inicio_str, fecha_fin_str))  → fechas validadas
        on_confirm(False)                              → cancelado / cerrado
    """

    # Variantes separadoras aceptadas: guion, punto, coma, o sin separador
    _RE_FECHA = re.compile(
        r'^(\d{2})[-.,]?(\d{2})[-.,]?(\d{4})$'
    )

    def __init__(self, parent, on_confirm, periodo=""):
        """
        Parámetros
        ----------
        parent      : ventana padre (Toplevel o Tk).
        periodo     : (13 nov 2026) texto «Periodo:» de la tarifa (C3); de él salen las fechas propuestas y los
                      avisos al confirmar (app/comun/fechas_tarifa.py).
        on_confirm  : callable(resultado) donde resultado es:
                        (str, str)  → (fecha_inicio, fecha_fin) en formato dd.mm.yyyy
                        False       → cancelado / cerrar sin crear RINFO
        """
        super().__init__(parent)
        self.title("Fechas de validez del Registro Info")
        self.configure(bg=C_BG)
        self.resizable(False, False)
        self._on_confirm = on_confirm
        self._periodo = (periodo or "").strip()

        cargar_icono_ventana(self)

        self._construir_ui()
        ajustar_geometria_pantalla(self, 500)
        self.protocol("WM_DELETE_WINDOW", self._cancelar)
        self.grab_set()

    # ── Geometría ─────────────────────────────────────────────────────────────

    # ── UI ────────────────────────────────────────────────────────────────────
    def _construir_ui(self):
        # Cabecera
        hdr = tk.Frame(self, bg="#285FA8", pady=10, padx=16)
        hdr.pack(fill="x")
        tk.Label(hdr, text="📅  Fechas de validez de la tarifa (Reg. Info)",
                 bg="#285FA8", fg="white", font=FONT_TITLE).pack(anchor="w")
        tk.Label(hdr,
                 text="Estos valores se escribirán en la hoja 3 del Registro Info.",
                 bg="#285FA8", fg="#BDD5F0", font=FONT_SUB).pack(anchor="w", pady=(3, 0))

        # Cuerpo
        cuerpo = tk.Frame(self, bg=C_BG, padx=22, pady=16)
        cuerpo.pack(fill="both", expand=True)

        tk.Label(cuerpo,
                 text="Formato requerido: [dd.mm.yyyy]",
                 bg=C_BG, fg="#666",
                 font=FONT_LABEL,
                 wraplength=440, justify="left").pack(anchor="w", pady=(0, 4))
        # (13 nov 2026) De dónde salen las fechas propuestas
        self._lbl_origen = tk.Label(cuerpo, text="", bg=C_BG, fg="#285FA8", font=("Segoe UI", 9, "bold"),
                                    wraplength=440, justify="left")
        self._lbl_origen.pack(anchor="w", pady=(0, 10))

        # ── Campo Fecha Inicio ─────────────────────────────────────────────────
        tk.Label(cuerpo,
                 text="Fecha Inicio Tarifa (Inicio de validez):",
                 bg=C_BG, fg=C_TEXT,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w")

        frm_inicio = tk.Frame(cuerpo, bg=C_BG)
        frm_inicio.pack(fill="x", pady=(4, 0))

        tk.Label(frm_inicio, text="Inicio:",
                 bg=C_BG, fg=C_TEXT,
                 font=FONT_BOLD, width=8, anchor="w").pack(side="left")

        self._var_inicio = tk.StringVar()
        self._entry_inicio = tk.Entry(frm_inicio,
                                      textvariable=self._var_inicio,
                                      font=("Segoe UI", 11),
                                      width=14, relief="solid", bd=1)
        self._entry_inicio.pack(side="left", padx=(4, 0))
        self._entry_inicio.bind("<Return>", lambda e: self._entry_fin.focus_set())
        # (13 nov 2026, petición de Víctor) Calendario para no escribir la fecha; el formato no cambia
        self._btn_cal_inicio = boton_calendario(frm_inicio, self._var_inicio, bg=C_BG, font=("Segoe UI", 12))
        self._btn_cal_inicio.pack(side="left", padx=(6, 0))

        tk.Label(frm_inicio, text="  p.ej. 01-01-2025 (o 📅)",
                 bg=C_BG, fg="#999", font=FONT_SMALL).pack(side="left", padx=(8, 0))

        self._lbl_error_inicio = tk.Label(cuerpo, text="",
                                          bg=C_BG, fg=C_ERROR,
                                          font=("Segoe UI", 8, "italic"))
        self._lbl_error_inicio.pack(anchor="w", pady=(2, 8))

        tk.Frame(cuerpo, bg="#E0E8F0", height=1).pack(fill="x", pady=(0, 10))

        # ── Campo Fecha Fin ────────────────────────────────────────────────────
        tk.Label(cuerpo,
                 text="Fecha Finalización Tarifa (Fin de validez):",
                 bg=C_BG, fg=C_TEXT,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w")

        frm_fin = tk.Frame(cuerpo, bg=C_BG)
        frm_fin.pack(fill="x", pady=(4, 0))

        tk.Label(frm_fin, text="Fin:",
                 bg=C_BG, fg=C_TEXT,
                 font=FONT_BOLD, width=8, anchor="w").pack(side="left")

        self._var_fin = tk.StringVar()
        self._entry_fin = tk.Entry(frm_fin,
                                   textvariable=self._var_fin,
                                   font=("Segoe UI", 11),
                                   width=14, relief="solid", bd=1)
        self._entry_fin.pack(side="left", padx=(4, 0))
        self._entry_fin.bind("<Return>", lambda e: self._confirmar())
        self._btn_cal_fin = boton_calendario(frm_fin, self._var_fin, bg=C_BG, font=("Segoe UI", 12))
        self._btn_cal_fin.pack(side="left", padx=(6, 0))

        tk.Label(frm_fin, text="  p.ej. 31-12-2025 (o 📅)",
                 bg=C_BG, fg="#999", font=FONT_SMALL).pack(side="left", padx=(8, 0))

        self._lbl_error_fin = tk.Label(cuerpo, text="",
                                       bg=C_BG, fg=C_ERROR,
                                       font=("Segoe UI", 8, "italic"))
        self._lbl_error_fin.pack(anchor="w", pady=(2, 8))

        tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=(0, 12))

        # ── Botones ────────────────────────────────────────────────────────────
        tk.Button(cuerpo,
                  text="✔  Confirmar fechas y continuar",
                  command=self._confirmar,
                  bg="#1A7A40", fg="white",
                  font=FONT_BOLD, padx=12, pady=10,
                  cursor="hand2", relief="flat",
                  activebackground="#228B4A").pack(fill="x", pady=(0, 6))

        tk.Button(cuerpo, text="Cancelar",
                  command=self._cancelar,
                  font=FONT_LABEL, padx=12, pady=7,
                  cursor="hand2", relief="flat",
                  bg="#D8DEE8", activebackground="#C0C8D8").pack(fill="x")

        # (13 nov 2026) Fechas propuestas: las del periodo de la tarifa; si no, las últimas usadas en este PC; si no,
        # el año en curso (lo de antes, que llevó a un RINFO de enero a diciembre para una tarifa del 4º trimestre)
        _ini, _fin, _origen = FT.propuesta(self._periodo, ultimas=FT.leer_ultimas())
        self._var_inicio.set(_ini.strftime("%d.%m.%Y"))
        self._var_fin.set(_fin.strftime("%d.%m.%Y"))
        self._lbl_origen.config(text={
            "periodo": f"Propuestas según el periodo de la tarifa: «{self._periodo}».",
            "ultimas": (f"La tarifa no indica un periodo reconocible («{self._periodo}»): se proponen las últimas "
                        "fechas usadas. Revísalas." if self._periodo else
                        "La tarifa no indica periodo: se proponen las últimas fechas usadas. Revísalas."),
            "anio": (f"No se reconoce el periodo de la tarifa («{self._periodo}»): se propone el año en curso. "
                     "Revísalas." if self._periodo else
                     "La tarifa no indica periodo: se propone el año en curso. Revísalas."),
        }[_origen], fg="#285FA8" if _origen == "periodo" else "#A15C00")

        self._entry_inicio.focus_set()

    # ── Validación ────────────────────────────────────────────────────────────
    @classmethod
    def _parsear_fecha(cls, texto: str):
        """
        Intenta parsear el texto introducido.
        Devuelve "dd.mm.yyyy" si es válido, None si no.
        Acepta: dd-mm-yyyy | dd.mm.yyyy | dd,mm,yyyy | ddmmyyyy
        """
        texto = texto.strip()
        # Quitar separadores mixtos (guion, punto, coma, espacio)
        m = cls._RE_FECHA.match(texto)
        if not m:
            return None
        dia, mes, anio = m.group(1), m.group(2), m.group(3)
        # Validar rangos básicos
        try:
            datetime(int(anio), int(mes), int(dia))
        except ValueError:
            return None
        return f"{dia}.{mes}.{anio}"

    @classmethod
    def _a_date(cls, texto: str):
        """Devuelve un objeto date a partir de texto ya validado 'dd.mm.yyyy', o None."""
        try:
            d, m, y = texto.split(".")
            return datetime(int(y), int(m), int(d)).date()
        except Exception:
            return None

    # ── Acciones ──────────────────────────────────────────────────────────────
    def _confirmar(self):
        # Limpiar errores previos
        self._lbl_error_inicio.config(text="")
        self._lbl_error_fin.config(text="")

        fecha_inicio = self._parsear_fecha(self._var_inicio.get())
        fecha_fin    = self._parsear_fecha(self._var_fin.get())

        ok = True

        if not self._var_inicio.get().strip():
            self._lbl_error_inicio.config(
                text="⚠  Campo obligatorio. Introduzca la fecha de inicio.")
            self._entry_inicio.focus_set()
            ok = False
        elif fecha_inicio is None:
            self._lbl_error_inicio.config(
                text="⚠  Fecha inválida — compruebe día, mes y año (p.ej. 01-01-2025).")
            self._entry_inicio.focus_set()
            ok = False

        if not self._var_fin.get().strip():
            self._lbl_error_fin.config(
                text="⚠  Campo obligatorio. Introduzca la fecha de fin.")
            if ok:
                self._entry_fin.focus_set()
            ok = False
        elif fecha_fin is None:
            self._lbl_error_fin.config(
                text="⚠  Fecha inválida — compruebe día, mes y año (p.ej. 31-12-2025).")
            if ok:
                self._entry_fin.focus_set()
            ok = False

        # ── Validación cruzada: fin no puede ser anterior a inicio ────────────
        if ok and fecha_inicio and fecha_fin:
            d_inicio = self._a_date(fecha_inicio)
            d_fin    = self._a_date(fecha_fin)
            if d_inicio is not None and d_fin is not None and d_fin < d_inicio:
                self._lbl_error_fin.config(
                    text=f"⚠  La fecha de fin ({fecha_fin}) no puede ser "
                         f"anterior a la de inicio ({fecha_inicio}).")
                self._entry_fin.focus_set()
                ok = False

        if not ok:
            return

        # (13 nov 2026) Confirmar si el inicio es anterior a hoy o no coincide con el periodo de la tarifa
        d_inicio, d_fin = self._a_date(fecha_inicio), self._a_date(fecha_fin)
        lista = FT.avisos(d_inicio, d_fin, self._periodo)
        if lista and not messagebox.askyesno(
                "Revisar las fechas de validez",
                "\n\n".join(f"⚠  {t}" for t in lista)
                + "\n\n¿Continuar con estas fechas?\n(«No» para corregirlas)", parent=self):
            self._entry_inicio.focus_set()
            return
        FT.guardar_ultimas(d_inicio, d_fin)

        self.destroy()
        self._on_confirm((fecha_inicio, fecha_fin))

    def _cancelar(self):
        self.destroy()
        self._on_confirm(False)

# ─────────────────────────────────────────────────────────────────────────────
#  PASO 4b.2 — VENTANA: SELECCIÓN DE CENTROS SAP
# ─────────────────────────────────────────────────────────────────────────────

class VentanaCentrosRINFO(tk.Toplevel):
    """
    Muestra los 10 centros SAP de Canarias con casillas de verificación.
    Todos los centros aparecen marcados por defecto.
    El usuario desmarca los centros que NO desea incluir en el Registro Info.

    Callbacks:
        on_confirm(lista)   → lista de (codigo, nombre) de los centros seleccionados
                              (al menos 1 debe quedar marcado)
        on_confirm(False)   → cancelado / ventana cerrada
    """

    def __init__(self, parent, on_confirm):
        """
        Parámetros
        ----------
        parent      : ventana padre (Toplevel o Tk).
        on_confirm  : callable(resultado) donde resultado es:
                        list[(str, str)]  → centros seleccionados [(codigo, nombre), ...]
                        False             → cancelado / cerrar sin crear RINFO
        """
        super().__init__(parent)
        self.title("Selección de centros — Registro Info")
        self.configure(bg=C_BG)
        self.resizable(False, False)
        self._on_confirm = on_confirm

        cargar_icono_ventana(self)

        self._vars = {}   # codigo → BooleanVar
        self._construir_ui()
        ajustar_geometria_pantalla(self, 520)
        self.protocol("WM_DELETE_WINDOW", self._cancelar)
        self.grab_set()

    # ── Geometría ─────────────────────────────────────────────────────────────

    # ── UI ────────────────────────────────────────────────────────────────────
    def _construir_ui(self):
        # Cabecera
        hdr = tk.Frame(self, bg=C_MID, pady=10, padx=16)
        hdr.pack(fill="x")
        tk.Label(hdr, text="🏨  Centros SAP — Selección para el Registro Info",
                 bg=C_MID, fg="white", font=FONT_TITLE).pack(anchor="w")
        tk.Label(hdr,
                 text="Desmarque los centros que NO se van a incluir. "
                      "Por defecto se incluyen todos.",
                 bg=C_MID, fg="#BDD5F0", font=FONT_SUB,
                 wraplength=460, justify="left").pack(anchor="w", pady=(3, 0))

        # Cuerpo
        cuerpo = tk.Frame(self, bg=C_BG, padx=22, pady=14)
        cuerpo.pack(fill="both", expand=True)

        # Botones rápidos Todos / Ninguno
        frm_rapido = tk.Frame(cuerpo, bg=C_BG)
        frm_rapido.pack(fill="x", pady=(0, 8))
        tk.Label(frm_rapido, text="Selección rápida:",
                 bg=C_BG, fg="#888", font=FONT_SMALL).pack(side="left")
        tk.Button(frm_rapido, text="✔ Todos",
                  command=self._seleccionar_todos,
                  bg="#E8F0FE", fg="#285FA8",
                  font=FONT_SMALL, padx=8, pady=2,
                  cursor="hand2", relief="flat",
                  activebackground="#D0E0F8").pack(side="left", padx=(8, 4))
        tk.Button(frm_rapido, text="✗ Ninguno",
                  command=self._deseleccionar_todos,
                  bg="#FEF0E8", fg="#8B3A00",
                  font=FONT_SMALL, padx=8, pady=2,
                  cursor="hand2", relief="flat",
                  activebackground="#F8D8C0").pack(side="left")

        tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=(0, 8))

        # Lista de centros con checkboxes
        frm_lista = tk.Frame(cuerpo, bg=C_BG)
        frm_lista.pack(fill="x")

        for i, (codigo, nombre) in enumerate(_CENTROS_LIST):
            var = tk.BooleanVar(value=True)
            self._vars[codigo] = var

            fila_bg = "#F4F7FB" if i % 2 == 0 else C_BG
            frm_fila = tk.Frame(frm_lista, bg=fila_bg, padx=6, pady=3)
            frm_fila.pack(fill="x")

            cb = tk.Checkbutton(frm_fila,
                                variable=var,
                                bg=fila_bg,
                                activebackground=fila_bg,
                                cursor="hand2",
                                relief="flat")
            cb.pack(side="left")

            tk.Label(frm_fila,
                     text=f"{codigo}",
                     bg=fila_bg, fg="#285FA8",
                     font=("Segoe UI", 9, "bold"),
                     width=6, anchor="w").pack(side="left")

            tk.Label(frm_fila,
                     text=nombre,
                     bg=fila_bg, fg=C_TEXT,
                     font=FONT_LABEL,
                     anchor="w").pack(side="left", padx=(4, 0))

        tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=(10, 0))

        # Contador de centros seleccionados (se actualiza en tiempo real)
        self._lbl_contador = tk.Label(cuerpo, text="",
                                      bg=C_BG, fg="#444",
                                      font=("Segoe UI", 8, "italic"))
        self._lbl_contador.pack(anchor="e", pady=(4, 0))
        self._actualizar_contador()

        # Trazar cambios en las variables para actualizar el contador
        for var in self._vars.values():
            var.trace_add("write", lambda *_: self._actualizar_contador())

        # Etiqueta de error
        self._lbl_error = tk.Label(cuerpo, text="",
                                   bg=C_BG, fg=C_ERROR,
                                   font=("Segoe UI", 8, "italic"))
        self._lbl_error.pack(anchor="w", pady=(2, 8))

        tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=(0, 10))

        # Botones principales
        tk.Button(cuerpo,
                  text="✔  Confirmar selección y continuar",
                  command=self._confirmar,
                  bg="#1A7A40", fg="white",
                  font=FONT_BOLD, padx=12, pady=10,
                  cursor="hand2", relief="flat",
                  activebackground="#228B4A").pack(fill="x", pady=(0, 6))

        tk.Button(cuerpo, text="Cancelar",
                  command=self._cancelar,
                  font=FONT_LABEL, padx=12, pady=7,
                  cursor="hand2", relief="flat",
                  bg="#D8DEE8", activebackground="#C0C8D8").pack(fill="x")

    # ── Helpers UI ────────────────────────────────────────────────────────────
    def _seleccionar_todos(self):
        for var in self._vars.values():
            var.set(True)

    def _deseleccionar_todos(self):
        for var in self._vars.values():
            var.set(False)

    def _actualizar_contador(self):
        n = sum(1 for v in self._vars.values() if v.get())
        total = len(self._vars)
        color = C_ERROR if n == 0 else ("#1A7A40" if n == total else "#285FA8")
        self._lbl_contador.config(
            text=f"{n} de {total} centro(s) seleccionado(s)",
            fg=color)

    # ── Acciones ──────────────────────────────────────────────────────────────
    def _confirmar(self):
        self._lbl_error.config(text="")
        seleccionados = [
            (codigo, nombre)
            for (codigo, nombre) in _CENTROS_LIST
            if self._vars[codigo].get()
        ]
        if not seleccionados:
            self._lbl_error.config(
                text="⚠  Debe seleccionar al menos un centro para continuar.")
            return
        self.destroy()
        self._on_confirm(seleccionados)

    def _cancelar(self):
        self.destroy()
        self._on_confirm(False)

# ─────────────────────────────────────────────────────────────────────────────
#  PASO 6 — VENTANA DE CONFIRMACIÓN FINAL
# ─────────────────────────────────────────────────────────────────────────────

class VentanaConfirmacionRINFO(tk.Toplevel):

    def __init__(self, parent, nombre_nuevo: str, ruta_nuevo: str,
                 nombre_sap: str, codigo_sap: str,
                 hojas_vaciadas: list, articulos_copiados: int):
        super().__init__(parent)
        self.title("Registro Info creado correctamente")
        self.configure(bg=C_BG)
        self.resizable(True, True)
        self._ruta_nuevo = ruta_nuevo
        cargar_icono_ventana(self)

        self._construir_ui(nombre_nuevo, ruta_nuevo, nombre_sap, codigo_sap,
                            hojas_vaciadas, articulos_copiados)
        self._ajustar_geometria()
        self.grab_set()

    def _ajustar_geometria(self):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w  = min(660, sw - 40)
        # Alto: contenido real + margen; mínimo 520 para que los botones sean visibles
        h  = max(520, min(self.winfo_reqheight() + 40, sh - 60))
        x  = (sw - w) // 2
        y  = max(20, (sh - h) // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.minsize(560, 520)

    def _construir_ui(self, nombre_nuevo, ruta_nuevo, nombre_sap, codigo_sap,
                       hojas_vaciadas, articulos_copiados):
        carpeta_guardado = os.path.dirname(ruta_nuevo)

        hdr = tk.Frame(self, bg="#1A7A40", pady=12, padx=16)
        hdr.pack(fill="x")
        tk.Label(hdr, text="✔  Registro Info creado correctamente",
                 bg="#1A7A40", fg="white", font=FONT_TITLE).pack(anchor="w")
        tk.Label(hdr,
                 text="El archivo ha sido generado y está listo para usarse en SAP.",
                 bg="#1A7A40", fg="#A8E6C0", font=FONT_SUB).pack(anchor="w", pady=(4, 0))

        # ── Zona de botones — declarada ANTES del cuerpo para que pack()
        #    la reserve en la parte inferior antes de que cuerpo expanda.
        frm_pie = tk.Frame(self, bg=C_BG, padx=22, pady=12)
        frm_pie.pack(side="bottom", fill="x")

        tk.Frame(frm_pie, bg="#D0D8E8", height=1).pack(fill="x", pady=(0, 10))

        tk.Label(frm_pie, text="¿Desea abrir el archivo ahora?",
                 bg=C_BG, fg=C_TEXT,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 8))

        tk.Button(frm_pie, text="✔  Sí, abrir ahora",
                  command=self._abrir_y_cerrar,
                  bg="#1A7A40", fg="white",
                  font=FONT_BOLD, padx=14, pady=11,
                  cursor="hand2", relief="flat",
                  activebackground="#228B4A").pack(fill="x", pady=(0, 6))

        tk.Button(frm_pie, text="No, cerrar",
                  command=self._cerrar,
                  font=FONT_LABEL, padx=14, pady=8,
                  cursor="hand2", relief="flat",
                  bg="#D8DEE8",
                  activebackground="#C0C8D8").pack(fill="x")

        cuerpo = tk.Frame(self, bg=C_BG, padx=22, pady=16)
        cuerpo.pack(fill="both", expand=True)

        tk.Label(cuerpo, text="✔  Archivo generado:",
                 bg=C_BG, fg=C_TEXT, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(cuerpo, text=f"  📄  {nombre_nuevo}",
                 bg="#EEF7EE", fg="#1B5E20",
                 font=("Segoe UI", 10, "bold"),
                 padx=10, pady=8, relief="flat").pack(fill="x", pady=(4, 10))

        tk.Label(cuerpo, text="Guardado en:",
                 bg=C_BG, fg=C_TEXT, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(cuerpo, text=f"  📁  {carpeta_guardado}",
                 bg="#F5F5F5", fg="#333",
                 font=FONT_LABEL, padx=10, pady=6, relief="flat",
                 wraplength=560, justify="left").pack(fill="x", pady=(4, 10))

        tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=(0, 10))

        tk.Label(cuerpo, text="Valores escritos en la hoja principal:",
                 bg=C_BG, fg=C_TEXT, font=("Segoe UI", 9, "bold")).pack(anchor="w")

        frm_vals = tk.Frame(cuerpo, bg="#F0F5FB", padx=12, pady=8)
        frm_vals.pack(fill="x", pady=(4, 10))
        _fila_valor(frm_vals, "D2 — Proveedor SAP:", nombre_sap)
        _fila_valor(frm_vals, "C2 — Código SAP:",    codigo_sap or "(sin código)")
        _fila_valor(frm_vals, "Columna E (desde E2):",
                    f"{articulos_copiados} artículo(s) copiado(s) de la tarifa")

        tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=(0, 10))

        if hojas_vaciadas:
            tk.Label(cuerpo,
                     text=f"Hojas creadas ({len(hojas_vaciadas)}):",
                     bg=C_BG, fg=C_TEXT,
                     font=("Segoe UI", 9, "bold")).pack(anchor="w")
            for h in hojas_vaciadas:
                tk.Label(cuerpo, text=f"     •  {h}",
                         bg=C_BG, fg=C_TEXT, font=FONT_LABEL).pack(anchor="w")

        tk.Frame(cuerpo, bg=C_BG, height=6).pack()

    def _abrir_y_cerrar(self):
        try:
            os.startfile(self._ruta_nuevo)
        except AttributeError:
            subprocess.call(["xdg-open", self._ruta_nuevo])
        except Exception:
            pass
        self._cerrar()

    def _cerrar(self):
        parent = self.master
        self.destroy()
        try:
            if parent and parent.winfo_exists():
                parent.deiconify()
        except Exception:
            pass

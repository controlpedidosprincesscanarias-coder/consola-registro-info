# -*- coding: utf-8 -*-

"""
crear_registro_info.py
----------------------
Módulo independiente — Botón «CREAR REGISTRO INFO»

Flujo completo según especificación:

  1. Seleccionar tarifa válida de la carpeta TARIFAS (sin RINFO, sin INCOMPLETO,
     con patrón  NOMBRE PROVEEDOR [dd-mm-yyyy].xlsx).
  2. Leer la celda C2 de la hoja TARIFA → nombre del proveedor.
  3. Buscar el proveedor en RECURSOS/Proveedores.xlsx por similitud (col A → D2,
     col B → C2 del nuevo RINFO).
  4. Buscar si ya existe un Purchasing Info Records*.xlsx compatible en TARIFAS.
     · Si existe  → VentanaPurchasingEncontrado: continuar con él (GeneradorDesdeExistente,
                    en generador_existente.py) o crear uno nuevo.
     · Si no      → VentanaOpcionesSinPurchasing: descargarlo de SAP (guía de 7 pasos) o crear
                    el RINFO desde cero (CreadorRINFO, en creador_rinfo.py).
  5. Creación del RINFO:
     · Crea un archivo NUEVO desde cero con las mismas 4 hojas del original.
     · Inserta las mismas cabeceras en la fila 1 de cada hoja.
     · Aplica el formato de la fila 2 del archivo original a todas las filas de datos.
     · Escribe D2 = nombre SAP (col A de Proveedores.xlsx).
     · Escribe C2 = código SAP (col B de Proveedores.xlsx).
     · Copia columna A de la tarifa → columna E del RINFO (datos desde fila 2).
     · Guarda como: RINFO NOMBRE ORIGINAL.xlsx  (mismo nombre de base que la tarifa).
  6. Confirmación final con opción de abrir el archivo.

Dependencias:  openpyxl, tkinter (stdlib)
Integración:   llamar a  iniciar_crear_registro_info(app, carpeta_proyecto)
               desde tarifa_app.App.

Desde nov 2026 (etapa 4) este archivo solo contiene las ventanas del flujo (pasos 1-4) e
iniciar_crear_registro_info; el resto está en comun_rinfo.py, ventanas_rinfo.py, creador_rinfo.py
y generador_existente.py. Reexporta todos los nombres de siempre.

Central Compras Princess Canarias — (c) VAMA 2026
"""

import os
import webbrowser
import tkinter as tk
from tkinter import ttk, messagebox
from app.comun.tarifa_constants import (
    C_MID,
    C_BG, C_PANEL, C_TEXT, C_TEXT_LIGHT, C_HOVER, C_ACCENT2, FONT_TITLE,
    FONT_SUB, FONT_LABEL, FONT_BOLD,
)
from app.comun.tarifa_core import cargar_icono_ventana
# normalizar_codigo_sap vive en sap_helpers.py (sin tkinter) para que
# también pueda importarla test_crear_registro_info.py sin arrastrar
# la UI. Ver ese fichero para la implementación y el porqué.

# (nov 2026, etapa 4) Este archivo se ha repartido en comun_rinfo.py, ventanas_rinfo.py,
# creador_rinfo.py y generador_existente.py. Aquí quedan las ventanas del flujo (que unen el
# Creador y el Generador) y iniciar_crear_registro_info. Se reexportan todos los nombres de
# siempre, así que «from app.rinfo.crear_registro_info import X» sigue funcionando.
from app.rinfo.comun_rinfo import (  # noqa: F401  (reexportados: tarifa_app, pruebas)
    ARCHIVO_PROV,
    ARCHIVO_TEMPLATE,
    CENTROS_SAP_CANARIAS,
    ESTRUCTURA_HOJAS,
    FECHA_FIN_FIJA_HOJA1,
    FECHA_INICIO_FIJA_HOJA1,
    FOLDER_RECURSOS,
    FOLDER_TARIFAS,
    HOJAS_PURCHASING,
    HOJA_TARIFA,
    UMBRAL_PURCHASING,
    UMBRAL_SIMILITUD,
    URL_SAP,
    _CENTROS_LIST,
    _LogAgrupado,
    _RE_FORMAS_SOC,
    _RE_INCOMPLETO,
    _RE_PATRON_TARIFA,
    _RE_RINFO,
    _buscar_proveedor_en_xlsx,
    _buscar_purchasing_compatible,
    _eliminar_filas_agrupadas,
    _fila_valor,
    _leer_proveedor_tarifa,
    _listar_tarifas_validas,
    _localizar_ruta_listado_dali,
    _localizar_ruta_productos,
    _log,
    _normalizar,
    _normalizar_sin_soc,
    _palabras_significativas,
    _parchear_column_dimension,
    _postprocesar_para_sap,
    _restaurar_column_dimension,
    _safe_after,
    _similitud,
    _sin_tildes,
    aplicar_formato_fila,
    copiar_template_registro_info,
)
from app.rinfo.ventanas_rinfo import (  # noqa: F401  (reexportados: tarifa_app, pruebas)
    VentanaCentrosRINFO,
    VentanaConfirmacionGenerado,
    VentanaConfirmacionRINFO,
    VentanaFechasRINFO,
    VentanaInstruccionesSAP,
    VentanaMaterialesParaReactivar,
    VentanaTelefonoRINFO,
    VentanaTemplateAusente,
    _ventana_sin_tarifa,
)
from app.rinfo.creador_rinfo import (  # noqa: F401  (reexportados: tarifa_app, pruebas)
    CreadorRINFO,
)
from app.rinfo.generador_existente import (  # noqa: F401  (reexportados: tarifa_app, pruebas)
    GeneradorDesdeExistente,
    _EstadoIndicesH3,
)


# ─────────────────────────────────────────────────────────────────────────────
#  PUNTO DE ENTRADA PÚBLICO
# ─────────────────────────────────────────────────────────────────────────────

def iniciar_crear_registro_info(app, carpeta_proyecto: str):
    """
    Punto de entrada llamado desde tarifa_app.App al pulsar «CREAR REGISTRO INFO».
    """
    if not carpeta_proyecto or not carpeta_proyecto.strip():
        messagebox.showwarning(
            "Sin proyecto",
            "Seleccione primero la carpeta del Proyecto antes de continuar.")
        return

    carpeta_tarifas = os.path.join(carpeta_proyecto.strip(), FOLDER_TARIFAS)
    if not os.path.isdir(carpeta_tarifas):
        messagebox.showerror(
            "Carpeta TARIFAS no encontrada",
            f"❌  No se encontró la carpeta TARIFAS en:\n\n"
            f"   {carpeta_tarifas}\n\n"
            "Verifique que la estructura del proyecto es correcta.")
        return

    archivos = _listar_tarifas_validas(carpeta_tarifas)
    _log(app, f"CREAR REGISTRO INFO: {len(archivos)} tarifa(s) válida(s) encontrada(s).", "INFO")

    if not archivos:
        messagebox.showwarning(
            "Sin tarifas disponibles",
            "❌  No se encontraron archivos de tarifa válidos en TARIFAS/.\n\n"
            "Los archivos deben cumplir:\n"
            "  • Patrón:  NOMBRE PROVEEDOR [dd-mm-yyyy].xlsx\n"
            "  • Sin etiqueta  [INCOMPLETO]\n"
            "  • Sin prefijo  RINFO\n\n"
            f"📁  Carpeta revisada:\n   {carpeta_tarifas}")
        return

    ventana = VentanaSeleccionTarifa(app, archivos, carpeta_tarifas)
    app.wait_window(ventana)

# ─────────────────────────────────────────────────────────────────────────────
#  PASO 1 — VENTANA DE SELECCIÓN DE TARIFA
# ─────────────────────────────────────────────────────────────────────────────

class VentanaSeleccionTarifa(tk.Toplevel):

    def __init__(self, parent, archivos: list, carpeta_tarifas: str):
        super().__init__(parent)
        self.title("Crear Registro Info — Seleccionar tarifa")
        self.resizable(True, True)
        self.minsize(680, 360)
        self.configure(bg=C_BG)
        self._intentar_icono()

        self._parent          = parent
        self._archivos        = archivos
        self._carpeta_tarifas = carpeta_tarifas
        self._procesando      = False   # guard anti-reentrada

        self._construir_ui()
        self.grab_set()
        self.transient(parent)
        self._centrar(parent)

    def _centrar(self, parent):
        self.update_idletasks()
        px = parent.winfo_rootx() + parent.winfo_width()  // 2
        py = parent.winfo_rooty() + parent.winfo_height() // 2
        w, h = self.winfo_width(), self.winfo_height()
        self.geometry(f"+{px - w // 2}+{py - h // 2}")

    def _intentar_icono(self):
        cargar_icono_ventana(self)

    def _construir_ui(self):
        hdr = tk.Frame(self, bg=C_MID, pady=12, padx=16)
        hdr.pack(fill="x")
        tk.Label(hdr, text="📋  Crear Registro Info SAP",
                 bg=C_MID, fg="white", font=FONT_TITLE).pack(anchor="w")
        tk.Label(hdr,
                 text="Seleccione el archivo de tarifa con el que desea trabajar:",
                 bg=C_MID, fg="#BDD5F0", font=FONT_SUB).pack(anchor="w", pady=(4, 0))

        frm_tabla = tk.Frame(self, bg=C_BG, padx=14, pady=10)
        frm_tabla.pack(fill="both", expand=True)

        estilo = ttk.Style()
        estilo.configure("CRI.Treeview",
                          background=C_PANEL, fieldbackground=C_PANEL,
                          foreground=C_TEXT, rowheight=26, font=FONT_LABEL)
        estilo.configure("CRI.Treeview.Heading",
                          background=C_MID, foreground="white",
                          font=FONT_BOLD, relief="flat")
        estilo.map("CRI.Treeview",
                    background=[("selected", "#285FA8")],
                    foreground=[("selected", "white")])

        cols = ("nombre", "fecha_mod")
        self.tree = ttk.Treeview(frm_tabla, columns=cols, show="headings",
                                  selectmode="browse", style="CRI.Treeview")
        self.tree.heading("nombre",    text="Archivo de tarifa")
        self.tree.heading("fecha_mod", text="Última modificación")
        self.tree.column("nombre",    width=450, anchor="w")
        self.tree.column("fecha_mod", width=170, anchor="center")

        sb = ttk.Scrollbar(frm_tabla, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.tree.tag_configure("par",   background="#F0F5FB", foreground=C_TEXT)
        self.tree.tag_configure("impar", background="white",   foreground=C_TEXT)

        primer_item = None
        for i, archivo in enumerate(self._archivos):
            tag = "par" if i % 2 == 0 else "impar"
            self.tree.insert("", "end", iid=archivo["ruta"],
                              values=(f"  {archivo['nombre']}", archivo["fecha_mod"]),
                              tags=(tag,))
            if primer_item is None:
                primer_item = archivo["ruta"]

        if primer_item:
            self.tree.selection_set(primer_item)

        self.tree.bind("<Double-1>", lambda e: self._confirmar())

        frm_btn = tk.Frame(self, bg=C_BG, padx=14, pady=10)
        frm_btn.pack(fill="x", side="bottom")

        tk.Button(frm_btn, text="✔  Continuar",
                  command=self._confirmar,
                  bg="#285FA8", fg="white", font=FONT_BOLD,
                  padx=16, pady=7, cursor="hand2", relief="flat",
                  activebackground="#3A7AC8").pack(side="right", padx=(6, 0))

        tk.Button(frm_btn, text="Cancelar", command=self.destroy,
                  font=FONT_LABEL, padx=14, pady=6, cursor="hand2",
                  relief="flat", bg="#D8DEE8",
                  activebackground="#C0C8D8").pack(side="right")

    def _confirmar(self):
        if self._procesando:          # evita reentrada por deiconify/grab de Tk
            return
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Sin selección",
                                   "Seleccione una tarifa de la lista antes de continuar.",
                                   parent=self)
            return
        ruta_sel   = sel[0]
        nombre_sel = next(
            (a["nombre_archivo"] for a in self._archivos if a["ruta"] == ruta_sel), "")

        _log(self._parent, f"Tarifa seleccionada: {nombre_sel}", "INFO")
        self._procesando = True        # bloquear re-llamadas mientras se procesa
        self._procesar_tarifa(ruta_sel, nombre_sel)

    def _procesar_tarifa(self, ruta_tarifa: str, nombre_archivo: str):
        nombre_proveedor_tarifa = _leer_proveedor_tarifa(ruta_tarifa)

        if not nombre_proveedor_tarifa:
            messagebox.showerror(
                "Proveedor no encontrado",
                f"❌  No se pudo leer el nombre de proveedor de la celda C2\n"
                f"   de la hoja '{HOJA_TARIFA}' en el archivo:\n\n"
                f"   {nombre_archivo}\n\n"
                "Compruebe que la hoja existe y que C2 contiene el nombre del proveedor.",
                parent=self)
            _log(self._parent, f"Proveedor no legible en C2 de: {nombre_archivo}", "ERROR")
            return

        _log(self._parent,
             f"Proveedor detectado en tarifa (C2): «{nombre_proveedor_tarifa}»", "OK")

        carpeta_proyecto  = os.path.dirname(self._carpeta_tarifas)
        ruta_proveedores  = os.path.join(carpeta_proyecto, FOLDER_RECURSOS, ARCHIVO_PROV)
        nombre_sap        = None
        codigo_sap        = None

        if not os.path.isfile(ruta_proveedores):
            _log(self._parent,
                 f"Proveedores.xlsx no encontrado en: {os.path.dirname(ruta_proveedores)}",
                 "WARN")
        else:
            candidatos = _buscar_proveedor_en_xlsx(nombre_proveedor_tarifa, ruta_proveedores)
            _log(self._parent,
                 f"Proveedores.xlsx — {len(candidatos)} coincidencia(s) (≥{UMBRAL_SIMILITUD}%).",
                 "INFO")

            if len(candidatos) == 1:
                _, nombre_sap, codigo_sap = candidatos[0]
                _log(self._parent,
                     f"Proveedor SAP: «{nombre_sap}» [cód. {codigo_sap}]", "OK")
            elif len(candidatos) > 1:
                nombre_sap, codigo_sap = self._dialogo_seleccion_proveedor(
                    nombre_proveedor_tarifa, candidatos)
                if nombre_sap:
                    _log(self._parent,
                         f"Usuario seleccionó proveedor SAP: «{nombre_sap}» "
                         f"[cód. {codigo_sap}]", "OK")
                else:
                    _log(self._parent,
                         "Usuario no seleccionó ningún proveedor SAP — "
                         "se usará el nombre de la tarifa.", "WARN")
            else:
                _log(self._parent,
                     f"Sin coincidencia ≥{UMBRAL_SIMILITUD}% para "
                     f"«{nombre_proveedor_tarifa}» en Proveedores.xlsx.", "WARN")
                messagebox.showwarning(
                    "Proveedor no localizado en Proveedores.xlsx",
                    f"⚠  No se encontró ninguna coincidencia suficiente\n"
                    f"   para el proveedor:\n\n"
                    f"   «{nombre_proveedor_tarifa}»\n\n"
                    "Se usará el nombre tal como aparece en la tarifa.\n\n"
                    "Si el nombre es incorrecto, revise Proveedores.xlsx\n"
                    "o corrija la celda C2 de la tarifa.",
                    parent=self)

        if not nombre_sap:
            nombre_sap  = nombre_proveedor_tarifa
            codigo_sap  = ""

        _log(self._parent,
             "Buscando Purchasing Info Records compatible en TARIFAS/…", "INFO")
        purchasing = _buscar_purchasing_compatible(nombre_proveedor_tarifa,
                                                   self._carpeta_tarifas)

        if purchasing:
            score_txt = str(purchasing["score"]) + "%"
            _log(self._parent,
                 f"Purchasing encontrado: {purchasing['archivo']} "
                 f"(score {score_txt}) — preguntando al usuario cómo continuar.", "OK")
            self.withdraw()
            VentanaPurchasingEncontrado(
                self,
                ruta_tarifa,
                nombre_archivo,
                nombre_proveedor_tarifa,
                nombre_sap,
                codigo_sap,
                purchasing,
                self._carpeta_tarifas,
                app_root=self._parent,
            )
            return

        _log(self._parent,
             "No se encontró Purchasing Info Records compatible.", "WARN")
        self.withdraw()
        VentanaOpcionesSinPurchasing(
            self, ruta_tarifa, nombre_archivo,
            nombre_proveedor_tarifa, nombre_sap, codigo_sap,
            self._carpeta_tarifas,
            app_root=self._parent)

    def _dialogo_seleccion_proveedor(self, nombre_dali: str,
                                     candidatos: list) -> tuple:
        resultado = [None, None]

        dlg = tk.Toplevel(self)
        dlg.title("Seleccionar proveedor SAP")
        dlg.configure(bg=C_BG)
        dlg.resizable(False, False)
        dlg.grab_set()
        dlg.transient(self)

        hdr_d = tk.Frame(dlg, bg="#285FA8", pady=10, padx=16)
        hdr_d.pack(fill="x")
        tk.Label(hdr_d, text="🔍  Seleccione el proveedor SAP correcto",
                 bg="#285FA8", fg="white", font=FONT_TITLE).pack(anchor="w")
        tk.Label(hdr_d,
                 text=f"Se encontraron {len(candidatos)} coincidencias en Proveedores.xlsx.",
                 bg="#285FA8", fg="#C0D8F8", font=FONT_SUB).pack(anchor="w", pady=(4, 0))

        cuerpo = tk.Frame(dlg, bg=C_BG, padx=18, pady=14)
        cuerpo.pack(fill="both", expand=True)

        tk.Label(cuerpo,
                 text=f"Nombre en la tarifa:  «{nombre_dali}»",
                 bg=C_BG, fg="#888",
                 font=("Segoe UI", 9, "italic")).pack(anchor="w")
        tk.Label(cuerpo,
                 text="Elija el proveedor SAP con el que desea crear el Registro Info:",
                 bg=C_BG, fg=C_TEXT,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(8, 10))

        for score, nombre, codigo in candidatos:
            etiqueta = (f"  {nombre}  [{codigo}]  —  {score}% coincidencia"
                        if codigo else f"  {nombre}  —  {score}% coincidencia")

            def _elegir(n=nombre, c=codigo):
                resultado[0] = n
                resultado[1] = c
                dlg.destroy()

            tk.Button(cuerpo, text=etiqueta,
                      command=_elegir,
                      bg="#EEF2F8", fg="#1A2E4A",
                      font=FONT_LABEL, padx=12, pady=9,
                      cursor="hand2", relief="flat",
                      activebackground="#D0D8F0",
                      anchor="w").pack(fill="x", pady=2)

        tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=(10, 6))

        tk.Button(cuerpo,
                  text="Ninguno — usar el nombre de la tarifa",
                  command=dlg.destroy,
                  bg="#D8DEE8", font=FONT_LABEL, padx=12, pady=7,
                  cursor="hand2", relief="flat",
                  activebackground="#C0C8D8").pack(fill="x", pady=(0, 4))

        dlg.update_idletasks()
        sw = dlg.winfo_screenwidth()
        w  = min(560, sw - 60)
        dlg.geometry(f"{w}+{(sw - w) // 2}+{dlg.winfo_screenheight() // 5}")
        self.wait_window(dlg)

        return (resultado[0], resultado[1])

# ─────────────────────────────────────────────────────────────────────────────
#  PASO 4 (PURCHASING ENCONTRADO) — VentanaPurchasingEncontrado
#
#  Informa al usuario de que se ha encontrado un Purchasing compatible,
#  pregunta si desea continuar usándolo o comenzar desde cero, y a
#  continuación ejecuta SIEMPRE el flujo estándar de recogida de datos:
#      teléfono → fechas → centros
#  Una vez completado ese flujo, bifurca según la elección del usuario:
#      · CONTINUAR  → GeneradorDesdeExistente (cruza materiales con el Purchasing)
#      · DESDE CERO → CreadorRINFO (crea desde la plantilla SAP)
# ─────────────────────────────────────────────────────────────────────────────

class VentanaPurchasingEncontrado(tk.Toplevel):
    """
    Ventana que aparece cuando se detecta un Purchasing Info Records compatible.

    Muestra la información del Purchasing encontrado, pregunta al usuario si
    desea continuar usándolo o empezar desde cero, y a continuación ejecuta
    el flujo estándar (teléfono → fechas → centros) independientemente de
    la elección.  Solo al final del flujo se bifurca el destino.

    Parámetros
    ----------
    parent_ventana   : VentanaSeleccionTarifa (llamante, ya ocultada con withdraw).
    ruta_tarifa      : ruta del archivo de tarifa seleccionado.
    nombre_archivo   : nombre de archivo de la tarifa (con extensión).
    nombre_dali      : nombre de proveedor leído de C2 de la tarifa.
    nombre_sap       : nombre SAP del proveedor (de Proveedores.xlsx).
    codigo_sap       : código SAP del proveedor.
    purchasing       : dict devuelto por _buscar_purchasing_compatible().
    carpeta_tarifas  : ruta de la carpeta TARIFAS del proyecto.
    app_root         : referencia al App raíz (para log y barra de progreso).
    """

    def __init__(self, parent_ventana, ruta_tarifa: str, nombre_archivo: str,
                 nombre_dali: str, nombre_sap: str, codigo_sap: str,
                 purchasing: dict, carpeta_tarifas: str, app_root=None):
        super().__init__(parent_ventana.master if hasattr(parent_ventana, "master")
                         else parent_ventana)
        self.title("Registro Info — Purchasing Info Records encontrado")
        self.resizable(True, True)
        self.configure(bg=C_BG)
        cargar_icono_ventana(self)

        self._parent_ventana  = parent_ventana
        self._ruta_tarifa     = ruta_tarifa
        self._nombre_archivo  = nombre_archivo
        self._nombre_dali     = nombre_dali
        self._nombre_sap      = nombre_sap
        self._codigo_sap      = codigo_sap
        self._purchasing      = purchasing
        self._carpeta_tarifas = carpeta_tarifas
        self._app_root        = app_root

        # La elección del usuario se guardará aquí antes de iniciar el flujo
        self._eleccion = None   # "continuar" | "desde_cero"

        self._construir_ui()
        self.grab_set()
        self.transient(self.master)
        self._ajustar_geometria()

    def _ajustar_geometria(self):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w  = min(700, sw - 40)
        h  = min(sh - 60, 560)
        x  = (sw - w) // 2
        y  = 20
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.minsize(560, 420)

    def _construir_ui(self):
        etiqueta_prov = (f"{self._nombre_sap}  [{self._codigo_sap}]"
                         if self._codigo_sap else self._nombre_sap)

        hdr = tk.Frame(self, bg="#285FA8", pady=12, padx=16)
        hdr.pack(fill="x")
        tk.Label(hdr,
                 text="✔  Se ha encontrado un Registro Info para este proveedor",
                 bg="#285FA8", fg="white", font=FONT_TITLE).pack(anchor="w")
        tk.Label(hdr,
                 text="Se ha localizado un Purchasing Info Records compatible en la carpeta TARIFAS.",
                 bg="#285FA8", fg="#BDD5F0", font=FONT_SUB).pack(anchor="w", pady=(4, 0))

        cuerpo = tk.Frame(self, bg=C_BG, padx=22, pady=14)
        cuerpo.pack(fill="both", expand=True)

        tk.Label(cuerpo,
                 text="Proveedor identificado en la tarifa (C2):",
                 bg=C_BG, fg="#999",
                 font=("Segoe UI", 8, "italic")).pack(anchor="w")
        tk.Label(cuerpo, text=f"  {self._nombre_dali}",
                 bg="#F8F4FF", fg="#555",
                 font=FONT_LABEL, padx=10, pady=5, relief="flat").pack(fill="x")

        tk.Frame(cuerpo, bg=C_BG, height=8).pack()

        tk.Label(cuerpo,
                 text="Proveedor SAP a usar:",
                 bg=C_BG, fg=C_TEXT,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(cuerpo, text=f"  {etiqueta_prov}",
                 bg="#EEF2F8", fg="#285FA8",
                 font=("Segoe UI", 11, "bold"),
                 padx=10, pady=8, relief="flat").pack(fill="x")

        tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=10)

        tk.Label(cuerpo,
                 text="Purchasing Info Records compatible encontrado:",
                 bg=C_BG, fg=C_TEXT,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(cuerpo,
                 text=f"  📋  {self._purchasing['archivo']}  "
                      f"({self._purchasing['score']}% coincidencia)",
                 bg="#EBF4FF", fg="#1A3A6A",
                 font=FONT_LABEL, padx=10, pady=7, relief="flat").pack(fill="x")

        tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=10)

        tk.Label(cuerpo,
                 text="¿Desea continuar usando este Purchasing o comenzar desde cero?",
                 bg=C_BG, fg=C_TEXT,
                 font=("Segoe UI", 10, "bold"),
                 wraplength=600, justify="left").pack(anchor="w", pady=(0, 10))

        tk.Button(cuerpo,
                  text="📋  Continuar usando este Purchasing Info Records",
                  command=self._elegir_continuar,
                  bg="#285FA8", fg="white",
                  font=FONT_BOLD, padx=12, pady=11,
                  cursor="hand2", relief="flat",
                  activebackground="#3A7AC8").pack(fill="x", pady=(0, 6))

        tk.Button(cuerpo,
                  text=f"🆕  Comenzar desde cero  —  {etiqueta_prov}",
                  command=self._elegir_desde_cero,
                  bg="#1A7A40", fg="white",
                  font=FONT_BOLD, padx=12, pady=11,
                  cursor="hand2", relief="flat",
                  activebackground="#228B4A").pack(fill="x", pady=(0, 6))

        tk.Button(cuerpo, text="Cancelar",
                  command=self._cancelar,
                  font=FONT_LABEL, padx=12, pady=8,
                  cursor="hand2", relief="flat",
                  bg="#D8DEE8",
                  activebackground="#C0C8D8").pack(fill="x")

    # ── Elección del usuario ───────────────────────────────────────────────────

    def _elegir_continuar(self):
        self._eleccion = "continuar"
        _log(self._parent_ventana,
             "Usuario eligió: Continuar usando el Purchasing existente.", "INFO")
        self._iniciar_flujo_estandar()

    def _elegir_desde_cero(self):
        # Verificar template SAP antes de continuar
        carpeta_proyecto = os.path.dirname(self._carpeta_tarifas)
        ruta_template    = os.path.join(carpeta_proyecto, FOLDER_RECURSOS, ARCHIVO_TEMPLATE)
        if not os.path.isfile(ruta_template):
            ruta_recursos = os.path.join(carpeta_proyecto, FOLDER_RECURSOS)
            _log(self._parent_ventana,
                 "Template SAP no encontrado en RECURSOS/ — creación abortada.", "ERROR")
            VentanaTemplateAusente(self, ruta_recursos, ARCHIVO_TEMPLATE)
            return
        self._eleccion = "desde_cero"
        _log(self._parent_ventana,
             "Usuario eligió: Comenzar desde cero (ignorar Purchasing).", "INFO")
        self._iniciar_flujo_estandar()

    def _cancelar(self):
        _log(self._parent_ventana, "Acción cancelada por el usuario.", "WARN")
        self.destroy()
        self._parent_ventana.withdraw()

    # ── Flujo estándar: teléfono → fechas → centros → bifurcación ─────────────

    def _iniciar_flujo_estandar(self):
        """
        Lanza el flujo estándar de recogida de datos (igual que en
        VentanaOpcionesSinPurchasing._opcion_crear_desde_cero).
        Al terminar, bifurca según self._eleccion.
        """
        self._telefono_pendiente = None
        self._fechas_pendientes  = (None, None)
        self.destroy()   # Cierra esta ventana; el flujo continúa en el padre
        self._parent_ventana.withdraw()   # re-ocultar: destroy() del hijo revierte el withdraw() previo
        VentanaTelefonoRINFO(self._parent_ventana, self._tras_telefono)

    def _tras_telefono(self, resultado):
        """Callback de VentanaTelefonoRINFO."""
        if resultado is False:
            _log(self._parent_ventana,
                 "Creación cancelada en la pantalla de teléfono.", "WARN")
            self._parent_ventana.withdraw()
            return
        self._telefono_pendiente = resultado
        VentanaFechasRINFO(self._parent_ventana, self._tras_fechas)

    def _tras_fechas(self, resultado):
        """Callback de VentanaFechasRINFO."""
        if resultado is False:
            _log(self._parent_ventana,
                 "Creación cancelada en la pantalla de fechas.", "WARN")
            self._parent_ventana.withdraw()
            return
        self._fechas_pendientes = resultado
        fecha_inicio, fecha_fin = resultado
        _log(self._parent_ventana,
             f"Fechas validez: inicio={fecha_inicio}  fin={fecha_fin}", "OK")
        VentanaCentrosRINFO(self._parent_ventana, self._tras_centros)

    def _tras_centros(self, resultado):
        """Callback de VentanaCentrosRINFO. Aquí bifurca según la elección."""
        if resultado is False:
            _log(self._parent_ventana,
                 "Creación cancelada en la pantalla de centros.", "WARN")
            self._parent_ventana.withdraw()
            return

        codigos = ", ".join(c for c, _ in resultado)
        _log(self._parent_ventana,
             f"Centros seleccionados ({len(resultado)}): {codigos}", "OK")

        if self._eleccion == "continuar":
            _log(self._parent_ventana,
                 "Iniciando generación desde Purchasing existente.", "INFO")
            _pv = self._parent_ventana
            _pv.withdraw()  # ocultar sin destruir antes de lanzar el generador
            GeneradorDesdeExistente(
                self._app_root,
                self._ruta_tarifa,
                self._nombre_archivo,
                self._nombre_sap,
                self._codigo_sap,
                self._purchasing,
                self._carpeta_tarifas,
                app_root=self._app_root,
                telefono_proveedor=self._telefono_pendiente,
                fecha_inicio=self._fechas_pendientes[0],
                fecha_fin=self._fechas_pendientes[1],
                centros_activos=resultado,
            )
        else:
            _log(self._parent_ventana,
                 "Iniciando creación de RINFO desde cero.", "INFO")
            _pv = self._parent_ventana
            _pv.withdraw()  # ocultar sin destruir antes de lanzar el creador
            CreadorRINFO(
                self._app_root,
                self._ruta_tarifa,
                self._nombre_archivo,
                self._nombre_sap,
                self._codigo_sap,
                self._carpeta_tarifas,
                app_root=self._app_root,
                telefono_proveedor=self._telefono_pendiente,
                fecha_inicio=self._fechas_pendientes[0],
                fecha_fin=self._fechas_pendientes[1],
                centros_activos=resultado,
            )

# ─────────────────────────────────────────────────────────────────────────────
#  PASO 4 — VENTANA DE OPCIONES CUANDO NO HAY PURCHASING
# ─────────────────────────────────────────────────────────────────────────────

class VentanaOpcionesSinPurchasing(tk.Toplevel):

    def __init__(self, parent_ventana, ruta_tarifa: str, nombre_archivo: str,
                 nombre_dali: str, nombre_sap: str, codigo_sap: str,
                 carpeta_tarifas: str, app_root=None):
        super().__init__(parent_ventana.master if hasattr(parent_ventana, "master")
                         else parent_ventana)
        self.title("Registro Info — Proveedor sin Purchasing en TARIFAS")
        self.resizable(True, True)
        self.configure(bg=C_BG)
        self._intentar_icono()

        self._parent_ventana  = parent_ventana
        self._ruta_tarifa     = ruta_tarifa
        self._nombre_archivo  = nombre_archivo
        self._nombre_dali     = nombre_dali
        self._nombre_sap      = nombre_sap
        self._codigo_sap      = codigo_sap
        self._carpeta_tarifas = carpeta_tarifas
        # Referencia directa al App raíz para actualizar la barra de progreso
        self._app_root        = app_root

        self._construir_ui()
        self.grab_set()
        self.transient(self.master)
        self._ajustar_geometria()

    def _intentar_icono(self):
        cargar_icono_ventana(self)

    def _ajustar_geometria(self):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w  = min(700, sw - 40)
        h  = min(sh - 60, 620)
        x  = (sw - w) // 2
        y  = 20
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.minsize(560, 420)

    def _construir_ui(self):
        etiqueta_prov = (f"{self._nombre_sap}  [{self._codigo_sap}]"
                         if self._codigo_sap else self._nombre_sap)

        hdr = tk.Frame(self, bg="#8B5A00", pady=12, padx=16)
        hdr.pack(fill="x")
        tk.Label(hdr,
                 text="⚠  No se encontró Purchasing Info Records para este proveedor",
                 bg="#8B5A00", fg="white", font=FONT_TITLE).pack(anchor="w")
        tk.Label(hdr,
                 text="No hay ningún archivo Purchasing Info Records*.xlsx en la carpeta TARIFAS.",
                 bg="#8B5A00", fg="#FFD99A", font=FONT_SUB).pack(anchor="w", pady=(4, 0))

        cuerpo = tk.Frame(self, bg=C_BG, padx=22, pady=14)
        cuerpo.pack(fill="both", expand=True)

        tk.Label(cuerpo,
                 text="Proveedor identificado en la tarifa (C2):",
                 bg=C_BG, fg="#999",
                 font=("Segoe UI", 8, "italic")).pack(anchor="w")
        tk.Label(cuerpo, text=f"  {self._nombre_dali}",
                 bg="#F8F4FF", fg="#555",
                 font=FONT_LABEL, padx=10, pady=5, relief="flat").pack(fill="x")

        tk.Frame(cuerpo, bg=C_BG, height=8).pack()

        tk.Label(cuerpo,
                 text="Proveedor SAP a usar (Proveedores.xlsx):",
                 bg=C_BG, fg=C_TEXT,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(cuerpo, text=f"  {etiqueta_prov}",
                 bg="#EEF2F8", fg="#285FA8",
                 font=("Segoe UI", 11, "bold"),
                 padx=10, pady=8, relief="flat").pack(fill="x")

        tk.Frame(cuerpo, bg="#D0D8E8", height=1).pack(fill="x", pady=10)

        frm_sap = tk.Frame(cuerpo, bg="#EBF4FF", padx=10, pady=8, relief="flat")
        frm_sap.pack(fill="x", pady=(0, 12))
        tk.Label(frm_sap,
                 text="🔗  Acceso directo a SAP S/4HANA — Registros Info de Compras:",
                 bg="#EBF4FF", fg="#1A3A6A",
                 font=("Segoe UI", 8, "bold")).pack(anchor="w")
        lnk = tk.Label(frm_sap, text=URL_SAP,
                        bg="#EBF4FF", fg="#1565C0",
                        font=("Segoe UI", 8, "underline"),
                        cursor="hand2")
        lnk.pack(anchor="w", pady=(2, 0))
        lnk.bind("<Button-1>", lambda e: webbrowser.open(URL_SAP))

        tk.Label(cuerpo,
                 text="Seleccione una de las siguientes opciones:",
                 bg=C_BG, fg=C_TEXT,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(0, 8))

        tk.Button(cuerpo,
                  text="📥  Descargar el Purchasing Info Records desde SAP primero",
                  command=self._opcion_descargar_sap,
                  bg=C_MID, fg=C_TEXT_LIGHT,
                  font=FONT_BOLD, padx=12, pady=11,
                  cursor="hand2", relief="flat",
                  activebackground=C_HOVER,
                  activeforeground=C_ACCENT2).pack(fill="x", pady=(0, 6))

        tk.Button(cuerpo,
                  text=f"🆕  Crear nuevo Registro Info desde cero  —  {etiqueta_prov}",
                  command=self._opcion_crear_desde_cero,
                  bg="#1A7A40", fg="white",
                  font=FONT_BOLD, padx=12, pady=11,
                  cursor="hand2", relief="flat",
                  activebackground="#228B4A").pack(fill="x", pady=(0, 6))

        tk.Button(cuerpo, text="Cancelar",
                  command=self._cancelar,
                  font=FONT_LABEL, padx=12, pady=8,
                  cursor="hand2", relief="flat",
                  bg="#D8DEE8",
                  activebackground="#C0C8D8").pack(fill="x")

    def _opcion_descargar_sap(self):
        _log(self._parent_ventana,
             "Usuario eligió: Descargar Purchasing desde SAP.", "INFO")
        self.destroy()
        self._parent_ventana.withdraw()
        VentanaInstruccionesSAP(self.master, self._nombre_sap)

    def _opcion_crear_desde_cero(self):
        _log(self._parent_ventana,
             "Usuario eligió: Crear Registro Info nuevo desde cero.", "INFO")

        # ── Verificar template SAP ANTES de abrir cualquier ventana ──────────
        carpeta_proyecto = os.path.dirname(self._carpeta_tarifas)
        ruta_template    = os.path.join(carpeta_proyecto, FOLDER_RECURSOS, ARCHIVO_TEMPLATE)

        if not os.path.isfile(ruta_template):
            ruta_recursos = os.path.join(carpeta_proyecto, FOLDER_RECURSOS)
            _log(self._parent_ventana,
                 "Template SAP no encontrado en RECURSOS/ — creación abortada.", "ERROR")
            VentanaTemplateAusente(self, ruta_recursos, ARCHIVO_TEMPLATE)
            return

        self.destroy()

        def _tras_centros(resultado_centros):
            """Callback de VentanaCentrosRINFO."""
            if resultado_centros is False:
                _log(self._parent_ventana,
                     "Creación de RINFO cancelada en la pantalla de centros.", "WARN")
                self._parent_ventana.withdraw()
                return
            codigos = ", ".join(c for c, _ in resultado_centros)
            _log(self._parent_ventana,
                 f"Centros seleccionados ({len(resultado_centros)}): {codigos}", "OK")
            CreadorRINFO(
                self._parent_ventana,
                self._ruta_tarifa,
                self._nombre_archivo,
                self._nombre_sap,
                self._codigo_sap,
                self._carpeta_tarifas,
                app_root=self._app_root,
                telefono_proveedor=self._telefono_pendiente,
                fecha_inicio=self._fechas_pendientes[0],
                fecha_fin=self._fechas_pendientes[1],
                centros_activos=resultado_centros,
            )

        def _tras_fechas(resultado_fechas):
            """Callback de VentanaFechasRINFO."""
            if resultado_fechas is False:
                _log(self._parent_ventana,
                     "Creación de RINFO cancelada en la pantalla de fechas.", "WARN")
                self._parent_ventana.withdraw()
                return
            self._fechas_pendientes = resultado_fechas
            fecha_inicio, fecha_fin = resultado_fechas
            _log(self._parent_ventana,
                 f"Fechas validez: inicio={fecha_inicio}  fin={fecha_fin}", "OK")
            VentanaCentrosRINFO(self._parent_ventana, _tras_centros)

        def _tras_telefono(resultado):
            """Callback de VentanaTelefonoRINFO."""
            if resultado is False:
                # El usuario canceló — no se crea el RINFO
                _log(self._parent_ventana,
                     "Creación de RINFO cancelada en la pantalla de teléfono.", "WARN")
                self._parent_ventana.withdraw()
                return
            # Guardar teléfono temporalmente y pasar a pantalla de fechas
            self._telefono_pendiente = resultado   # str validado o None
            VentanaFechasRINFO(self._parent_ventana, _tras_fechas)

        self._telefono_pendiente = None
        self._fechas_pendientes  = (None, None)
        VentanaTelefonoRINFO(self._parent_ventana, _tras_telefono)

    def _cancelar(self):
        _log(self._parent_ventana, "Acción cancelada por el usuario.", "WARN")
        self.destroy()
        self._parent_ventana.withdraw()


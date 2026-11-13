# -*- coding: utf-8 -*-
"""
accion_listado_dali.py
Acción 2 — Actualizar los artículos DALI del catálogo.

Mixin de App (tarifa_app.py): sus métodos se usan como self.metodo(). Separado de tarifa_app.py
en la etapa 4 (nov 2026) sin cambiar el código.
Central Compras Princess Canarias · La idea y el desarrollo de la aplicación son de
Víctor Alejandro Martín Andrade (VAMA).

v10.2.54 (decisión de Víctor, 13 nov 2026): ya no existe «LISTADO CODIGOS DALI - SAP.xlsx». Antes esta acción
copiaba código, descripción, unidad y activo de ARTICULOS_DALI.xlsx al LISTADO POR POSICIÓN DE FILA (fila 8+n),
dejando el código SAP de la columna D en su fila: si DALI quitaba o intercalaba un código, los SAP de debajo
quedaban emparejados con el artículo de al lado. Ahora regenera directamente el catálogo completo, emparejando
siempre por código DALI: los datos DALI salen de ARTICULOS_DALI.xlsx y los SAP, el código de proveedor y el precio
se conservan del catálogo anterior (ver _regenerar_catalogo_dali_completo_impl).
"""

import os
from tkinter import messagebox
import threading
from app.comun.tarifa_core import (
    FOLDER_RECURSOS,
)


class AccionListadoDali:
    def _actualizar_listado_dali(self):
        """
        PUNTO 2 — Actualizar los artículos DALI del catálogo.
        Origen : RECURSOS/ARTICULOS_DALI.xlsx  (generado por Punto 1)
        Destino: RECURSOS/catalogo-dali-completo.xlsx (por código DALI)
        """
        carpeta = self._carpeta_proyecto.get().strip()
        if not carpeta:
            messagebox.showwarning("Sin proyecto",
                                   "Seleccione primero la carpeta del Proyecto.")
            return

        if not self._confirmar_catalogo_actualizado():
            self.log_consola(
                "Actualización DALI cancelada por el usuario (catálogo desactualizado).", "WARN")
            return

        # ── Rutas ────────────────────────────────────────────────
        ruta_origen   = os.path.join(carpeta, FOLDER_RECURSOS, "ARTICULOS_DALI.xlsx")
        ruta_catalogo = os.path.join(carpeta, FOLDER_RECURSOS, "catalogo-dali-completo.xlsx")

        # ── Verificar ARTICULOS_DALI.xlsx ────────────────────────
        if not os.path.isfile(ruta_origen):
            messagebox.showwarning(
                "Archivo no encontrado",
                "❌  No se encontró el archivo:\n\n"
                f"   {ruta_origen}\n\n"
                "Debe ejecutar primero la opción\n"
                "▶  1 — Organizar artículos DALI\n\n"
                "para generar este archivo antes de continuar.")
            return

        # ── Leer fecha del encabezado (celda B3) ─────────────────
        try:
            import openpyxl
            wb_orig = openpyxl.load_workbook(ruta_origen, read_only=True, data_only=True)
            celda_b3 = wb_orig.active["B3"].value or ""
            wb_orig.close()
        except Exception as e:
            messagebox.showerror("Error al leer origen",
                                 f"No se pudo leer ARTICULOS_DALI.xlsx:\n{e}")
            return
        texto_fecha = str(celda_b3).strip() if celda_b3 else "(fecha no disponible)"

        # ── Confirmación al usuario ───────────────────────────────
        if not messagebox.askyesno(
                "Confirmar actualización DALI",
                f"Archivo origen encontrado:\n"
                f"   ARTICULOS_DALI.xlsx\n\n"
                f"   {texto_fecha}\n\n"
                f"Destino:\n"
                f"   RECURSOS/catalogo-dali-completo.xlsx\n\n"
                "¿Desea continuar?"):
            self.log_consola("Actualización DALI cancelada por el usuario.", "WARN")
            return

        # ── Proceso en hilo ──────────────────────────────────────
        def proceso():
            self.consola.separador("INICIO ACTUALIZACIÓN DALI")
            self.log_consola("Iniciando actualización de los artículos DALI del catálogo...", "INFO")
            self.after(0, lambda: self._barra.indeterminate("Leyendo ARTICULOS_DALI..."))
            try:
                import openpyxl
                wb = openpyxl.load_workbook(ruta_origen, read_only=True, data_only=True)
                n_filas = 0
                for fila in wb.active.iter_rows(min_row=8, max_col=1, values_only=True):
                    if fila and fila[0] not in (None, ""):
                        n_filas += 1
                wb.close()
                self.log_consola(f"Artículos en ARTICULOS_DALI.xlsx: {n_filas}", "OK")
                if n_filas == 0:
                    self.after(0, lambda: self._barra.reset("Sin datos en origen"))
                    self.after(0, lambda: messagebox.showwarning(
                        "Sin datos",
                        "No se encontraron filas con datos a partir de la fila 8\n"
                        "en el archivo ARTICULOS_DALI.xlsx."))
                    return

                self._regenerar_catalogo_dali_completo()
                if not os.path.isfile(ruta_catalogo):
                    raise RuntimeError("no se ha podido crear catalogo-dali-completo.xlsx")

                self.log_consola(f"  └ {texto_fecha}", "INFO")
                self.consola.separador("FIN ACTUALIZACIÓN DALI")
                self.after(0, lambda: self._barra.completado(
                    f"DALI actualizado — {n_filas} artículos ✔"))
                self.after(0, lambda n=n_filas: messagebox.showinfo(
                    "Actualizacion completada",
                    f"✔  Artículos DALI actualizados en el catálogo\n\n"
                    f"   Artículos DALI : {n}\n"
                    f"   {texto_fecha}\n\n"
                    f"\U0001f4c1  {ruta_catalogo}"))

            except PermissionError:
                self.log_consola("ERROR: catalogo-dali-completo.xlsx está abierto en Excel.", "ERROR")
                self.consola.separador("FIN ACTUALIZACIÓN DALI (con error)")
                self.after(0, lambda: self._barra.reset("Error al guardar"))
                self.after(0, lambda: messagebox.showerror(
                    "Archivo bloqueado por Excel",
                    "❌  No se pudo guardar RECURSOS/catalogo-dali-completo.xlsx:\n"
                    "está abierto en Excel.\n\n"
                    "Ciérrelo y vuelva a ejecutar la Acción 2."))
            except Exception as e:
                self.log_consola(f"ERROR: {e}", "ERROR")
                self.consola.separador("FIN ACTUALIZACIÓN DALI (con error)")
                self.after(0, lambda: self._barra.reset("Error en actualización DALI"))
                self.after(0, lambda err=str(e): messagebox.showerror(
                    "Error en actualización DALI",
                    f"✘  El proceso no se completó correctamente.\n\n"
                    f"Detalle:\n{err}\n\n"
                    "Revise la ventana de log para más información."))

        threading.Thread(target=proceso, daemon=True).start()

# -*- coding: utf-8 -*-
"""
accion_listado_dali.py
Pasar ARTICULOS_DALI.xlsx al catálogo completo y a Catálogo Asignaciones (antes, Acción 2).

Mixin de App (tarifa_app.py): sus métodos se usan como self.metodo(). Separado de tarifa_app.py
en la etapa 4 (nov 2026) sin cambiar el código.
Central Compras Princess Canarias · La idea y el desarrollo de la aplicación son de
Víctor Alejandro Martín Andrade (VAMA).

v10.2.50 (decisión de Víctor, 13 nov 2026): ya no existe «LISTADO CODIGOS DALI - SAP.xlsx». Antes esta acción
copiaba código, descripción, unidad y activo de ARTICULOS_DALI.xlsx al LISTADO POR POSICIÓN DE FILA (fila 8+n),
dejando el código SAP de la columna D en su fila: si DALI quitaba o intercalaba un código, los SAP de debajo
quedaban emparejados con el artículo de al lado. Ahora regenera directamente el catálogo completo, emparejando
siempre por código DALI: los datos DALI salen de ARTICULOS_DALI.xlsx y los SAP, el código de proveedor y el precio
se conservan del catálogo anterior (ver _regenerar_catalogo_dali_completo_impl).

13 nov 2026 (petición de Víctor, «reducir pasos y que todo esté actualizado en la web lo antes posible»): ya no hay
botón 2. Esto lo lanza el botón «Actualizar Catálogo Asignaciones» del final de la Acción 1
(app/ui/ventana_volcado_dali.py), que ya ha pedido la confirmación; por eso aquí no se pregunta otra vez.
_regenerar_catalogo_dali_completo envía después a la web solo lo que haya cambiado.
"""

import os
from tkinter import messagebox
import threading
from app.comun import volcado_dali
from app.comun.tarifa_core import (
    FOLDER_RECURSOS,
)


class AccionListadoDali:
    def _volcar_articulos_dali_al_catalogo(self):
        """
        Final de la Acción 1 — Actualizar el catálogo y Catálogo Asignaciones con los artículos DALI.
        Origen : RECURSOS/ARTICULOS_DALI.xlsx  (generado por la Acción 1 y, si se quiso, revisado en Excel)
        Destino: RECURSOS/catalogo-dali-completo.xlsx (por código DALI) y, después, Catálogo Asignaciones
        """
        carpeta = self._carpeta_proyecto.get().strip()
        if not carpeta:
            messagebox.showwarning("Sin proyecto",
                                   "Seleccione primero la carpeta del Proyecto.")
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
                "Ejecute la Acción 1 — Organizar artículos DALI\n"
                "para generar este archivo antes de continuar.")
            return

        # ── Leer fecha del encabezado (celda B3) ─────────────────
        try:
            texto_fecha = volcado_dali.fecha_encabezado(ruta_origen)
        except Exception as e:
            messagebox.showerror("Error al leer origen",
                                 f"No se pudo leer ARTICULOS_DALI.xlsx:\n{e}")
            return

        # ── Proceso en hilo ──────────────────────────────────────
        def proceso():
            self.consola.separador("INICIO ACTUALIZACIÓN DALI")
            self.log_consola("Iniciando actualización de los artículos DALI del catálogo...", "INFO")
            self.after(0, lambda: self._barra.indeterminate("Leyendo ARTICULOS_DALI..."))
            try:
                n_filas = volcado_dali.contar_articulos(ruta_origen)
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
                    f"\U0001f4c1  {ruta_catalogo}\n\n"
                    "Los cambios se envían ahora a Catálogo Asignaciones:\n"
                    "verás el resumen antes de aplicarlos."))

            except PermissionError:
                self.log_consola("ERROR: catalogo-dali-completo.xlsx está abierto en Excel.", "ERROR")
                self.consola.separador("FIN ACTUALIZACIÓN DALI (con error)")
                self.after(0, lambda: self._barra.reset("Error al guardar"))
                self.after(0, lambda: messagebox.showerror(
                    "Archivo bloqueado por Excel",
                    "❌  No se pudo guardar RECURSOS/catalogo-dali-completo.xlsx:\n"
                    "está abierto en Excel.\n\n"
                    "Ciérrelo y vuelva a pulsar la Acción 1: ofrecerá usar el\n"
                    "ARTICULOS_DALI.xlsx ya generado para actualizar."))
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

# -*- coding: utf-8 -*-
"""
accion_asignar_sap.py — Acción 8: «Asignar SAP pendientes» (v10.2.50).

Mixin de App (tarifa_app.py). Central Compras Princess Canarias · La idea y el desarrollo de la aplicación,
incluidas sus pruebas, son de Víctor Alejandro Martín Andrade (VAMA).

Petición de Víctor (13 nov 2026): ventana con todos los códigos DALI activos sin SAP del catálogo, para grabar los
SAP en el orden que se quiera (app/ui/ventana_asignar_sap.py; lógica en app/comun/asignar_sap.py).
- Se guardan en el catálogo (RECURSOS/catalogo-dali-completo.xlsx), con su bloqueo, resolviendo antes los códigos
  SAP repetidos con la misma ventana de siempre.
- Después se completan sus datos SAP con el mismo cruce de la Acción 3 (Productos.xlsx, Listas de productos.xlsx
  y ARTICULOS_DALI.xlsx), que además envía los cambios a Catálogo Asignaciones. Si falta alguno de esos archivos,
  solo se envían los códigos.
- Un código que no está en Productos.xlsx se guarda igualmente y se avisa: hay que descargar de nuevo
  Productos.xlsx y Listas de productos.xlsx en RECURSOS y pulsar otra vez la Acción 3.
"""
import os
import threading
from datetime import datetime
from tkinter import messagebox

from app.comun import asignar_sap
from app.comun import catalogo_local
from app.comun.sap_helpers import leer_productos_xlsx
from app.comun.tarifa_core import FOLDER_RECURSOS


class AccionAsignarSap:
    def _asignar_sap_pendientes(self):
        carpeta = self._carpeta_proyecto.get().strip()
        if not carpeta:
            messagebox.showwarning("Sin proyecto", "Seleccione primero la carpeta del Proyecto.")
            return
        ruta_cat = catalogo_local.ruta_catalogo(carpeta)
        if not os.path.isfile(ruta_cat):
            messagebox.showwarning(
                "Catálogo no encontrado",
                "No se encontró RECURSOS/catalogo-dali-completo.xlsx.\n\n"
                "Ejecute la Acción 1 (o descárguelo de Catálogo Asignaciones) y vuelva a intentarlo.")
            return
        self._pasar_listado_al_catalogo()
        self.consola.separador("ASIGNAR SAP PENDIENTES")
        self.log_consola("Buscando códigos DALI activos sin SAP en el catálogo...", "INFO")
        self._barra.indeterminate("Leyendo el catálogo...")

        def _cargar():
            self._traer_cambios_web("la Acción 8")   # (13 nov 2026) los SAP puestos en la web ya no salen pendientes
            try:
                lista = asignar_sap.pendientes(ruta_cat)
                ruta_prod = os.path.join(carpeta, FOLDER_RECURSOS, "Productos.xlsx")
                productos = leer_productos_xlsx(ruta_prod)[0] if os.path.isfile(ruta_prod) else None
                self.after(0, lambda: self._abrir_ventana_asignar_sap(ruta_cat, lista, productos))
            except Exception as e:  # noqa: BLE001
                self.log_consola(f"No se pudo leer el catálogo: {e}", "ERROR")
                self.after(0, lambda: self._barra.reset("Error al leer el catálogo"))
                self.after(0, lambda err=str(e): messagebox.showerror("Error al leer el catálogo", err))
        threading.Thread(target=_cargar, daemon=True, name="AsignarSapCargar").start()

    def _abrir_ventana_asignar_sap(self, ruta_cat, lista, productos):
        from app.ui.ventana_asignar_sap import VentanaAsignarSap
        self._barra.reset("")
        self.log_consola(f"  └ {len(lista)} código(s) DALI activos sin SAP.", "OK" if lista else "INFO")
        if productos is None:
            self.log_consola("  ⚠  Sin RECURSOS/Productos.xlsx: no se comprobará si los códigos existen en SAP.",
                             "WARN")
        if not lista:
            messagebox.showinfo("Asignar SAP pendientes", "✔  Todos los códigos DALI activos tienen código SAP.")
            return
        ven = VentanaAsignarSap(self, lista, productos,
                                lambda nuevos, resolver: self._guardar_sap_asignados(ruta_cat, lista, productos,
                                                                                     nuevos, resolver),
                                consola_fn=self.log_consola, listados_fn=self._avisar_listados_sap)
        self.wait_window(ven)
        self.log_consola(f"  └ Asignar SAP pendientes: {ven.guardados} código(s) guardados en esta sesión.", "OK")

    def _guardar_sap_asignados(self, ruta_cat, lista, productos, nuevos, resolver):
        """Se llama desde el hilo de guardado de la ventana. Devuelve el resumen, o None si se canceló."""
        descripciones = {p["dali"]: p["descripcion"] for p in lista}
        self._catalogo_lock.acquire()
        try:
            wb, ws, mapa = catalogo_local.abrir_para_editar(ruta_cat)
            try:
                fila_de, sap_de = catalogo_local.mapas_dali_sap(ws, mapa)
                for sap in set(nuevos.values()):
                    if sap in sap_de and "descripcion" in mapa:
                        dali_e, fila = sap_de[sap]
                        descripciones.setdefault(dali_e, ws.cell(row=fila, column=mapa["descripcion"]).value or "")
                conflictos = asignar_sap.calcular_conflictos(nuevos, sap_de, descripciones)
                a_borrar = set()
                if conflictos:
                    self.log_consola(f"  ⚠ {len(conflictos)} código(s) SAP repetidos: pidiendo cuál es el correcto...",
                                     "WARN")
                    a_borrar = resolver(conflictos)
                    if a_borrar is None:
                        self.log_consola("  ✘ Guardado cancelado (códigos repetidos sin resolver).", "WARN")
                        return None
                escritos, limpiados, no_encontrados = asignar_sap.aplicar(ws, mapa, fila_de, nuevos, conflictos,
                                                                          a_borrar)
                wb.save(ruta_cat)
            finally:
                wb.close()
        finally:
            self._catalogo_lock.release()

        for dali, sap in escritos.items():
            self.log_consola(f"  ✔ DALI {dali:06d} → SAP {sap}", "OK")
        fuera = asignar_sap.fuera_de_productos(escritos, productos)
        if fuera:
            self.log_consola(
                f"  ⚠ {len(fuera)} código(s) no están en Productos.xlsx (guardados igualmente): descarga de nuevo "
                "Productos.xlsx y Listas de productos.xlsx en RECURSOS y pulsa la Acción 3.", "WARN")
        completando = False
        if escritos or limpiados:
            completando = self._completar_datos_sap_tras_asignar(list(escritos) + list(limpiados))
        return {"escritos": escritos, "limpiados": limpiados, "no_encontrados": no_encontrados,
                "fuera_productos": fuera, "completando": completando and bool(set(escritos) - set(fuera))}

    def _completar_datos_sap_tras_asignar(self, cambiados):
        """Completa los datos SAP con el mismo cruce de la Acción 3 (en segundo plano), que también regenera el
        catálogo y envía los cambios a la web. Si falta algún archivo, solo se envían los códigos.
        Devuelve True si se ha lanzado el cruce."""
        carpeta = self._carpeta_proyecto.get().strip()
        rec = os.path.join(carpeta, FOLDER_RECURSOS)
        rutas = {n: os.path.join(rec, n) for n in ("Productos.xlsx", "ARTICULOS_DALI.xlsx", "Listas de productos.xlsx")}
        faltan = [n for n, r in rutas.items() if not os.path.isfile(r)]
        if faltan:
            self.log_consola(f"  ℹ Para completar los datos SAP falta(n) en RECURSOS: {', '.join(faltan)}. Se envían "
                             "solo los códigos a Catálogo Asignaciones.", "WARN")
            self._tras_cambiar_sap_en_catalogo(cambiados)
            return False

        def _cruce():
            try:
                self.log_consola("Completando los datos SAP de los códigos asignados (cruce de la Acción 3)...",
                                 "INFO")
                tabla_productos, filas_prod = self._leer_productos_sap(rutas["Productos.xlsx"])
                if tabla_productos is None:
                    self._tras_cambiar_sap_en_catalogo(cambiados)
                    return
                tabla_dali = self._leer_articulos_dali_sap(rutas["ARTICULOS_DALI.xlsx"])
                tabla_categoria, conflictos_cat = self._leer_categoria_valoracion(rutas["Listas de productos.xlsx"],
                                                                                    tabla_productos)
                dt = datetime.fromtimestamp(os.path.getmtime(rutas["Productos.xlsx"]))
                self._cruzar_y_escribir_listado_sap(
                    catalogo_local.ruta_catalogo(carpeta), tabla_productos, tabla_dali, tabla_categoria,
                    conflictos_cat, filas_prod, dt.strftime("%d/%m/%Y"), dt.strftime("%H:%M"))
            except Exception as e:  # noqa: BLE001 — en un hilo: que nunca se pierda en silencio
                self.log_consola(f"  ⚠  No se pudieron completar los datos SAP: {e}. Pulse la Acción 3.", "WARN")
                self._tras_cambiar_sap_en_catalogo(cambiados)
        threading.Thread(target=_cruce, daemon=True, name="AsignarSapCruce").start()
        return True

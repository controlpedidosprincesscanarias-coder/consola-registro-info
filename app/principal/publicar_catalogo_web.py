# -*- coding: utf-8 -*-
"""
publicar_catalogo_web.py — después de las Acciones 2, 3 y 4 (y de resolver un SAP duplicado), envía a
Catálogo Asignaciones solo lo que ha cambiado (etapa 6, fase 1, v10.2.50).

Mixin de App (tarifa_app.py). Central Compras Princess Canarias · La idea y el desarrollo de la aplicación,
incluidas sus pruebas, son de Víctor Alejandro Martín Andrade (VAMA).

- Corre en su propio hilo: la acción que lo lanza termina como siempre y no espera a la web.
- Si ya hay un envío en marcha, no se lanza otro: se apunta y, al terminar, se repite una vez (así dos
  tarifas seguidas no abren dos ventanas).
- La lógica está en app/comun/catalogo_web.py (con pruebas); aquí solo van la ventana, los mensajes y el hilo.
- Fase 1 (doble escritura): el catálogo local (RECURSOS/catalogo-dali-completo.xlsx) se sigue escribiendo
  igual; lo que queda distinto tras el envío se cuenta en la consola, para ver que la web y el Excel cuadran.
- (13 nov 2026, petición de Víctor) Antes de cada envío, y al empezar las Acciones 3, 4 y 8
  (`_traer_cambios_web`), lo cambiado en la web (datos SAP, precio y código de proveedor) baja al Excel; en cada
  campo gana la última modificación (catalogo_web.fusionar). Ya no hay casilla para «sustituir lo editado en la
  web»: lo decide la fecha.
"""
import os
import threading
from contextlib import nullcontext

from app.comun import catalogo_local
from app.comun import catalogo_web as W
from app.comun.credenciales import borrar_credenciales_dali
from app.comun.tarifa_core import FOLDER_RECURSOS, guardar_config
from app.ui import ventana_publicar_catalogo


class PublicarCatalogoWeb:
    def _publicar_catalogo_web(self, motivo, archivo=None):
        """Lanza el envío a la web en segundo plano (o lo deja apuntado si ya hay uno en marcha)."""
        with self._publicacion_web_estado:
            if self._publicacion_web_activa:
                self._publicacion_web_pendiente = (motivo, archivo)
                return
            self._publicacion_web_activa = True
        # (13 nov 2026) La barra no se pone en verde hasta que acaba el envío a la web
        self._barra_ocupar("Enviando cambios a Catálogo Asignaciones…")
        threading.Thread(target=self._publicar_catalogo_web_hilo, args=(motivo, archivo), daemon=True).start()

    def _barra_ocupar(self, texto):
        barra = getattr(self, "_barra", None)
        if barra is not None and hasattr(barra, "ocupar"):
            try:
                self.after(0, lambda: barra.ocupar(texto))
            except Exception:  # noqa: BLE001 — sin ventana (pruebas): nada que mover
                pass

    def _barra_liberar(self):
        barra = getattr(self, "_barra", None)
        if barra is not None and hasattr(barra, "liberar"):
            try:
                self.after(0, barra.liberar)
            except Exception:  # noqa: BLE001
                pass

    def _publicar_catalogo_web_hilo(self, motivo, archivo):
        try:
            while True:
                try:
                    self._publicar_catalogo_web_una_vez(motivo, archivo)
                except Exception as e:  # noqa: BLE001 — en un hilo: que nunca se pierda en silencio
                    self.log_consola(f"Catálogo Asignaciones: error inesperado al enviar los cambios ({e}). "
                                     "Volverán a salir la próxima vez.", "ERROR")
                with self._publicacion_web_estado:
                    siguiente = self._publicacion_web_pendiente
                    self._publicacion_web_pendiente = None
                    if siguiente is None:
                        self._publicacion_web_activa = False
                        self._barra_liberar()
                        return
                motivo, archivo = siguiente
        except BaseException:
            with self._publicacion_web_estado:
                self._publicacion_web_activa = False
            self._barra_liberar()
            raise

    def _lock_cache_web(self):
        lock = getattr(self, "_cache_web_lock", None)
        return lock if lock is not None else nullcontext()

    def _conectar_catalogo_web(self):
        """Inicia sesión en Catálogo Asignaciones. Devuelve el cliente o None (ya explicado en la consola)."""
        log = self.log_consola
        email, password = self._ejecutar_en_hilo_principal_y_esperar(self._obtener_credenciales_dali_app)
        if not email or not password:
            log("Catálogo Asignaciones: sin credenciales, no se han enviado los cambios a la web "
                "(saldrán la próxima vez).", "WARN")
            return None
        cliente = W.ClienteCatalogo()
        try:
            cliente.login(email, password)
        except W.ErrorWeb as e:
            if e.codigo in (401, 403):
                borrar_credenciales_dali(self.config_data)
                guardar_config(self.config_data)
                log(f"Catálogo Asignaciones: inicio de sesión rechazado ({e}). Se han borrado las credenciales "
                    "guardadas; se pedirán otra vez.", "ERROR")
            else:
                log(f"Catálogo Asignaciones: {e}. Los cambios saldrán la próxima vez.", "WARN")
            return None
        return cliente

    def _bajar_de_la_web(self, cache, ruta_local):
        """Baja al Excel lo que gana la web (catalogo_web.fusionar) y apunta como base lo que ya coincide.
        Devuelve (catálogo local tal como queda, fusión); en la fusión, "bajar" se queda solo con lo que NO se pudo
        escribir (Excel abierto). Hay que llamarlo con la copia de la web ya sincronizada."""
        log = self.log_consola
        with self._catalogo_lock:                 # que ninguna acción lo esté escribiendo a la vez
            local = W.leer_catalogo_local(ruta_local)
            hora_excel = W.hora_iso(os.path.getmtime(ruta_local))
            fus = W.fusionar(local, cache, hora_excel)
            if fus["bajar"]:
                n_campos = sum(len(c) for c in fus["bajar"].values())
                try:
                    catalogo_local.escribir_campos(ruta_local, W.bajar_para_excel(fus["bajar"]))
                    log(f"  ⬇ Traído de Catálogo Asignaciones al catálogo: {n_campos} dato(s) de "
                        f"{len(fus['bajar'])} artículo(s) cambiados en la web.", "OK")
                    for cod, campos in list(fus["bajar"].items())[:20]:
                        detalle = ", ".join(f"{W.CLAVE_LOCAL.get(c, c)}={v}" for c, v in campos.items())
                        log(f"     DALI {cod}: {detalle}", "INFO")
                    if len(fus["bajar"]) > 20:
                        log(f"     … y {len(fus['bajar']) - 20} artículo(s) más.", "INFO")
                    local = W.leer_catalogo_local(ruta_local)
                    fus["bajar"] = {}             # ya está en el Excel
                except PermissionError:
                    log("  ⚠ No se ha podido traer lo cambiado en la web: catalogo-dali-completo.xlsx está abierto "
                        "en Excel. Ciérralo; se traerá la próxima vez (mientras, no se envía a la web lo que "
                        "cambió allí).", "WARN")
            for cod, campo, vl, vw, gana in fus["conflictos"][:20]:
                log(f"     ⚖ DALI {cod} · {campo}: cambiado en la web («{vw}») y en la Consola («{vl}»); "
                    f"gana {'la web' if gana == 'web' else 'la Consola'} (cambio más reciente).", "WARN")
        W.acordar(cache, local)
        return local, fus

    def _traer_cambios_web(self, motivo):
        """(13 nov 2026) Al empezar las acciones que escriben datos SAP, precio o código de proveedor (3, 4 y 8): baja
        al Excel lo cambiado en la web, para trabajar ya con ello. Síncrono (se llama desde el hilo de la acción).
        Sin conexión o sin credenciales, la acción sigue igual: el envío de después tampoco pisa lo de la web."""
        carpeta = self._carpeta_proyecto.get().strip()
        ruta_local = os.path.join(carpeta, FOLDER_RECURSOS, "catalogo-dali-completo.xlsx") if carpeta else ""
        if not ruta_local or not os.path.isfile(ruta_local):
            return
        try:
            with self._lock_cache_web():
                self.log_consola(f"Catálogo Asignaciones: trayendo lo cambiado en la web antes de {motivo}...", "INFO")
                cliente = self._conectar_catalogo_web()
                if cliente is None:
                    return
                cache = W.Cache().cargar()
                cache.sincronizar(cliente)
                self._bajar_de_la_web(cache, ruta_local)
                cache.guardar()
        except W.ErrorWeb as e:
            self.log_consola(f"Catálogo Asignaciones: {e}. Se sigue sin traer lo de la web (se traerá la próxima "
                             "vez).", "WARN")
        except Exception as e:  # noqa: BLE001 — que la acción siga aunque falle esto
            self.log_consola(f"Catálogo Asignaciones: no se ha podido traer lo cambiado en la web ({e}).", "WARN")

    def _publicar_catalogo_web_una_vez(self, motivo, archivo):
        with self._lock_cache_web():
            self._publicar_catalogo_web_una_vez_impl(motivo, archivo)

    def _publicar_catalogo_web_una_vez_impl(self, motivo, archivo):
        log = self.log_consola
        carpeta = self._carpeta_proyecto.get().strip()
        ruta_local = os.path.join(carpeta, FOLDER_RECURSOS, "catalogo-dali-completo.xlsx") if carpeta else ""
        if not ruta_local or not os.path.isfile(ruta_local):
            return

        cliente = self._conectar_catalogo_web()
        if cliente is None:
            return
        log("Catálogo Asignaciones: comprobando qué ha cambiado...", "INFO")

        try:
            cache = W.Cache().cargar()
            primera_vez = cache.hasta is None
            if primera_vez:
                log("  Primera vez en este equipo: se descarga una copia del catálogo de la web "
                    "(solo esta vez; después, solo lo que cambie).", "INFO")
            n = cache.sincronizar(cliente)
            if not primera_vez and n:
                log(f"  {n} artículo(s) cambiado(s) en la web desde la última vez.", "INFO")

            # (13 nov 2026) Lo que gana la web baja antes al Excel: lo que queda distinto es lo que sube
            local, fus = self._bajar_de_la_web(cache, ruta_local)
            dif = W.calcular_diferencias(local, cache.articulos, cache.publicado_hasta, cache.protegidos)
            W.quitar_campos(dif, fus["bajar"])  # no bajó (Excel abierto): tampoco se sube, para no pisar la web
            if not any(dif.values()):
                cache.guardar()
                log("  └ La web ya coincide con el catálogo local: no hay nada que enviar.", "OK")
                return

            simulaciones = {}
            for grupo in W.GRUPOS:
                filas = W.filas_para_enviar(dif, grupo, incluir_editados=True)
                if filas:
                    simulaciones[grupo] = W.enviar_grupo(cliente, grupo, filas, simular=True, archivo=archivo,
                                                         sincronizado_desde=cache.publicado_hasta)

            aplicar, _ = self._ejecutar_en_hilo_principal_y_esperar(
                lambda: ventana_publicar_catalogo.pedir_confirmacion(self, dif, simulaciones, motivo))
            incluir_editados = True       # lo que ganaba la web ya ha bajado al Excel: lo que queda es de la Consola
            if not aplicar:
                cache.guardar()
                log("  └ No se ha enviado nada a la web («Ahora no»): los cambios saldrán la próxima vez.", "WARN")
                return

            for grupo in W.GRUPOS:                    # en orden: altas (accion2), SAP (accion3), precios (accion4)
                filas = W.filas_para_enviar(dif, grupo, incluir_editados)
                if not filas:
                    continue
                r = W.enviar_grupo(cliente, grupo, filas, simular=False, archivo=archivo,
                                   sincronizado_desde=cache.publicado_hasta)
                res = r["resumen"]
                log(f"  {W.NOMBRE_GRUPO[grupo]}: {res['nuevos']} alta(s), {res['cambiados']} cambiado(s), "
                    f"{res['irregulares']} sin aplicar.", "OK" if not res["irregulares"] else "WARN")
                nuevos = {d["codigo_dali"] for d in dif["accion2"] if d["nuevo"]}
                for irr in r["irregulares"]:
                    if irr.get("tipo") == "no_existe" and irr.get("codigo_dali") in nuevos:
                        continue
                    log(f"    ⚠ DALI {irr.get('codigo_dali')}: {irr.get('detalle') or irr.get('tipo')} "
                        "— corrígelo para que la web y el catálogo local coincidan.", "WARN")

            W.actualizar_protegidos(cache, dif, incluir_editados)
            cache.sincronizar(cliente)                # trae lo recién aplicado
            cache.publicado_hasta = cache.hasta
            restantes = W.calcular_diferencias(local, cache.articulos, cache.publicado_hasta, cache.protegidos)
            W.actualizar_protegidos(cache, restantes, incluir_editados=True)
            W.acordar(cache, local)
            cache.guardar()
            pendientes = {g: len(v) for g, v in restantes.items() if v}
            if pendientes:
                detalle = ", ".join(f"{W.NOMBRE_GRUPO[g]}: {n}" for g, n in pendientes.items())
                log(f"  └ Catálogo Asignaciones actualizado. Siguen distintos de la web: {detalle} "
                    "(irregulares que la web no aplica: revísalos).", "WARN")
            else:
                log("  └ Catálogo Asignaciones actualizado: la web y el catálogo local coinciden.", "OK")
        except W.ErrorWeb as e:
            log(f"Catálogo Asignaciones: {e}. Los cambios que falten saldrán la próxima vez.",
                "ERROR" if e.codigo == 403 else "WARN")

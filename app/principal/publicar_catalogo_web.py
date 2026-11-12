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
"""
import os
import threading

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
        threading.Thread(target=self._publicar_catalogo_web_hilo, args=(motivo, archivo), daemon=True).start()

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
                        return
                motivo, archivo = siguiente
        except BaseException:
            with self._publicacion_web_estado:
                self._publicacion_web_activa = False
            raise

    def _publicar_catalogo_web_una_vez(self, motivo, archivo):
        log = self.log_consola
        carpeta = self._carpeta_proyecto.get().strip()
        ruta_local = os.path.join(carpeta, FOLDER_RECURSOS, "catalogo-dali-completo.xlsx") if carpeta else ""
        if not ruta_local or not os.path.isfile(ruta_local):
            return

        email, password = self._ejecutar_en_hilo_principal_y_esperar(self._obtener_credenciales_dali_app)
        if not email or not password:
            log("Catálogo Asignaciones: sin credenciales, no se han enviado los cambios a la web "
                "(saldrán la próxima vez).", "WARN")
            return

        log("Catálogo Asignaciones: comprobando qué ha cambiado...", "INFO")
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
            return

        try:
            cache = W.Cache().cargar()
            primera_vez = cache.hasta is None
            if primera_vez:
                log("  Primera vez en este equipo: se descarga una copia del catálogo de la web "
                    "(solo esta vez; después, solo lo que cambie).", "INFO")
            n = cache.sincronizar(cliente)
            if not primera_vez and n:
                log(f"  {n} artículo(s) cambiado(s) en la web desde la última vez.", "INFO")

            with self._catalogo_lock:                 # que ninguna acción lo esté escribiendo a la vez
                local = W.leer_catalogo_local(ruta_local)
            dif = W.calcular_diferencias(local, cache.articulos, cache.publicado_hasta, cache.protegidos)
            if not any(dif.values()):
                log("  └ La web ya coincide con el catálogo local: no hay nada que enviar.", "OK")
                return

            simulaciones = {}
            for grupo in W.GRUPOS:
                filas = W.filas_para_enviar(dif, grupo, incluir_editados=True)
                if filas:
                    simulaciones[grupo] = W.enviar_grupo(cliente, grupo, filas, simular=True, archivo=archivo,
                                                         sincronizado_desde=cache.publicado_hasta)

            aplicar, incluir_editados = self._ejecutar_en_hilo_principal_y_esperar(
                lambda: ventana_publicar_catalogo.pedir_confirmacion(self, dif, simulaciones, motivo))
            if not aplicar:
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
            W.actualizar_protegidos(cache, restantes, incluir_editados=False)
            cache.guardar()
            pendientes = {g: len(v) for g, v in restantes.items() if v}
            if pendientes:
                detalle = ", ".join(f"{W.NOMBRE_GRUPO[g]}: {n}" for g, n in pendientes.items())
                log(f"  └ Catálogo Asignaciones actualizado. Siguen distintos de la web: {detalle} "
                    "(editados en la web que dejaste fuera o irregulares).", "WARN")
            else:
                log("  └ Catálogo Asignaciones actualizado: la web y el catálogo local coinciden.", "OK")
        except W.ErrorWeb as e:
            log(f"Catálogo Asignaciones: {e}. Los cambios que falten saldrán la próxima vez.",
                "ERROR" if e.codigo == 403 else "WARN")

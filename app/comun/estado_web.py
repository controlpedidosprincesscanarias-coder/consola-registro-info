# -*- coding: utf-8 -*-
"""
estado_web.py — estado de la conexión con Catálogo Asignaciones (18 nov 2026, petición de Víctor: «¿se puede lanzar
un aviso en la consola cuando no se pueda acceder a Catálogo Asignaciones?»).

Central Compras Princess Canarias · La idea y el desarrollo de la aplicación, incluidas sus pruebas,
son de Víctor Alejandro Martín Andrade (VAMA).

Sin tkinter (pruebas en tests/test_estado_web.py). La ventana está en app/ui/ventana_sin_conexion.py y el enganche
con las acciones en app/principal/publicar_catalogo_web.py.
- `clasificar(error)`: qué tipo de fallo es (sin red, la web arrancando, error de la web, credenciales, permiso…) y
  cómo explicarlo.
- `EstadoConexion`: lo que enseña el indicador de la cabecera y si toca sacar la ventana. La ventana sale la primera
  vez que falla; mientras siga sin conexión no se repite en cada acción (solo el indicador y el log), salvo que el
  usuario pulse «Reintentar». Al volver a conectar, se reinicia.
"""
import threading
import time
from datetime import datetime

# Fallos que no son de conexión: no tiene sentido ofrecer «Reintentar» (el log ya explica qué hacer).
TIPOS_SIN_REINTENTO = ("credenciales", "permiso")

_SIN_RED = ("getaddrinfo", "name or service", "nodename", "no such host", "network is unreachable",
            "no route to host", "connection refused", "connection reset", "winerror 10051", "winerror 10061",
            "winerror 10065", "winerror 11001", "winerror 11004", "errno 11001", "temporary failure in name")
_TARDA = ("timed out", "timeout", "winerror 10060", "tiempo de espera")


def clasificar(error):
    """{"tipo", "titulo", "explicacion"} de un error de la conexión (catalogo_web.ErrorWeb u otra excepción)."""
    codigo = getattr(error, "codigo", None)
    texto = str(error).lower()
    if codigo in (401, 428):
        return {"tipo": "credenciales", "titulo": "Catálogo Asignaciones ha rechazado el inicio de sesión",
                "explicacion": "El usuario o la contraseña ya no valen, o la web pide un código de verificación. "
                               "Vuelve a iniciar sesión."}
    if codigo == 403:
        return {"tipo": "permiso", "titulo": "Sin permiso en Catálogo Asignaciones",
                "explicacion": "Tu cuenta no tiene el permiso «Escribir desde la Consola Registro Info». "
                               "Pídeselo al administrador principal."}
    if codigo is not None and codigo >= 500:
        return {"tipo": "servidor", "titulo": "Catálogo Asignaciones tiene un problema",
                "explicacion": f"La web ha respondido con un error (HTTP {codigo}). Suele ser pasajero: prueba de "
                               "nuevo en unos minutos."}
    if any(s in texto for s in _TARDA):
        return {"tipo": "arrancando", "titulo": "Catálogo Asignaciones tarda en responder",
                "explicacion": "Si llevaba un rato sin usarse, la web está arrancando y puede tardar hasta un "
                               "minuto. Espera un momento y pulsa «Reintentar»."}
    if any(s in texto for s in _SIN_RED):
        return {"tipo": "sin_red", "titulo": "Sin conexión con Catálogo Asignaciones",
                "explicacion": "No hay conexión a internet o la red no deja llegar a la web. Comprueba la red "
                               "del equipo."}
    return {"tipo": "otro", "titulo": "No se ha podido acceder a Catálogo Asignaciones",
            "explicacion": "La conexión con la web ha fallado."}


class EstadoConexion:
    """Estado de la conexión, compartido por los hilos de las acciones (con su propio candado)."""

    def __init__(self, reloj=time.time):
        self._reloj = reloj
        self._lock = threading.Lock()
        self.conectado = None          # None: aún no se ha comprobado; True / False
        self.desde = None              # hora (segundos) del primer fallo seguido
        self.error = None
        self._avisado = False

    def marcar_ok(self):
        """Ha conectado. Devuelve True si ha cambiado (hay que repintar el indicador)."""
        with self._lock:
            cambia = self.conectado is not True
            self.conectado, self.desde, self.error, self._avisado = True, None, None, False
            return cambia

    def marcar_fallo(self, error):
        """Ha fallado. Devuelve True si hay que sacar la ventana (primer fallo desde que conectó o desde el último
        «Reintentar»)."""
        with self._lock:
            if self.conectado is not False:
                self.desde = self._reloj()
            self.conectado, self.error = False, error
            avisar = not self._avisado
            self._avisado = True
            return avisar

    def permitir_nuevo_aviso(self):
        """Tras «Reintentar»: si vuelve a fallar, la ventana sale otra vez."""
        with self._lock:
            self._avisado = False

    def indicador(self):
        """(texto, estado) para la cabecera; estado: "pendiente", "ok" o "fallo"."""
        with self._lock:
            if self.conectado is None:
                return "● Catálogo Asignaciones: sin comprobar", "pendiente"
            if self.conectado:
                return "● Catálogo Asignaciones: conectado", "ok"
            hora = datetime.fromtimestamp(self.desde).strftime("%H:%M") if self.desde else "?"
            return f"● Catálogo Asignaciones: sin conexión desde las {hora}", "fallo"

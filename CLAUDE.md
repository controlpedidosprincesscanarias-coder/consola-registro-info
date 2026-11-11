# CLAUDE.md — Consola Registro Info SAP (TARIFA Finder)

> **Autoría:** la idea y el desarrollo de la aplicación, incluidas sus pruebas, son de **Víctor Alejandro Martín Andrade**, para el Departamento de Central de Compras de Princess Hotels en Canarias. Ver `AUTORIA.md`.

Instrucciones para Claude Code en este repositorio. Responde siempre en **español**.

## Qué es
App de escritorio para Windows (Python + Tkinter) de la Central de Compras: prepara las tarifas de los proveedores y genera los **Registros de Información de compras de SAP (RINFO)** a partir de la tarifa y del «Purchasing Info Records» descargado de SAP. Mantiene también los listados de artículos DALI y de materiales SAP. Ventana «CREACIÓN REGISTRO INFO SAP»; ejecutable `TARIFA_Finder.exe`.
- **Compilación en GitHub** (desde nov 2026, como el Organizador): `.github/workflows/compilar.yml`, **solo a mano** desde la pestaña Actions («Compilar Consola Registro Info»). Corre en Windows y hace esto:
  1. Calcula la versión.
  2. Pasa las pruebas.
  3. Compila con `build_exe.py`, que llama a PyInstaller con `--onefile --windowed` y genera `dist/TARIFA_Finder.exe` con su `.sha256`.

  Tiene dos destinos:
  - «solo-descarga»: el `.exe` queda como descarga del propio workflow.
  - «borrador»: además crea una Release en **borrador**, con las novedades generadas a partir de los PR. Se publica a mano con «Publish release».

  Nunca publicar una Release directamente desde el workflow ni lanzarlo al subir código. Dependencias de la compilación: `compilacion/requisitos.txt`.
- **Versión automática** (`compilacion/version.py`, pruebas en `tests/test_version.py`):
  - La versión de cada compilación es la mayor entre `APP_VERSION` de `tarifa_app.py` y las Releases ya creadas (también los borradores), **+2**, siempre par: 10.2.36 → 10.2.38.
  - Se escribe en la copia que se compila, así que el `.exe` la muestra y la Release lleva esa etiqueta (`v10.2.38`). El workflow **no** sube ese cambio al repositorio: las versiones publicadas quedan en las etiquetas de las Releases.
  - El campo «version» del workflow permite forzar otra (por ejemplo, 10.3.0); tiene que ser mayor y par.
  - `APP_VERSION` (formato `"v X.Y.Z"`) sigue en `tarifa_app.py`, que es de donde lo leen el workflow, `_bat_helper.py` y `build_exe.py`.
- **Compilación local** (en Windows, si hace falta): `COMPILAR_TARIFA.bat` → `build_exe.py`. Ya no hace falta para poner la versión. El `.spec` lo genera PyInstaller y no se sube.
- **Actualizaciones automáticas** a los compañeros: pendiente, igual que el Organizador desde `releases/latest`. Hará falta decidir si el repositorio vuelve a ser público, porque sin iniciar sesión solo se leen las Releases de un repositorio público.
- **Datos:** carpeta de proyecto en la unidad de red (`TARIFAS/`, `RECURSOS/`), elegida en la app y guardada en `tarifa_config.json`, junto al `.exe` o al script. Ese archivo **no se sube nunca** (está en `.gitignore`).
- **Catálogo Asignaciones** (repo `dali-sap-articulos-app`): antes de algunas acciones se descarga su catálogo a través del Worker `DALI_APP_API_BASE` (`app/principal/catalogo_dali.py`), con el usuario y la contraseña de cada persona.

## Estructura
Desde la etapa 2 (nov 2026) el código está en el paquete `app/`, y `tarifa_app.py` sigue en la raíz como punto de arranque: lo usan `COMPILAR_TARIFA.bat`, `_bat_helper.py` (lee `APP_VERSION`) y `build_exe.py`. Cada módulo conserva su nombre de siempre; para encontrar una función: `grep -rn "def nombre" tarifa_app.py app`.
- `tarifa_app.py` (~200 líneas) — punto de arranque y clase `App(InterfazApp, AccionOrganizarDali, AccionListadoDali, AccionListadoSap, CatalogoDali, AccionesTarifa, tk.Tk)`. Aquí quedan `__init__`, la cabecera y el pie, porque usan `APP_VERSION`, que **tiene que seguir en este archivo** (lo leen `_bat_helper.py` y `COMPILAR_TARIFA.bat`).
- `app/principal/` — los métodos de `App`, repartidos por temas en clases *mixin* (etapa 4, nov 2026, sin cambiar el código). Se llaman siempre como `self.metodo()`, así que un método puede vivir en cualquiera de ellas:
  - `interfaz.py`: construcción de la ventana, consola de log, botones y acciones sencillas (carpeta, buscar `TARIFA.xls`, abrir resultado). El botón 7 (tarifa vacía) está en construcción.
  - `accion_organizar_dali.py`: Acción 1. Su lógica está en `app/acciones/organizar_dali.py`.
  - `accion_listado_dali.py`: Acción 2.
  - `accion_listado_sap.py`: Acción 3.
  - `accion_tarifa.py`: Acciones 4 y 6 (TARIFA final, archivos INC / [INCOMPLETO], reinscripción SAP, botón Registro Info).
  - `catalogo_dali.py`: Catálogo DALI completo y Catálogo Asignaciones (`DALI_APP_API_BASE`).
  - Un método nuevo va al mixin de su tema. Dos mixins no pueden tener un método con el mismo nombre, porque ganaría el primero de la lista sin avisar.
- `app/comun/` — `tarifa_constants.py` (colores, fuentes, tablas fijas, `resource_path` y `RAIZ_PROYECTO`), `tarifa_core.py` (configuración, consola de log, barra de progreso), `sap_helpers.py` (funciones de negocio puras, sin tkinter) y `credenciales.py` (DPAPI, sin tkinter).
- `app/ui/tarifa_dialogs.py` — ventanas modales.
- `app/rinfo/` — motor de RINFO, repartido en la etapa 4 (nov 2026) sin cambiar el código:
  - `comun_rinfo.py`: constantes, `_log`, `_safe_after` y funciones auxiliares (Excel, similitud de nombres, búsqueda de tarifas, Purchasing y listados). No importa de otros módulos de `app.rinfo`.
  - `ventanas_rinfo.py`: ventanas que no dependen de otras (fechas, centros, teléfono, confirmaciones, materiales sin tarifa / para reactivar…). Importa solo de `comun_rinfo`.
  - `creador_rinfo.py` (`CreadorRINFO`, RINFO nuevo) y `generador_existente.py` (`GeneradorDesdeExistente`, desde un Purchasing existente). Importan de `comun_rinfo` y de `ventanas_rinfo`.
  - `crear_registro_info.py`: las ventanas del flujo (selección de tarifa, Purchasing encontrado, opciones sin Purchasing), que unen el Creador y el Generador, e `iniciar_crear_registro_info`. **Reexporta todos los nombres**, así que `from app.rinfo.crear_registro_info import X` sigue valiendo.
  - En las pruebas, una función como `_log` se parchea con `_parchear_rinfo()`, que la cambia en todos los módulos de `app.rinfo` (cada uno tiene su propia referencia).
  - El antiguo autotest `--test` está en `tests/test_casos_limite_pir.py`.
- `app/tarifas/transformar_tarifa.py` — transforma la tarifa del proveedor al formato de trabajo.
- `app/acciones/organizar_dali.py` — lógica de la Acción 1, sin tkinter.
- `assets/` (en la raíz): `logo.png`, `logo.ico`, `logo_dali_articulos.png`, `logo_sap.png` y `marcar_borrado_sap.png`. Todos se usan y van dentro del `.exe` (`ASSETS_NECESARIOS` en `build_exe.py`). Las imágenes van en `assets/` y se cargan con `resource_path()`; no se incrustan en base64 en el código.
- `tests/`.
- **Arranque:** lo pesado (pandas) se importa dentro de la acción que lo usa, no al arrancar. No volver a importarlo arriba del todo en `tarifa_app.py`.
- **Dependencias** (sin importaciones circulares, siempre con la ruta completa `from app.comun.tarifa_core import …`): `app.comun` no importa de otros paquetes del proyecto; `app.ui`, `app.rinfo` y `app.tarifas` importan de `app.comun`; `app.principal` importa de los anteriores, y `tarifa_app.py` de `app.principal`. Los módulos sin tkinter (`sap_helpers`, `credenciales`, `app/acciones`) no importan de la interfaz. Nada importa de `tarifa_app.py`.
- **Trampa conocida:** no usar `__file__` para localizar archivos del proyecto (assets, `tarifa_config.json`). Usar `resource_path()` y `RAIZ_PROYECTO`, porque los módulos ya no están en la raíz.
- **Saltos de línea:** `.gitattributes` guarda LF en el repositorio y Git pone CRLF al sacar los archivos en Windows; los `.bat` siempre con CRLF. Al editar desde Linux, no convertir archivos enteros.

## Seguridad
- **Nunca** escribir contraseñas, tokens, claves ni correos de cuentas en el código, en los commits ni en los documentos. El 9 nov 2026 se subió por error `tarifa_config.json` con la contraseña de Catálogo Asignaciones en claro; desde entonces está en `.gitignore` y la contraseña va cifrada (`credenciales.py`).
- La contraseña de Catálogo Asignaciones solo se guarda cifrada con DPAPI. Si no se puede cifrar, no se guarda. No reintroducir el guardado en claro.
- Datos reales (tarifas, Purchasing, listados `.xls`/`.xlsx`, logs) **nunca** al repositorio.

## Reglas de trabajo
1. Leer el código real antes de proponer cambios y buscar todos los usos de una función antes de modificarla.
2. Trabajar en rama y con PR; **no hacer push a `main` sin confirmación de Víctor**. Commits pequeños con mensaje claro.
3. Los errores que se encuentren se reparan en un **commit propio**, aunque ya estuvieran, sin mezclarlos con código que solo se mueve.
4. **`HISTORIAL_REORGANIZACION_RINFO.md` es el historial de esta app**: se le añade una sección numerada por cada tanda de trabajo, y se mantiene al día «Próximos pasos». Antes de empezar, leer sus secciones «Patrón de trabajo» y «Próximos pasos».
5. Al trocear un método largo, auditar **todas** las variables del bloque extraído (si se usan después). Ya hubo dos fallos así. Muchos procesos corren en hilos sin captura: si la app «se queda colgada», sospechar primero de una excepción silenciosa en el hilo.
6. Sin Windows ni datos reales no se prueba de verdad: compilar, pasar las pruebas y probar la lógica aislada (`object.__new__(Clase)`, ver el historial). Después, decir claramente qué queda por probar con datos reales en Windows.
7. Conservar siempre la línea de autoría en documentos y cabeceras, y añadirla en los nuevos.

## Pruebas
```bash
python -m unittest discover -s tests -v
```
- 112 pruebas (desde la raíz). 18 de ellas importan `crear_registro_info` (que usa tkinter): sin tkinter se saltan. En Linux hay que lanzarlas con un Python que tenga tkinter y con `xvfb-run`.
- Dependencias para probar: `pandas openpyxl xlrd pillow psutil`. Para compilar, además `pyinstaller pywin32`.
- Código muerto: `vulture tarifa_app.py app --min-confidence 60`, comprobando cada resultado con `grep` (también en cadenas) antes de quitar nada.
- `tests/test_crear_registro_info.py` prueba el código real a través de `sap_helpers.py` y de `crear_registro_info`; no se deben reimplementar funciones dentro de la prueba.
- Comprobación de la compilación: `build_exe.py` también funciona en Linux (genera un ejecutable de Linux). Sirve para ver que PyInstaller incluye todo `app/` y que la app arranca.

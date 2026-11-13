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
- **Actualizaciones automáticas** (desde nov 2026; lógica en `app/comun/actualizaciones.py`, ventana en `app/ui/ventana_actualizacion.py`, pruebas en `tests/test_actualizaciones.py`). Cada compañero tiene su propia copia del `.exe` en su PC.
  1. Al arrancar, en segundo plano y sin iniciar sesión, se consulta `releases/latest`. Los borradores no salen ahí, así que una versión no llega a nadie hasta que se pulsa «Publish release».
  2. Si hay una más nueva, se avisa con las novedades. «Actualizar ahora» descarga `TARIFA_Finder.exe` a la **misma carpeta** (`.nuevo.exe`) y comprueba el SHA-256 con el `.sha256` publicado: sin `.sha256` o con hash distinto no se instala.
  3. Después renombra el `.exe` en uso a `.anterior.exe`, coloca el nuevo, lo abre con el entorno de PyInstaller limpio (`entorno_limpio`) y cierra el viejo.
  4. Al arrancar, la versión nueva borra el `.anterior.exe` y lo dice en la consola. Si aún está en uso, lo reintenta cada 10 s, hasta 6 veces.

  Un 404 (ninguna Release publicada) no es un error. Desde el código (`python tarifa_app.py`) nunca se actualiza. **El repositorio tiene que seguir siendo público** y no se borran ni se renombran Releases, igual que en el Organizador.
- **Datos:** carpeta de proyecto en la unidad de red (`TARIFAS/`, `RECURSOS/`), elegida en la app y guardada en `tarifa_config.json`, junto al `.exe` o al script. Ese archivo **no se sube nunca** (está en `.gitignore`).
- **Catálogo Asignaciones** (repo `dali-sap-articulos-app`): antes de algunas acciones se descarga su catálogo a través del Worker `DALI_APP_API_BASE` (`app/principal/catalogo_dali.py`), con el usuario y la contraseña de cada persona.
- **Acción 8: Asignar SAP pendientes** (v10.2.50, petición de Víctor, 13 nov 2026): `AccionAsignarSap` (`app/principal/accion_asignar_sap.py`) abre `app/ui/ventana_asignar_sap.py` con todos los DALI **activos sin SAP** del catálogo (lógica sin tkinter en `app/comun/asignar_sap.py`: `pendientes`, `filtrar`, `validar_codigo`, `calcular_conflictos`, `aplicar`, `fuera_de_productos`). Tabla (Treeview) con buscador y filtros; Intro asigna y pasa al siguiente. Guardar: en un hilo, bajo `_catalogo_lock`, resolviendo los SAP repetidos con `VentanaResolverSAPDuplicados`; la ventana atiende esa pregunta y el final desde el hilo principal sondeando (`_vigilar_guardado`), sin llamar a Tk desde el hilo. Después completa los datos SAP con el cruce de la Acción 3 (`_leer_productos_sap` + `_leer_articulos_dali_sap` + `_leer_categoria_valoracion` + `_cruzar_y_escribir_listado_sap`, que regenera y envía a la web); si falta algún archivo, solo envía los códigos. Un código que no está en `Productos.xlsx` se guarda igualmente y se avisa: descargar de nuevo `Productos.xlsx` y `Listas de productos.xlsx` en RECURSOS y pulsar la Acción 3 (decisión de Víctor). La lectura de `Productos.xlsx` está en `sap_helpers.leer_productos_xlsx` (compartida con la Acción 3).
- **Un solo archivo de artículos** (v10.2.50, decisión de Víctor, 13 nov 2026): ya no existe `LISTADO CODIGOS DALI - SAP.xlsx`; todo está en `RECURSOS/catalogo-dali-completo.xlsx`. `app/comun/catalogo_local.py` (sin tkinter, `tests/test_catalogo_local.py`) lo ofrece **con la misma forma que tenía el LISTADO** (filas desde la 8, columnas A..O por posición: `COLUMNAS_LISTADO`), para no cambiar la lógica que ya funcionaba: los que solo leían (Acción 4 e INC, paso 27b de la tarifa, RINFO) usan `filas_formato_listado` / `abrir_como_listado`; la Acción 3 cruza sobre `libro_formato_listado` (en memoria) y pasa el resultado con `tabla_sap_desde_libro` + `_regenerar_catalogo_dali_completo(tabla_sap=…)`; la ventana de SAP faltantes escribe con `abrir_para_editar` / `mapas_dali_sap` bajo `_catalogo_lock`; la Acción 2 solo regenera el catálogo (los SAP, el código de proveedor y el precio salen del catálogo anterior, **siempre por código DALI**: antes escribía el LISTADO por posición de fila). `_pasar_listado_al_catalogo()` (al abrir y al regenerar) pasa una vez el LISTADO que quede (`migrar_listado`: manda el LISTADO, informe en RECURSOS y lo renombra a «(antiguo, ya no se usa)»). Al descargar de la web se conservan los códigos SAP locales que la web no tiene (`conservar_sap_local`). Se quitó la protección de hojas con contraseña escrita en el código. `tests/test_catalogo_unico.py` prueba los flujos de punta a punta con las clases reales. No volver a leer ni escribir el LISTADO.
- **Inicio de sesión** (v10.2.50, decisión de Víctor, 13 nov 2026): la ventana principal arranca oculta (`withdraw`) y `_iniciar_sesion_consola()` (`app/principal/catalogo_dali.py`) abre `app/ui/ventana_login.py`. Solo entra quien inicia sesión en Catálogo Asignaciones con una cuenta activa (la web no deja entrar a las inactivas), de rol `admin` y con el permiso `admin-catalogo-escribir` o administrador principal (`catalogo_web.usuario_autorizado` / `comprobar_acceso`, con pruebas); si no, «Usuario no autorizado.», sin anotar nada. Las credenciales guardadas (DPAPI) se prueban solas y se borran si dejan de valer; «Salir» cierra la app. Lo que se hacía al arrancar va en `_tras_iniciar_sesion()`, salvo la comprobación de actualizaciones, que va **antes** a propósito (si el inicio de sesión fallara, la versión que lo arregle llegaría igual; por eso `VentanaActualizacion` sale aunque la principal esté oculta).
- **Etapa 6: la Consola escribe en Catálogo Asignaciones** (v10.2.50; en la web, v1.20.63 y su `docs/propuestas/etapa6-consola-escribe-catalogo.md`). Tras regenerar el catálogo (Acciones 2 y 3), completar una tarifa (Acción 4) o resolver un SAP duplicado, `_publicar_catalogo_web()` (`app/principal/publicar_catalogo_web.py`, en su propio hilo; si ya hay uno, se repite una vez al acabar) compara `RECURSOS/catalogo-dali-completo.xlsx` con una copia de la web guardada en `%LOCALAPPDATA%\ConsolaRegistroInfo\catalogo_web.json`, que solo pide lo cambiado (`GET /admin/catalogo/cambios?desde=`). Envía solo lo distinto a `POST /admin/catalogo/lote`, primero simulando, con la ventana de resumen (`app/ui/ventana_publicar_catalogo.py`). La lógica, sin tkinter, en `app/comun/catalogo_web.py` (`tests/test_catalogo_web.py`, `tests/test_publicar_catalogo_web.py`). Reglas: nunca se borra nada en la web; un precio o código de proveedor vacío en local no borra el de la web; lo editado en la web (o enviado por otro compañero) queda fuera por defecto y sigue protegido (`Cache.protegidos`); el código DALI y su descripción solo se cargan desde la Consola (decisión de Víctor, 12 nov 2026), así que la descripción local siempre manda. **Fase 1:** el Excel se sigue escribiendo igual y la consola dice si al final coinciden.

## Estructura
Desde la etapa 2 (nov 2026) el código está en el paquete `app/`, y `tarifa_app.py` sigue en la raíz como punto de arranque: lo usan `COMPILAR_TARIFA.bat`, `_bat_helper.py` (lee `APP_VERSION`) y `build_exe.py`. Cada módulo conserva su nombre de siempre; para encontrar una función: `grep -rn "def nombre" tarifa_app.py app`.
- `tarifa_app.py` (~200 líneas) — punto de arranque y clase `App(InterfazApp, AccionOrganizarDali, AccionListadoDali, AccionListadoSap, CatalogoDali, AccionesTarifa, PublicarCatalogoWeb, AccionAsignarSap, tk.Tk)`. Aquí quedan `__init__`, la cabecera y el pie, porque usan `APP_VERSION`, que **tiene que seguir en este archivo** (lo leen `_bat_helper.py` y `COMPILAR_TARIFA.bat`).
- `app/principal/` — los métodos de `App`, repartidos por temas en clases *mixin* (etapa 4, nov 2026, sin cambiar el código). Se llaman siempre como `self.metodo()`, así que un método puede vivir en cualquiera de ellas:
  - `interfaz.py`: construcción de la ventana, consola de log, botones y acciones sencillas (carpeta, buscar `TARIFA.xls`, abrir resultado). El botón 7 (tarifa vacía) está en construcción.
  - `accion_organizar_dali.py`: Acción 1. Su lógica está en `app/acciones/organizar_dali.py`.
  - `accion_listado_dali.py`: Acción 2.
  - `accion_listado_sap.py`: Acción 3.
  - `accion_tarifa.py`: Acciones 4 y 6 (TARIFA final, archivos INC / [INCOMPLETO], reinscripción SAP, botón Registro Info).
  - `catalogo_dali.py`: Catálogo DALI completo y Catálogo Asignaciones (`DALI_APP_API_BASE`, que desde v10.2.50 vive en `app/comun/catalogo_web.py`).
  - `publicar_catalogo_web.py`: envío a Catálogo Asignaciones de lo cambiado (etapa 6, v10.2.50).
  - `accion_asignar_sap.py`: Acción 8, Asignar SAP pendientes (v10.2.50).
  - Un método nuevo va al mixin de su tema. Dos mixins no pueden tener un método con el mismo nombre, porque ganaría el primero de la lista sin avisar.
- `app/comun/` — `asignar_sap.py` (Acción 8, sin tkinter, v10.2.50), `catalogo_local.py` (el catálogo como único archivo de artículos, v10.2.50), `catalogo_web.py` (envío a Catálogo Asignaciones, v10.2.50), `manual.py` (abrir y consultar el manual), `tarifa_constants.py` (colores, fuentes, tablas fijas, `resource_path` y `RAIZ_PROYECTO`), `tarifa_core.py` (configuración, consola de log, barra de progreso), `sap_helpers.py` (funciones de negocio puras, sin tkinter), `credenciales.py` (DPAPI, sin tkinter) y `actualizaciones.py` (actualización automática, sin tkinter).
- `app/ui/` — `tarifa_dialogs.py` (ventanas modales: selección de .xls, crear proyecto, SAP duplicados y SAP faltantes), `ventana_actualizacion.py` (aviso de versión nueva), `ventana_consulta_manual.py` («Consultar el manual»), `ventana_login.py` (inicio de sesión al abrir, v10.2.50), `ventana_asignar_sap.py` (Acción 8, v10.2.50), `ventana_publicar_catalogo.py` (resumen de cambios para la web, v10.2.50) y `ventana_notas_version.py` (novedades de la versión al pulsar el número de versión de la cabecera o del pie; las pide a la Release de GitHub de esa versión con `actualizaciones.notas_de_version`).
- **Manual de usuario** (nov 2026): `docs/Manual_Usuario_ConsolaRegistroInfo.docx` / `.pdf`, con el estilo de los de Control Pedidos y el Organizador. Se regenera con `docs/manuales_fuente/` (capturas reales con datos inventados en Xvfb y maquetación con `docx` de Node; ver su README). El PDF y `manual_indice.json` van también en `assets/` (dentro del `.exe`): el acceso rápido «Manual de usuario» abre el PDF y la barra «🔎 Consultar el manual» busca en el índice (`app/comun/manual.py`, sin tkinter, pruebas en `tests/test_manual.py`). Si un cambio altera una pantalla o un texto del manual, avisar de que conviene regenerarlo (y copiar otra vez el PDF y el índice a `assets/`).
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
- **Icono** (`assets/logo.ico`, nov 2026): corona dorada sobre un cuadrado azul marino con borde dorado, para distinguirlo del Organizador (la corona sola). Se genera con `python compilacion/generar_icono.py` a partir de `compilacion/corona_icono.png`, un tamaño a la vez para que se vea nítido a 16 y 32 px. Lo ponen en las ventanas y en la barra de tareas `poner_icono_aplicacion()` y `cargar_icono_ventana()` (`tarifa_core.py`).
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
- 185 pruebas (desde la raíz). 18 de ellas importan `crear_registro_info` (que usa tkinter): sin tkinter se saltan. En Linux hay que lanzarlas con un Python que tenga tkinter y con `xvfb-run`.
- Dependencias para probar: `pandas openpyxl xlrd pillow psutil`. Para compilar, además `pyinstaller pywin32`.
- Código muerto: `vulture tarifa_app.py app --min-confidence 60`, comprobando cada resultado con `grep` (también en cadenas) antes de quitar nada.
- `tests/test_crear_registro_info.py` prueba el código real a través de `sap_helpers.py` y de `crear_registro_info`; no se deben reimplementar funciones dentro de la prueba.
- Comprobación de la compilación: `build_exe.py` también funciona en Linux (genera un ejecutable de Linux). Sirve para ver que PyInstaller incluye todo `app/` y que la app arranca.

# Historial de reorganización — Proyecto CONSOLA REG INFO SAP

**Propósito de este documento:** dar continuidad a un trabajo de refactorización y corrección de bugs que se lleva haciendo a lo largo de varias sesiones con Claude. Al empezar una sesión nueva, pega este documento entero como primer mensaje. Al terminar cada sesión (o cada vez que se cierre una tanda de trabajo), pide a Claude que actualice este mismo documento añadiendo lo nuevo, para mantenerlo siempre al día.

**Proyecto:** Central Compras Princess Canarias — VAMA. Dos aplicaciones interconectadas:
- `tarifa_app.py` — app de escritorio (ttkbootstrap/tkinter) "Organizador Princess", varias Acciones (organizar DALI, actualizar listados DALI/SAP, generar tarifa/RINFO, etc.)
- `crear_registro_info.py` — motor de generación de Registros de Información SAP (RINFO) a partir de tarifas de proveedor + Purchasing Info Records existentes.

**Ruta del proyecto:** `G:\CARPETA COMPRADORES\COMPRADOR 1 - VICTOR MARTIN\01.CONSOLA REG INFO SAP\CODIGO` (unidad de red).

---

## Estado del control de versiones

Git montado desde cero en esta sesión. `master` es siempre la versión estable y probada. El flujo de trabajo establecido:

1. `git checkout -b nombre-rama` antes de tocar nada.
2. Se entregan los ficheros modificados, se corren los tests (`python -m unittest test_crear_registro_info -v`, deben salir todos en verde).
3. Se prueba en real con `python tarifa_app.py` y datos reales, comparando el log contra ejecuciones anteriores conocidas-buenas.
4. Solo si todo cuadra: `git checkout master` → `git merge nombre-rama` → `git branch -d nombre-rama`.

`.gitignore` excluye: `build/`, `dist/`, `*.xlsx`/`*.xls` (datos reales, con excepción de `assets/`), `RECURSOS/`, logs (`log_*.txt`), `tarifa_config.json` (estado de sesión local), `*.zip` (backups manuales antiguos), `SEGURIDAD/` (carpeta de backups zip manuales previa a git).

**Regla de oro aprendida:** comprobar SIEMPRE `git branch` antes de copiar ficheros nuevos, para no acabar commiteando cambios grandes directamente en `master`.

**Segunda regla de oro aprendida:** commitear el fix validado en la rama de trabajo ANTES de darlo por "listo para fusionar" — no dejarlo solo como cambio sin commitear en el working directory. Ya pasó una vez: `git merge` dio `Already up to date.` en vez de `Fast-forward` porque la rama no tenía ningún commit propio (el cambio viajaba pegado al working directory, no al historial). No se perdió nada, pero tocó detectarlo con `git status` + `git log --oneline` y commitear a posteriori en `master`. Desde entonces, verificar siempre que el `merge` diga `Fast-forward` (o liste el fichero actualizado) y no `Already up to date.` antes de borrar la rama.

---

## Patrón de trabajo establecido (repetir en cada tanda nueva)

1. Buscar primero código duplicado o bloques aislables de bajo riesgo antes de trocear un método gigante de golpe (probado en `_organizar_dali` de `tarifa_app.py` y en `crear_registro_info.py`).
2. Extraer a función/método con un contrato de entrada/salida claro.
3. **Auditoría sistemática de variables**: antes de dar por bueno un troceo, listar TODAS las variables asignadas en el bloque extraído y comprobar una por una si se usan después en el código que se queda fuera. Dos bugs reales de este tipo ya han ocurrido (`output_cols`/`errores` en tarifa_app.py, `nombre_nuevo` en crear_registro_info.py) — en ambos casos, el fallo no se detectó hasta ejecutar en real, porque corría en un hilo sin captura de excepción y moría en silencio.
4. Probar con datos reales de forma aislada cuando sea posible (instanciar la clase con `object.__new__()` saltándose `__init__`, rellenar solo los atributos necesarios, invocar el método directamente sin hilo — así se ven los `tracebacks` al momento).
5. Añadir tests automáticos cuando la función se preste a ello (sin tkinter).
6. Entregar, el usuario prueba en real con `python tarifa_app.py`, compara logs, confirma.

---

## 1. `tarifa_app.py` — Acción 1 (Organizar DALI) — CERRADO Y FUSIONADO A MASTER

**Antes:** `_organizar_dali()` tenía 1008 líneas, mezclando UI/hilos con toda la lógica de negocio.
**Después:** 322 líneas (solo orquestación). Lógica de negocio extraída a `acciones/organizar_dali.py` (nuevo módulo, 581 líneas, sin tkinter):
`procesar_datos_dali`, `generar_excel_dali`, `eliminar_archivos_origen`, `validar_diccionarios_sap`, `detectar_codigos_faltantes`, `guardar_log_incidencias`.

**Bugs reales corregidos de paso (no solo reorganización):**
- **Duplicados de código de origen sin detectar**: si ARTICULOS1/2.xls tenían el mismo código repetido, el merge de pandas generaba un efecto cartesiano (multiplicaba filas). Ahora se deduplica ANTES del cruce, y si dos apariciones del mismo código tienen datos distintos (ej. proveedor distinto), se documenta en el log Y se genera un fichero de incidencias en `RECURSOS/INCIDENCIAS_DUPLICADOS_DALI_<timestamp>.txt` + aviso `messagebox` al usuario.
- **Filas fantasma en "LISTADO CODIGOS DALI - SAP.xlsx"**: al reducirse el nº de filas entre ejecuciones, solo se vaciaban las columnas A/B/C/L de las filas sobrantes, dejando basura en D-K y un `ws.max_row` inflado que arrastraba conteos erróneos a Acción 3 (`"30093 artículos cruzados"` en vez de 30090 reales). Solucionado con `ws_dest.delete_rows()` (borra la fila entera, no celda a celda).
- **Bloqueo de hoja**: se añadió `_proteger_hoja_destino()`, llamado antes de cada guardado de "LISTADO CODIGOS DALI - SAP.xlsx" en Acciones 2 y 3, para que quede siempre protegida contra edición, la encuentre como la encuentre.
- Logo de Acción 1 (150 líneas de base64 embebido) movido a `assets/logo_dali_articulos.png`, cargado con el `resource_path()` ya existente en el proyecto.
- Catálogos `DALI_CLAVES_SUBFAMILIA` / `DALI_CLAVES_IGIC` movidos de estar hardcodeados dentro del método a `tarifa_constants.py`.

Todo validado en real, varias veces, con logs de producción comparados línea a línea. **Fusionado a `master`.**

---

## 2. `crear_registro_info.py` — consolidación de duplicados — CERRADO Y FUSIONADO A MASTER

**`test_crear_registro_info.py` — hallazgo importante:** los 39 tests que ya existían **no importaban `crear_registro_info.py`** (para evitar arrastrar tkinter, que se importa a nivel de módulo ahí). Eran reimplementaciones paralelas de la lógica, que podían divergir del código real sin que ningún test lo detectara.

**Solución de fondo:** nuevo módulo `sap_helpers.py`, sin tkinter, con funciones de negocio puras reutilizables tanto desde `crear_registro_info.py` como desde el test:
- `normalizar_codigo_sap(valor)` — normaliza floats/int/str a código SAP limpio. **Consolidaba 31 copias idénticas** del mismo patrón de 3 líneas (`isinstance(x, float) and x==int(x): x=int(x); str(x).strip()`) repartidas por todo el fichero.
- `ultima_fila_con_datos(ws, columna, fila_inicio)` — detecta la última fila con datos en una columna. **Consolidaba 5 copias idénticas.**

`crear_registro_info.py`: 8263 → 8223 líneas. `test_crear_registro_info.py` ahora tiene 43 tests (39 originales + 4 nuevos de `ultima_fila_con_datos`), y **TC13 (normalización SAP) ahora protege código de producción real**, no una copia.

Todo validado con datos reales (generación completa de RINFO con absorciones, solapamientos, etc.). **Fusionado a `master`.**

---

## 3. `crear_registro_info.py` — troceo de `GeneradorDesdeExistente._ejecutar()` — CERRADO Y FUSIONADO A MASTER

`GeneradorDesdeExistente` es la clase más grande del proyecto (3613 líneas). Su método `_ejecutar()` (el orquestador de generación de RINFO) tenía ~1502 líneas. Internamente ya estaba dividido en **12 pasos numerados y comentados** (1 a 12), lo cual dio un mapa claro para trocear por partes a lo largo de varias tandas.

### Las 5 extracciones, todas fusionadas a `master`:

**Pasos 1-3** (nombre del archivo destino, copiar Purchasing, abrir libro y validar hoja) → `_preparar_archivo_destino()`. Contrato: devuelve `(wb, ws, ruta_nuevo, nombre_nuevo)` o `None` si falla (con su propio manejo de error/limpieza ya resuelto dentro).

- **Bug real cometido y corregido en el camino:** al extraer, `nombre_nuevo` se quedó sin incluir en lo que se devolvía, pero se seguía usando más abajo en `_ejecutar` (log final + ventana de confirmación). Como corre en un hilo sin try/except propio, el fallo (`NameError`) mataba el hilo en silencio — la app se quedaba "colgada" a ojos del usuario (barra parada al 95%, sin ningún error visible), hasta que se relanzó con la terminal visible y salió el traceback. **Lección aprendida: cuando algo "se queda colgado" sin motivo aparente en esta app, sospechar primero de una excepción silenciosa en el hilo de fondo, no de lentitud de red/antivirus.**

**Paso 6c** (el bloque más grande, ~647 líneas: actualizar fechas y precio en Hoja 3 para materiales heredados con match en tarifa) → `_actualizar_h3_heredados(self, wb, ws, codigos_existentes, codigos_tarifa, dic_tarifa) -> dict`. Devuelve `_absorbidas_con_tarifa` (mismo dict que antes alimentaba el paso 9b). Auditoría de variables sin hallazgos — ninguna fuga detectada. Fusionado en commit `51911e0`.

**Paso 4c** (~272 líneas: limpiar marca AJ=X heredada en materiales que sí continúan, gestionar bajas automáticas/decisión de usuario para los que se quedan sin tarifa) → `_limpiar_marca_aj_heredados(self, wb, ws, codigos_existentes, codigos_tarifa_set) -> list`. Devuelve `sin_tarifa`, usada mucho más adelante solo para `self._audit["n_sin_tarifa"]`. Fusionado en commit `cea8dee`.

**Pasos 4-6** (~163 líneas: leer materiales existentes del Purchasing, limpiar campo NUEVO residual en las 3 hojas, leer la tarifa nueva, calcular materiales sin match) → `_leer_purchasing_y_tarifa(self, wb, ws, ruta_nuevo)`. Devuelve una tupla de 6 elementos (`codigos_existentes, codigos_tarifa, codigos_tarifa_set, nuevos, ultima_fila_datos, max_id_col_a`) — todas se usaban más adelante en `_ejecutar`. Tenía un `return` de salida temprana (si falla la lectura de la tarifa); al extraer, ese `return` pasó a `return None`, y `_ejecutar` comprueba `if _resultado_lectura is None: return` — mismo contrato que `_preparar_archivo_destino()`. Fusionado en commit `31ac1b6`.

**Paso 6b-bis** (~169 líneas: propagar marca de borrado de Hoja1 col AJ a Hoja2 col AN y Hoja3 col Y, cerrando col J en la fila más reciente de cada grupo material+centro) → `_propagar_borrado_h1_a_h2_h3(self, wb, ws) -> None`. El más limpio de los cinco: ninguna variable local cruzaba la frontera, todo se resuelve escribiendo directamente sobre `wb`/`ws` y `self._audit["materiales_baja"]` — no necesita devolver nada. Fusionado en commit `2337832`.

### Resultado final
`_ejecutar()` pasó de ~1502 líneas a **231 líneas** de pura orquestación de los 12 pasos, cada uno delegando en un método con contrato de entrada/salida claro. Las 5 extracciones se validaron con 43 tests en verde y con datos reales de dos proveedores distintos (Ahumados Canarios — que ejerce bajas automáticas, absorciones y materiales sin fila previa en H3 — y LORMACAN CONSTRUCCIONES — que ejerce inserción masiva de materiales nuevos), comparando logs y `.xlsx` resultantes celda a celda contra ejecuciones conocidas-buenas.

**Cierre de versión:** dado el volumen del cambio (5 extracciones + 2 fixes reales de negocio + fix de UM/Z heredado de sesión previa), se marcó como salto mayor: `APP_VERSION` en `tarifa_app.py` pasó de `"v 9.0.6"` a `"v 10.0.0"` (commit `92eeaec`).

---

## 4. BUG REAL DE NEGOCIO encontrado y corregido (no relacionado con refactor) — CERRADO Y FUSIONADO A MASTER

Durante las pruebas reales de la Tanda de `_ejecutar`, el usuario detectó que un material (código SAP `10101839`) con precio válido en la tarifa nueva, y presente en Hoja1/Hoja2, **no aparecía en absoluto en Hoja3** (la hoja de condiciones de precio que SAP realmente usa).

### Causa raíz confirmada:
El paso 6c (línea ~4212 de `_ejecutar`) **solo recorre filas que YA EXISTEN en Hoja3** para actualizarlas (`for fila_row in ws_h3_act.iter_rows(...)`). Si un material tiene match en la tarifa pero **nunca tuvo ninguna fila en Hoja3** (comprobado con el `Purchasing Info Records (2).xlsx` original: el hueco ya venía de antes, no lo causa el código), nunca entra en ese bucle, y su condición de precio simplemente no se genera. **Confirmado que afecta a 50 materiales** en el caso de prueba real (Ahumados Canarios).

### Arreglo aplicado (fusionado a `master` en commit `0709925`):
1. Tras construir `periodos_h3` (diccionario de todas las claves PIR existentes en Hoja3), se calcula `heredados_con_match - materiales_en_periodos_h3` → materiales con tarifa que nunca tuvieron fila en H3.
2. Esos materiales se añaden al mismo diccionario `_absorbidas_con_tarifa` que ya alimentaba la función existente `_generar_filas_absorbidas_con_tarifa()` (reutilizada tal cual, sin duplicar lógica — para esa función da igual si la fila nueva hace falta por absorción o porque nunca existió).
3. Log agrupado por material (no material×centro) para no generar cientos de líneas — ya corregido y probado (de ~500 líneas a ~50).

**Validado con los 3 ficheros reales** (Purchasing, tarifa, ejecución completa de `_ejecutar()` fuera de la UI): el material 10101839 pasó a tener sus 10 filas nuevas en Hoja3 (una por centro), con el precio correcto (11,95€) y el período de la tarifa nueva.

### ✅ RESUELTO — duplicados en Hoja2 (era el problema abierto de la sesión anterior):
**Causa confirmada:** `_generar_filas_absorbidas_con_tarifa()` escribía SIEMPRE una fila nueva en Hoja2 para cada combinación material×centro recibida, sin comprobar si ya existía. El docstring original asumía que un material absorbido "necesitaba" una fila H2 nueva — pero H2 no lleva fechas/períodos (eso vive solo en H3), así que esa fila nueva nunca hacía falta, ni para el caso nuevo (nunca tuvo H3) ni para el caso de absorción genuina.

**Arreglo:** antes del bucle, se lee `ws_h2` una vez y se construye un set de combinaciones `(material, centro)` ya existentes. Dentro del bucle, si la combinación ya está en Hoja2, se **salta la escritura física en Hoja2** (evita el duplicado) pero se **sigue generando la tupla para Hoja3** igualmente (esa parte nunca dependía de haber escrito una fila H2 nueva). Log agrupado actualizado para distinguir "fila(s) generada(s)" de "ya existían en H2, sin duplicar".

**Validado con los 3 ficheros reales:** Hoja2 pasó de 20 a 10 celdas para el material 10101839 (las 10 originales, sin duplicar), Hoja3 sigue con sus 10 filas nuevas correctas. **Hallazgo adicional confirmado en la misma prueba: el bug también afectaba a las absorciones genuinas** (10101583, 10100148, 10101136) — el arreglo los corrige a todos a la vez, con la misma comprobación. 43 tests siguen en verde.

**Estado:** confirmado en real por el usuario (43 tests OK + `.xlsx` revisado celda a celda) y **fusionado a `master`** (commit `0709925`).

---

## 5. `crear_registro_info.py` — `_sincronizar_um_hoja3` no debe marcar Z ("Nueva") en filas heredadas — CERRADO (sesión anterior, no documentada hasta ahora)

**Nota:** este punto no se trabajó en esta sesión — se reconstruye a partir de tres `CHANGELOG_UM_HOJA3*.md` sueltos en la carpeta del proyecto, de una sesión previa que nunca se volcó a este historial. Verificado en esta sesión que las tres iteraciones **ya están aplicadas en el código actual** (`RINFO_9_0_4.zip`), en su versión final correcta.

**Síntoma original:** tras corregir un bug gemelo en Hoja 1 (marca AK), 3 de 4 materiales con UM desactualizada se corrigieron en SAP; el 4º (material `10103364`, registro `5300001868`) seguía fallando con "Por favor, adecúe la UM-base y el factor de conversión" en las 10 organizaciones/centros.

**Causa raíz (mismo patrón que el bug de AK en Hoja 1):** `Z="X"` (col 26, "Nueva") en una fila heredada de Hoja 3 le dice a SAP "crea una condición de precio nueva", pero esa organización/centro ya tiene una condición existente — SAP no puede crear una nueva sobre una que ya existe sin indicación de reemplazo.

**Tres iteraciones hasta la corrección final:**
1. **v1:** `_sincronizar_um_hoja3` deja de marcar Z al corregir R/S. Insuficiente: `10103364` seguía fallando porque Z la marcaba una función distinta, `_marcar_modificado()`, en un paso anterior (ajuste de Fin validez por cambio real de fecha).
2. **v2:** `_sincronizar_um_hoja3` fuerza Z a vacío en la fila heredada si la encuentra en "X" (sin importar quién la marcó), y descuenta el material de `materiales_modificados` para que el resumen no mienta. Seguía sin funcionar en la práctica.
3. **v3 (bug real, no de diseño):** la causa de que v2 no funcionara era `ws_h3.cell(row=fila_num, column=COL_Z_H3_NUEVA, value=None)` — openpyxl **no borra** la celda con `value=None` (`Worksheet.cell()` solo escribe si `value is not None`), así que la "X" quedaba intacta pese a que el log decía lo contrario. Corregido cogiendo la celda y asignando `.value = None` directamente:
   ```python
   celda_z = ws_h3.cell(row=fila_num, column=COL_Z_H3_NUEVA)
   celda_z.value = None
   ```
   Verificado con prueba aislada (`cell(value=None)` deja "X"; `.value = None` sí la vacía). Revisado el resto del fichero por si el mismo patrón se repetía en otro punto — no se encontró ningún otro caso.

**Nota pendiente de la sesión original:** si SAP sigue rechazando `10103364` en concreto tras este fix, revisar el maestro de materiales (MM03) por si tiene definida una unidad alternativa o algún bloqueo específico que el RINFO no pueda resolver desde este lado.

**Estado:** las tres iteraciones están en el código actual (confirmado línea a línea en `_sincronizar_um_hoja3`, con el mismo comentario explicativo del bug de openpyxl incluido). El fix hermano de AK en Hoja 1 también sigue presente. No consta el commit exacto en el historial de git en el momento de escribir esto.

---

## 6. `crear_registro_info.py` / `sap_helpers.py` — auditoría de limpieza post-v10.0.0 — CERRADO Y FUSIONADO A MASTER

Tras cerrar el troceo completo de `_ejecutar()` (v10.0.0), se pidió una auditoría general del proyecto para detectar duplicados reales antes de seguir añadiendo funcionalidad. Búsqueda automática de bloques de código duplicados en los ~17.500 líneas del proyecto, más revisión manual de los candidatos con más apariciones.

**Confirmado sin novedad:** ningún duplicado nuevo de los patrones ya consolidados antes (`normalizar_codigo_sap`, `ultima_fila_con_datos`) — la lección de la Tanda 2 se mantiene.

**Dos duplicados reales encontrados y consolidados:**

1. **Parseo de fecha de celda** (`None`/`datetime`/`date`/string en 3 formatos → `datetime.date`): la misma lógica de ~15 líneas copiada **5 veces** en `crear_registro_info.py` (`_parse_fecha`, `_a_date_h3`, `_parse_date_bis`, y 2 variantes inline). Consolidado en `parsear_fecha_celda(val)` dentro de `sap_helpers.py`, con 9 tests nuevos (`TC18_ParsearFechaCelda`). Los 4 puntos que definían la función anidada ahora solo hacen `_nombre_original = parsear_fecha_celda` (alias, cero riesgo en las llamadas existentes); el quinto (cálculo de `_fecha_cierre_h3` en el paso 6b-bis) se reescribió para usar la función directamente.

2. **Localización del LISTADO CODIGOS DALI-SAP** (buscar el fichero por ruta exacta en 3 carpetas y, si no aparece, por patrón de nombre): la misma lógica de ~30 líneas copiada en **4 métodos** (`_cargar_listado_dali`, `_cruzar_con_listado_dali`, `_rellenar_hoja2`, `_rellenar_hoja3`) — solo uno de ellos tenía la lógica como método reutilizable; los otros tres la reimplementaban entera. Consolidado en `_localizar_ruta_listado_dali()`. **Bug real corregido de paso:** solo `_cargar_listado_dali` no tenía el *fallback* `getattr(self, "_carpeta_proyecto", ...)` que las otras tres copias sí tenían — podía mirar en una carpeta distinta si `self._carpeta_proyecto` estaba definido. Ahora las 4 usan la versión correcta y consistente. La extracción solo tocó la parte de "localizar y abrir" — la lógica de extracción de columnas de cada método (distinta en cada uno) no se tocó.

**Resultado:** `crear_registro_info.py` con **115 líneas menos**, sin cambiar comportamiento. 52/52 tests en verde (43 + 9 nuevos). Validado en real con Ahumados Canarios, comprobando específicamente que las columnas derivadas del LISTADO (Hoja1 col F/N, Hoja3 col R) seguían rellenas al 100% tras el cambio. Fusionado a `master` en commit `babdeae`.

**Candidatos de limpieza identificados pero NO abordados (baja prioridad, quedan para otra sesión):**
- Boilerplate de `__init__`/geometría de ventana repetido entre las clases `VentanaTelefonoRINFO`, `VentanaFechasRINFO`, `VentanaCentrosRINFO` y varios diálogos de `tarifa_dialogs.py` — cosmético, extraíble a una clase base común si se quiere.
- Métodos grandes fuera de `_ejecutar()` que siguen siendo monolitos: `transformar_tarifa.py::transformar` (866 líneas), `tarifa_app.py::_actualizar_listado_sap` (752 líneas), `crear_registro_info.py::_actualizar_h3_heredados` (656 líneas — el propio paso 6c, ya extraído a método pero aún grande por dentro), `_rellenar_h2_nuevas_filas` (480 líneas), `tarifa_app.py::_ejecutar_reinscripcion_sap_desde_inc` (453 líneas). Ninguno formaba parte de los 12 pasos de `_ejecutar()`; candidatos para el mismo tratamiento si en el futuro dan problemas de mantenimiento.

---

## 7. Roadmap de limpieza priorizado — EN CURSO, orden establecido para próximas sesiones

Tras cerrar v10.0.0 y la auditoría de duplicados (punto 6), se estableció un orden priorizado de limpieza para ir ejecutando en sesiones sucesivas. **Al retomar, basta con decir "continúa" y se sigue por el primer punto pendiente de la lista, con la misma metodología ya rodada.**

### Fase A — Métodos gigantes de negocio (mayor impacto/riesgo latente)
| # | Método | Líneas | Estado |
|---|---|---|---|
| 1 | ~~`transformar_tarifa.py::transformar`~~ | ~~866~~ → 349 | **CERRADO** — ver sección 8 |
| 2 | ~~`tarifa_app.py::_actualizar_listado_sap`~~ | ~~752~~ → 163 | **CERRADO** — ver sección 11 |
| 3 | ~~`crear_registro_info.py::_actualizar_h3_heredados`~~ | 619 → 3 componentes (30+144+468+48) | **CERRADO** — ver sección 17 |
| 4 | `tarifa_app.py::_ejecutar_reinscripcion_sap_desde_inc` | 453 | Pendiente — ya se corrigió un bug real dentro de este método (sección 8) sin llegar a trocearlo |
| 5 | `crear_registro_info.py::_rellenar_h2_nuevas_filas` | 480 | Pendiente |

### Fase B — Boilerplate cosmético (bajo riesgo, bajo beneficio inmediato)
6. Clase base común para ventanas tkinter duplicadas: `VentanaTelefonoRINFO`, `VentanaFechasRINFO`, `VentanaCentrosRINFO` y varios diálogos de `tarifa_dialogs.py` comparten `__init__`/geometría casi idénticos.

**Criterio de orden:** primero lo de mayor tamaño e impacto de negocio (datos SAP reales), con el historial de bugs reales como desempate; lo puramente cosmético (punto 6) al final.

### Metodología a aplicar en cada punto (idéntica a la usada en los 12 pasos de `_ejecutar()`)
1. Auditoría de variables: qué entra al bloque, qué sale, qué se usa después.
2. Extraer a método/función con contrato de entrada/salida claro.
3. Tests unitarios cuando la función se preste a ello (sin tkinter, vía `sap_helpers.py` si es lógica pura).
4. Entregar; el usuario prueba en real con datos conocidos (Ahumados Canarios / LORMACAN como referencia) y compara logs/`.xlsx` contra ejecuciones conocidas-buenas.
5. Rama Git por cada extracción (`git checkout -b`), commit descriptivo, merge a `master`, actualizar este historial.

---

## 8. Roadmap punto 1 — `transformar_tarifa.py::transformar` — CERRADO Y FUSIONADO A MASTER

Primer punto del roadmap de la sección 7, ejecutado en tres tandas sucesivas dentro de la misma sesión. `transformar()` ya estaba numerada en 31 pasos (igual que `_ejecutar()` en sus 12), lo que dio un mapa claro para trocear.

### Las 3 extracciones:

**PASO 14b — Logo en A1:A5** (213 líneas, casi toda base64 embebido) → `_insertar_logo_cabecera(ws, log_fn)`. El bloque más autocontenido de los tres: solo entra `ws`/`log_fn`, nada sale. Riesgo mínimo. Validado en real (proveedor COCAMPA). Commit `c66f0c5`.

**PASO 27b — Cruzar CODIGO DALI con LISTADO DALI-SAP** (157 líneas) → `_cruzar_con_listado_dali_sap(ws, ruta_xls, ultima, log_fn) -> list`. **Fuga real encontrada:** `_sin_desc_und` (códigos DALI sin descripción/unidad SAP en el listado) se usaba mucho más adelante, tanto en el diccionario `_analisis` final como para decidir si el archivo se marca `[INCOMPLETO]`. La función ahora la devuelve explícitamente. De paso se simplificaron 2 usos de `"_sin_desc_und" in dir()` que ya no hacían falta (la variable se inicializa siempre al principio de la función). Validado en real (proveedor PALMAFRUT SL, confirmando que el marcado `[INCOMPLETO]` seguía disparándose correctamente). Commit `bb84f37`.

**PASO 31 en adelante — autoajuste, formato condicional, análisis de duplicados, guardado y borrado del .xls original** (162 líneas) → `_finalizar_y_guardar(wb, ws, ultima, ruta_xls, nombre_base, sin_desc_und, log_fn)`. Es literalmente el resto de `transformar()` hasta su `return` final — mismo contrato de salida `(ruta_salida, _analisis)`. Cero riesgo de fuga porque no hay nada después. Validado en real (COCAMPA y PALMAFRUT SL). Commit `4154259`.

### Resultado final
`transformar()` pasó de **866 a 349 líneas** (-60%). Nota: este módulo importa `win32com`/`pythoncom` (Windows-only) y no tiene suite de tests — toda la validación de las 3 extracciones se hizo con pruebas reales (`python tarifa_app.py` + `TARIFA.xls` reales de varios proveedores), no con `unittest`.

### Bug real encontrado y corregido durante las pruebas (no relacionado con el troceo)
Durante la prueba real del PASO 31, el usuario detectó que un archivo `[INCOMPLETO]` corregido a mano (eliminando duplicados) y reprocesado con Acción 4 **no eliminaba el `[INCOMPLETO]` sobrante** cuando ya existía un archivo con el nombre final en la carpeta — el código antiguo, al encontrar el destino ya ocupado, simplemente avisaba en el log y no hacía nada más, dejando ambos archivos (el completo y el `[INCOMPLETO]` huérfano) conviviendo para siempre.

**Causa:** el mismo patrón "si ya existe, avisar y no renombrar" estaba duplicado **4 veces** en `tarifa_app.py`, todas dentro de la lógica de Acción 4 (`_ejecutar_reinscripcion_sap_desde_inc` y alrededores).

**Arreglo:** las 4 ocurrencias ahora, si el destino ya existe, lo eliminan y sustituyen por el archivo recién corregido, dejando solo el correcto. Comportamiento confirmado seguro porque el flujo de Acción 6 (edición manual del `[INCOMPLETO]`) solo permite **eliminar** líneas duplicadas/incorrectas, nunca añadir — así que el archivo final siempre es una versión estrictamente derivada y correcta, sin riesgo de perder ediciones manuales propias del archivo final antiguo. Validado en real (proveedor PALMAFRUT SL): log confirma `🗑 Sustituido archivo final anterior` seguido de `✔ Archivo renombrado`, con un único archivo final resultante. Commit `6485123`.

---

## 9. `crear_registro_info.py` — regresión propia en `_localizar_ruta_listado_dali` — CERRADO Y FUSIONADO A MASTER

**Contexto:** bug introducido por la propia consolidación de la sección 6 (misma sesión), detectado por el usuario con una prueba real inmediatamente después.

**Síntoma:** el flujo "Crear Registro Info nuevo desde cero" (cuando no hay Purchasing Info Records compatible) se quedaba colgado en la barra de progreso, en el paso "Cruzando con LISTADO CODIGOS DALI-SAP...", sin ningún error visible.

**Causa:** `_localizar_ruta_listado_dali()` se dejó como **método** de `GeneradorDesdeExistente` (flujo "desde Purchasing existente"). Pero de los 4 puntos que se sustituyeron por la llamada a ese método, **3 pertenecían a `CreadorRINFO`** (flujo "crear desde cero") — una clase completamente distinta, sin relación de herencia. `CreadorRINFO` no tiene ese método → `AttributeError` en cuanto se llama → como corre en un hilo de fondo sin captura de excepción, el hilo muere en silencio y la app se queda "colgada". Exactamente el mismo patrón de fallo ya documentado en el punto 3 de este historial (excepción silenciosa en hilo de fondo).

**Arreglo:** `_localizar_ruta_listado_dali()` convertida en **función de módulo independiente** (`_localizar_ruta_listado_dali(carpeta_tarifas, carpeta_proyecto=None, parent=None)`), sin depender de `self`. Las 4 llamadas actualizadas para pasar `self._carpeta_tarifas`, `getattr(self, "_carpeta_proyecto", None)` y `self._parent` como argumentos explícitos.

**Validación:** se instaló `tkinter` en el entorno de pruebas para poder importar el módulo completo y simular la llamada real desde ambas clases (`GeneradorDesdeExistente` y `CreadorRINFO`) sin `AttributeError`. Validado también en real por el usuario: generación completa de RINFO desde cero para AMD MAYORISTAS EN FRUTAS Y VERDURAS (132 artículos, sin Purchasing), log confirma `LISTADO DALI-SAP localizado` y el resto del flujo hasta `RINFO guardado correctamente`, sin colgarse. 52/52 tests en verde. Fusionado a `master` en commit `343178a`.

**Lección aprendida:** al consolidar código duplicado entre varios puntos del fichero, verificar primero que **todos** los puntos pertenecen a la misma clase (o no pertenecen a ninguna) antes de decidir si la consolidación debe ser un método (depende de `self`) o una función de módulo independiente. Un vistazo rápido a los números de línea no basta — hay que confirmar explícitamente los límites de clase (`class ... :`) de cada punto a consolidar.

---

## 10. Sesión 2026-08-05 — fixes de empaquetado/errores + cambios propios del usuario

### Fixes de Claude, fusionados a `master`:

**`assets/logo_dali_articulos.png` faltaba en el empaquetado del `.exe`.** El logo de Acción 1 (Organizar artículos DALI) se omitía siempre en el `.exe` compilado (aviso `WARN`, no bloqueante) porque nunca se incluía como recurso al compilar — solo `logo.png` y `logo.ico` se empaquetaban. Funcionaba en modo `.py` porque `resource_path()` usa la carpeta del proyecto directamente en ese caso; en el `.exe`, resolvía bien la ruta a la carpeta temporal `_MEI...` de PyInstaller, pero el fichero nunca llegaba ahí.

- **Primer intento (fallido):** se añadió la entrada que faltaba a `TARIFA_Finder.spec`. No sirvió — `COMPILAR_TARIFA.bat` **borra `TARIFA_Finder.spec` en su PASO 3** antes de cada compilación, y `build_exe.py` genera el comando de PyInstaller desde cero mediante `--add-data` (sin leer el `.spec`), así que cualquier edición manual del `.spec` se pierde siempre. Confirmado con una segunda compilación real: el aviso `Logo omitido` seguía apareciendo.
- **Fix correcto:** añadido `logo_dali_articulos.png` directamente en `build_exe.py` (ruta, comprobación de existencia, línea de resumen y `--add-data`), con el mismo patrón defensivo que ya usaban `logo.png`/`logo.ico`. Verificado ejecutando `build_exe.py::main()` real (interceptando la llamada a PyInstaller para no compilar de verdad) para confirmar que el `--add-data` se generaba correctamente antes de pedir al usuario que recompilara. **Confirmado en real por el usuario tras recompilar: el logo ya aparece correctamente.**

Fusionado a `master` en dos commits: `117d15f` (intento con el `.spec`, inocuo pero insuficiente) y `95da005` (fix real en `build_exe.py`).

**Mensaje de error genérico al guardar `ARTICULOS_DALI.xlsx` con el archivo abierto en Excel.** Un compañero tenía el archivo de destino abierto en Excel (carpeta compartida de red) y el guardado falló con `[Errno 13] Permission denied`, mostrando el diálogo genérico "El proceso no se completó correctamente" con el traceback crudo. Se revisó `generar_excel_dali()` (en `acciones/organizar_dali.py`) y se confirmó que el código no tiene ningún handle propio abierto que pudiera causar el bloqueo — la causa es siempre externa (archivo abierto por alguien, sincronización de red, antivirus). Se añadió un manejador específico de `PermissionError` en el bloque de Acción 1 de `tarifa_app.py`, antes del genérico, con un mensaje claro: *"El archivo de destino está abierto en Excel... Cierra el archivo y vuelve a ejecutar la acción."* — mismo patrón que ya existía en otro punto del código para un caso similar.

Fusionado a `master` en commit `117d15f` (el aviso de PermissionError; el fix del logo se completó después, ver arriba).

### Cambios propios del usuario (fuera de esta sesión, confirmados por él, commiteados en la misma tanda):

- **Columnas fijas de Hoja2 adaptadas al manual SAP "Crear Registros Info" (punto 3.1.2):** eliminada la escritura de la columna **AK** (37, "Entrega ilimitada permitida") — el manual detalla explícitamente todas las columnas fijas de la pestaña 2 (A, B, C, D/E, F/G, H/I, J/K, L, M, N, S, W/X, AD/AE, AF, AH, AJ, AO) y AK no figura entre ellas. Queda documentado en el propio código (comentario "AUDITORÍA 2026-08") para que, si algún día SAP/central confirma que sí hace falta, se reintroduzca con la referencia exacta del manual que lo respalde. Commit `756c3b4`.
- **`COMPILAR_TARIFA.bat`** modificado para verificar e instalar automáticamente las dependencias que falten antes de compilar. Commit `ae07bf8`.

### Versión
`APP_VERSION` la gestiona el usuario manualmente (instrucción explícita: Claude no la toca; en su lugar, pregunta al cerrar cada fase si conviene subir de versión). Compilada en la sesión del 2026-08-05 como **v 10.2.4**. Al cierre de la sesión del 2026-08-06 (avance del punto 2 del roadmap + fixes de orden/protección), el usuario sube la versión a **v 10.0.8**.

---

## 11. Roadmap punto 2 — `tarifa_app.py::_actualizar_listado_sap` — CERRADO

Sesión 2026-08-06. Primer avance sobre el método de 752 líneas (Acción "Actualizar listado materiales SAP").

**Auditoría inicial:** el método está estructurado en 6 secciones numeradas (igual que `_ejecutar()`/`transformar()`). La sección 5 ("Diccionarios de referencia — equivalentes al VBA") contenía dos tablas estáticas hardcodeadas sin ninguna lógica: `TABLA_SUBFAMILIA` (265 entradas: subfamilia DALI → código+nombre de grupo de compras SAP) y `TABLA_IGIC` (6 entradas: texto IGIC → indicador de impuesto SAP). Comprobado que solo se usaban dentro de `proceso()` (la función anidada que corre en el hilo de fondo), un único punto cada una, sin fugas.

**Extracción:** ambas tablas movidas a `tarifa_constants.py` como `TABLA_SUBFAMILIA_GRUPO_COMPRAS` / `TABLA_IGIC_GRUPO_COMPRAS` — mismo patrón ya usado en la Tanda 1 con los catálogos `DALI_CLAVES_SUBFAMILIA`/`DALI_CLAVES_IGIC` (datos estáticos sin lógica no deben vivir dentro de un método). `_actualizar_listado_sap()` baja de 752 a **477 líneas**.

**Validación:** como `tarifa_app.py` no tiene tests automáticos, se hizo una prueba funcional real en el sandbox (import completo del módulo con `win32com`/`pythoncom` simulados) confirmando que las 265+6 entradas llegaban intactas con el mismo contenido exacto. Validado también en real por el usuario: ejecución de la Acción "Actualizar listado materiales SAP" con datos reales, cruces de SUBFAMILIA e IGIC correctos, sin discrepancias. Fusionado a `master` en commit `120deb5`.

**Pendiente dentro de este mismo punto:** la sección 6 ("Proceso en hilo") sigue siendo un bloque grande (LECTURA 1, LECTURA 2, CRUCE Y ESCRITURA, LIMPIEZA PREVIA, contadores de resumen) sin trocear todavía — próximo candidato dentro del punto 2 cuando se retome.

### Dos bugs reales adicionales, detectados durante las pruebas de esta misma sesión (no relacionados con el troceo del punto 2)

**`ARTICULOS_DALI.xlsx` no salía ordenado por CODIGO.** El usuario detectó que Acción 1 estaba generando el archivo con los códigos en orden descendente (30167, 30166, ... 1) en vez de ascendente, que es como siempre debió ser. Revisado `acciones/organizar_dali.py`: no había ningún `sort()` en todo el flujo — el `merge(how="left")` simplemente conservaba el orden de filas de `ARTICULOS1.xls` tal cual viniera del fichero de origen (sin garantía de orden). Corregido añadiendo `df_final = df_final.sort_values("CODIGO", ascending=True).reset_index(drop=True)` justo después del filtro de códigos válidos, en `procesar_datos_dali()`. Validado con una prueba funcional real (códigos sintéticos 10→1 de entrada, salida 1→10) y confirmado en real por el usuario tras ejecutar Acción 1.

**`LISTADO CODIGOS DALI - SAP.xlsx` se protegía sin contraseña.** `_proteger_hoja_destino()` (que ya se llamaba correctamente al final de Acción 2 y Acción 3, protegiendo el archivo esté como esté al empezar) aplicaba protección "blanda" sin contraseña. El archivo real usa la contraseña `"Princess2026"` para desproteger en Excel. Añadida `ws_dest.protection.password = "Princess2026"` junto al `sheet=True` ya existente — confirmado con una prueba real que openpyxl hashea la contraseña automáticamente (no queda en texto plano dentro del `.xlsx`) y que sobrevive al ciclo guardar/releer.

De paso se aclaró (sin cambiar código) que `D4`/`D5` en el resumen de Acción 3 muestran la fecha de última modificación de `Productos.xlsx` (frescura del dato de origen SAP), no la hora de ejecución de la acción — comportamiento intencional, confirmado con el usuario que se deja tal cual.

Ambos fixes fusionados a `master` en commit `d5dd5d0`.

### Segundo avance del punto 2: LECTURA 1 y LECTURA 2 extraídas

Dentro de la sección 6 ("Proceso en hilo"), los dos bloques de lectura de ficheros de origen extraídos a métodos independientes:

- **`_leer_productos_sap(ruta_productos)`** (49 líneas): lee `Productos.xlsx`, construye `tabla_productos` (clave=código SAP).
- **`_leer_articulos_dali_sap(ruta_dali)`** (49 líneas): lee `ARTICULOS_DALI.xlsx`, construye `tabla_dali` (clave=código DALI normalizado).

`_actualizar_listado_sap()` baja de 477 a **411 líneas**.

**Bug real detectado en producción por el usuario y corregido:** la primera versión de `_leer_productos_sap()` solo devolvía `tabla_productos`, asumiendo que `len(tabla_productos)` equivaldría al recuento de filas leídas (`filas_prod`) — pero `filas_prod` se usaba más adelante, fuera de la función, para la celda `H5` y el resumen final (`H1=... H2=... H3=... H5=...`), y **no es lo mismo** si `Productos.xlsx` trae códigos SAP duplicados (el diccionario deduplica por clave, el contador de filas no). El fallo saltó en real como `NameError: name 'filas_prod' is not defined` al ejecutar Acción 3. Corregido devolviendo `(tabla_productos, filas_prod)` explícitamente; validado con una prueba sintética con un código duplicado que confirmó la diferencia real entre ambos valores (3 filas leídas, 2 claves únicas), y revalidado en real por el usuario (`H1=6159 H2=543 H3=562 H5=8295`, sin errores). Fusionado a `master` en commit `2fb2df2`.

### Tercer y último avance: cierre del punto 2

**"CRUCE Y ESCRITURA EN DESTINO" + "Resumen en consola"** (270 líneas) extraídas juntas a `_cruzar_y_escribir_listado_sap(ruta_destino, tabla_productos, tabla_dali, filas_prod, fecha_productos_solo, hora_productos_solo)` — mismo patrón que `_finalizar_y_guardar()` en `transformar_tarifa.py`: es literalmente el resto del cuerpo del `try` de `proceso()`, sin nada después salvo el `except` que envuelve a `proceso()` entero, así que no hay riesgo de fuga hacia adelante.

**Auditoría reforzada tras la fuga anterior:** esta vez, antes de tocar nada, se identificaron explícitamente las 6 variables de entrada (`ruta_destino`, `tabla_productos`, `tabla_dali`, `filas_prod`, `fecha_productos_solo`, `hora_productos_solo`) y se confirmó que ninguna otra variable local del bloque se usa después. Además, se hizo una **prueba funcional de extremo a extremo** antes de entregar nada: las tres funciones nuevas (`_leer_productos_sap`, `_leer_articulos_dali_sap`, `_cruzar_y_escribir_listado_sap`) encadenadas con ficheros `.xlsx` sintéticos completos (incluyendo un caso con SAP, un caso sin SAP, match de SUBFAMILIA e IGIC), verificando manualmente cada columna del resultado (E-L) y los contadores (H1/H2/H3/H5) antes de dar el fichero por bueno.

**Resultado:** `_actualizar_listado_sap()` pasa de 752 a **163 líneas** (pura orquestación: validaciones + confirmación al usuario + las 3 llamadas en cadena). Validado también en real por el usuario: `H1=6159 H2=543 H3=562 H5=8295`, idéntico a la ejecución conocida-buena anterior, sin errores ni avisos nuevos. Fusionado a `master` en commit `70412e8`.

### Resumen completo del punto 2 (las 4 extracciones)
| Extracción | Líneas | Commit |
|---|---|---|
| `TABLA_SUBFAMILIA_GRUPO_COMPRAS`/`TABLA_IGIC_GRUPO_COMPRAS` → `tarifa_constants.py` | -275 | `120deb5` |
| `_leer_productos_sap()` | 53 | `2fb2df2` |
| `_leer_articulos_dali_sap()` | 49 | `2fb2df2` |
| `_cruzar_y_escribir_listado_sap()` | 270 | `70412e8` |

`_actualizar_listado_sap()`: **752 → 163 líneas** (-78%).

---

---

## 12. Roadmap punto 3 — `crear_registro_info.py::_actualizar_h3_heredados` — EN CURSO (limpieza conservadora)

**Auditoría inicial (sesión 2026-08-07):** a diferencia de `_ejecutar()` y `transformar()`, este método (656 líneas, el propio paso 6c) **no se presta a un troceo mecánico seguro**. Internamente tiene dos fases documentadas como "PRIORIDAD 2" (construye 10 estructuras mutables: `periodos_h3`, `precio_heredado_h3`, `i_heredada_h3`, `j_heredada_h3`, `fila_mas_reciente_h3`, 5 contadores, `filas_a_eliminar_6c`, `_absorbidas_con_tarifa` —el valor de retorno del método—, y 4 diccionarios de warnings) y "PRIORIDAD 1" (un único bucle de 417 líneas que lee y muta esas 10 estructuras simultáneamente, fila a fila). Separarlas en dos métodos exigiría un contrato de ~10 parámetros/retornos — alto riesgo de fuga (el mismo tipo de bug que `filas_prod` en el punto 2, pero aquí con más en juego, al ser el código de los 2 bugs reales más serios del proyecto).

**Decisión (con el usuario):** en vez de forzar un troceo arriesgado, se optó por una limpieza conservadora esta sesión, dejando el rediseño más a fondo (agrupar las 10 estructuras en un objeto de estado, y entonces sí trocear limpio) como tarea aparte pendiente para una sesión dedicada.

### Limpieza conservadora realizada
4 funciones auxiliares puras (sin dependencia de `self`/`wb`/`ws`) extraídas de `_actualizar_h3_heredados()` a `sap_helpers.py`:
- **`formatear_fecha_dmy(d)`** — consolida 2 copias duplicadas (`_date_str` en `_ajustar_fechas_fin_h3`, `_date_a_str_h3` aquí mismo).
- **`clave_pir_h3(fila_row)`**, **`material_de_clave_pir(clave_pir)`**, **`formatear_clave_pir(clave_pir)`** — lógica pura sin duplicar en otro sitio, movida para hacerla testable y reducir ruido. Las 2 funciones que sí dependían de estado compartido (`_material_en_baja`, `_marcar_modificado`) se dejaron intactas donde estaban.

**Validación en dos niveles:** 9 tests nuevos (`TC19_HelpersH3`, casos límite incluidos: material como float, fila corta, clave con campos vacíos) — 61/61 en verde; y una **prueba funcional real** invocando `_actualizar_h3_heredados()` directamente (técnica ya documentada más arriba: `object.__new__()` + datos sintéticos) con dos materiales (mismo precio → solo cambia fecha; precio distinto → nuevo período completo), verificando celda a celda que Hoja 3 quedaba exactamente como se esperaba.

**No hay reducción de líneas del método** (sigue en ~640 líneas) — el objetivo fue reducir riesgo/duplicación, no tamaño.

### Bonus: log de centros descartados en Hoja2 abreviado (toca el punto 5 del roadmap)
Durante las pruebas, el usuario detectó que la generación de una tarifa con muchos materiales heredados sin match (LP de Ahumados Canarios) generaba **cientos de líneas de log**, una por cada fila de Hoja2 con centro descartado (`H2 fila 489: centro '0710' descartado → ...`). Se aplicó el mismo mecanismo `_LogAgrupado` que ya usa el proyecto en otros puntos, agrupando por material: **1144 filas individuales → 128 líneas resumen** (una por material, con la lista de centros y el rango de filas), validado en real sin pérdida de información. Este bloque vive dentro de `_rellenar_h2_nuevas_filas()` — el punto 5 del roadmap — así que queda parcialmente adelantado sin haberlo empezado formalmente.

Todo fusionado a `master` en commit `d8742e0`.

---

## 13. Optimización de rendimiento — borrado físico de filas — CERRADO Y FUSIONADO A MASTER

Sesión 2026-08-10. **Origen distinto al resto del historial:** este trabajo lo inició otra sesión/herramienta (no esta conversación) sobre `crear_registro_info.py`, y el usuario trajo el resultado aquí para auditarlo, testearlo y fusionarlo siguiendo el mismo flujo de siempre.

**Problema real detectado en producción:** generar un RINFO desde un Purchasing grande (proveedor EMICELA, S.A.) tardaba varios minutos en el paso de "eliminación física (coherencia con Hoja2)" de Hoja 3. Causa raíz: `ws.delete_rows(fila)` llamado una vez por cada fila a eliminar (4.518 filas dispersas) — cada llamada desplaza hacia arriba todas las filas por debajo, así que repetirlo fila a fila da un coste ~O(n²) (decenas de millones de operaciones de celda sobre una hoja de ~10.000 filas).

**Fix 1 — `eliminar_filas_agrupadas()`:** agrupa los índices de fila consecutivos en rangos contiguos y llama a `delete_rows(inicio, cantidad)` una sola vez por rango, en orden descendente. Sustituye el patrón fila-a-fila en los 4 puntos donde aparecía (Hoja2, Hoja3-propagación, paso 4, paso 6c). Resultado real: 4.518 filas, de ~6 minutos a **16 segundos**.

**Fix 2 — log agrupado en la propagación a Hoja3:** una vez resuelto el `delete_rows`, el cuello de botella se movió a `_log()` llamado una vez por fila (5.366 llamadas repintando la consola Tkinter, ≈4 minutos). Se aplicó `_LogAgrupado` (la misma clase ya usada para el caso equivalente en Hoja2) — ahora una línea resumen por material en vez de una por centro.

### Auditoría propia realizada en esta sesión (sobre el trabajo recibido)
Antes de fusionar nada, se verificó **matemáticamente** que `eliminar_filas_agrupadas()` produce un resultado idéntico al método fila-a-fila original: 8 casos aleatorios (semillas fijas) + 8 casos límite explícitos (lista vacía, una sola fila, bloque consecutivo único, filas alternas, extremos primera/última fila, múltiples bloques dispersos, índices duplicados), comparando ambos métodos celda a celda sobre la misma hoja de partida.

**Único hallazgo:** con índices duplicados en la lista de entrada, el método nuevo (que deduplica con `set()`) difiere del original — el original habría repetido `delete_rows` sobre la misma posición y borrado una fila distinta la segunda vez (bug latente). Confirmado que las 4 listas reales se construyen en una única pasada por fila (nunca se añade el mismo número dos veces), así que este caso no debería darse en la práctica — mejora de seguridad, no una regresión.

**Pulido adicional:** la función se movió de `crear_registro_info.py` a `sap_helpers.py` (mismo patrón que el resto de funciones puras del proyecto, perfil similar a `ultima_fila_con_datos`) y se le añadieron **8 tests permanentes** (`TC20_EliminarFilasAgrupadas`), en vez de dejarla validada solo manualmente.

**Validación:** 69/69 tests en verde (61 + 8 nuevos). Confirmado en real dos veces por el usuario con logs de producción del proveedor EMICELA, S.A. (ambas ejecuciones terminan en `RESUMEN VALIDACIÓN SAP` + `Archivo generado` sin errores). Fusionado a `master` en commit `8602ecb`.

---

## 14. `tarifa_app.py` — resincronización de Acción 3 tras insertar 3 columnas nuevas en LISTADO CODIGOS DALI-SAP — CERRADO Y FUSIONADO A MASTER

Sesión 2026-08-12. El usuario insertó **3 columnas nuevas** en `LISTADO CODIGOS DALI - SAP.xlsx`: **F** (Texto UM expl.), **K** (Categoría valoración) y **L** (Tipo de producto) — desplazando todo lo que había a partir de ahí. Había que localizar y actualizar todo el código que dependía de la estructura antigua de columnas, fundamentalmente en Acción 3 (`_actualizar_listado_sap` / `_cruzar_y_escribir_listado_sap`).

### Desplazamiento de columnas (antigua → nueva)
| Antigua | Nueva | Contenido |
|---|---|---|
| F (UNIDAD SAP) | **G** | +1 |
| G (GRUPO PRODUCTOS SAP) | **H** | +1 |
| H (Desc. grupo productos) | **I** | +1 |
| I (Indicador impuestos / IGIC) | **J** | +1 |
| J (GRUPO COMPRAS SAP) | **M** | +3 |
| K (NOMBRE GRUPO COMPRAS SAP) | **N** | +3 |
| L (ACTIVO DALI) | **O** | +3 |

### Nuevo mapeo implementado
- **E, F, G, H, I, L** ← `Productos.xlsx` columnas C, J, I, F, G, D respectivamente (Cruce 1, vía `_leer_productos_sap()` — ahora devuelve 6 campos en vez de 4).
- **K** ← nueva fuente **`Listas de productos.xlsx`** (carpeta RECURSOS, mismo patrón que Productos.xlsx), columna F (Categoría valoración), cruzada por columna A (código SAP) — nuevo método `_leer_categoria_valoracion()`. El fichero trae una fila por combinación código+centro; cuando un código SAP tiene valores distintos según el centro (3 casos de 6.903 en los datos reales), se usa el **valor más frecuente** (moda) asumiendo que el resto son errores, y se lanza un **mensaje visual dedicado** listando los códigos y valores en conflicto para que el usuario los revise en el origen.
- **J** (IGIC) y **M, N** (Grupo de Compras) — misma lógica de cruce con `ARTICULOS_DALI.xlsx` de siempre, solo desplazadas a sus nuevas columnas.
- **O** (ACTIVO DALI) — solo lectura (para los contadores H1/H2/H3), nunca se escribe en este proceso.
- Limpieza previa de columnas ampliada de E:K a **E:N**.

### Validación
Prueba funcional de extremo a extremo (las 3 lecturas — Productos, ARTICULOS_DALI, Listas de productos — encadenadas con `_cruzar_y_escribir_listado_sap()`, ficheros sintéticos incluyendo un caso con conflicto de Categoría valoración) verificando columna por columna (E a O) que el resultado era exactamente el esperado, y que el aviso visual de conflictos se disparaba con el detalle correcto. Confirmado también en real por el usuario: `6704 materiales actualizados desde Productos.xlsx`, `6699 materiales con Categoría valoración`, `30173 artículos cruzados con ARTICULOS_DALI`, 3 conflictos detectados y mostrados en el diálogo visual con los mismos códigos y valores que en los datos reales, y la cabecera/celdas del `.xlsx` resultante revisadas manualmente coincidiendo con lo esperado.

**Hallazgo colateral, no corregido (pre-existente, fuera del alcance de este cambio):** dentro del bucle de `_cruzar_y_escribir_listado_sap()`, si un artículo no tiene coincidencia en `ARTICULOS_DALI.xlsx`, un `continue` salta directamente a la siguiente fila sin llegar a contar esa fila en H1/H2/H3 — aunque tenga código SAP y esté activa. Ya existía antes de esta sesión; queda anotado por si se quiere revisar en el futuro.

Fusionado a `master` en commit `f589f13`.

---

## 15. `crear_registro_info.py` — resincronización de ambos flujos RINFO con la nueva estructura del LISTADO — CERRADO Y FUSIONADO A MASTER

Continuación directa de la sección 14, mismo cambio estructural (3 columnas nuevas F, K, L insertadas en `LISTADO CODIGOS DALI - SAP.xlsx`), pero esta vez en el otro fichero grande del proyecto: `crear_registro_info.py`, donde viven **ambos flujos de creación de RINFO**.

### Puntos corregidos
Se localizaron **4 funciones** que leían columnas fijas del LISTADO con el layout antiguo, con el mismo patrón que ya se había arreglado en `tarifa_app.py`:

- **`_cargar_listado_dali()`** (`GeneradorDesdeExistente`, flujo "desde Purchasing existente") — usada por `_rellenar_h1_nuevas_filas`, `_rellenar_h2_nuevas_filas`, `_generar_filas_absorbidas_con_tarifa` y las dos funciones de sincronización de UM (`_sincronizar_um_hoja1`/`_hoja3`).
- **`_cruzar_con_listado_dali()`** (`CreadorRINFO`, flujo "crear desde cero") — Hoja1: texto breve, grupo artículos, desc. grupo, unidad (N/P).
- **`_rellenar_hoja2()`** (`CreadorRINFO`) — Grupo de Compras, indicador de impuestos (IGIC), texto breve, grupo artículos.
- **`_rellenar_hoja3()`** (`CreadorRINFO`) — Unidad SAP (R/S).

Todas con el mismo desplazamiento de índices ya documentado en la sección 14 (antigua F/G/H/I → G/H/I/J; antigua J/K → M/N). Los consumidores que leen estos diccionarios **por índice de tupla** en vez de por columna de fichero (como `_sincronizar_um_hoja1`/`_hoja3`) no necesitaron ningún cambio — se corrigen automáticamente al arreglar la fuente.

### Validación
**4 pruebas funcionales aisladas**, una por función corregida, usando exactamente los valores reales de la fila 8 del `LISTADO_CODIGOS_DALI_-_SAP.xlsx` ya validado en la sección 14 (mismo material, mismos valores esperados en cada caso) — las 4 coincidieron exactamente con lo esperado. 69/69 tests generales en verde (no afectados, tocan otra parte del código).

**Confirmado también en real por el usuario, en los dos flujos:**
- **Desde Purchasing existente** (proveedor EMICELA SA GC): `LISTADO DALI-SAP cargado: 6702 entradas`, `Hoja1 — UM (N/P): sin discrepancias`, `Hoja3 — UM (R/S): sin discrepancias`. Revisado el `.xlsx` resultante directamente: columnas N/P de Hoja1 con valores reales no vacíos (`PI`, `BID`), confirmando que el diccionario corregido trae datos de verdad, no solo "sin discrepancias" por casualidad entre dos vacíos.
- **Desde cero** (proveedor CAFES CANDELAS CANARIAS, S.L.): `Cruce DALI-SAP: 3 coincidencia(s) de 3 artículo(s)` (Hoja1), `Hoja2 rellenada: 30 filas ... 30 cruces LISTADO` (Grupo Compras, IGIC), `Hoja3 rellenada: 30 fila(s) ... 30 cruces LISTADO (R/S)` (Unidad SAP) — sin errores en ninguno de los tres puntos corregidos.

Fusionado a `master` en commit `d2d0c62`. Con esto queda cerrada del todo la resincronización del cambio estructural del LISTADO CODIGOS DALI-SAP en toda la aplicación (Acción 3 en `tarifa_app.py` + ambos flujos de RINFO en `crear_registro_info.py`).

---

## 16. `tarifa_app.py` — Acción 2 (`_actualizar_listado_dali`) — dos bugs colaterales del cambio estructural — CERRADO Y FUSIONADO A MASTER

Sesión 2026-08-17. El usuario detectó que el material más reciente de `ARTICULOS_DALI.xlsx` (DALI 30174) llegaba con la columna **ACTIVO DALI vacía** en `LISTADO CODIGOS DALI - SAP.xlsx`, pese a que el origen sí traía el valor. Investigado a fondo, con una vuelta en falso incluida (una copia de trabajo desactualizada llevó a una hipótesis inicial equivocada, corregida al comparar directamente con los ficheros reales que subió el usuario).

### Bug 1 — ACTIVO DALI en la columna equivocada, con colisión real entre Acción 2 y Acción 3
`_actualizar_listado_dali()` (Acción 2) seguía escribiendo el indicador ACTIVO en la **columna L** — la posición que tenía antes de insertar las 3 columnas nuevas (sección 14), ahora ocupada por "Tipo de producto". Esto por sí solo ya sería un error de columna, pero además provocaba una **colisión real** con Acción 3: su "Limpieza previa" borra el rango **E:N** (que incluye L) en cada ejecución, y solo lo vuelve a rellenar si el material tiene código SAP asignado. Para cualquier material **sin match SAP** — el caso más común entre materiales recién añadidos — Acción 3 borraba lo que Acción 2 acababa de escribir en L, y nunca se reponía. Los valores `'SI'` que aún se veían en filas antiguas eran dato **residual**: habían quedado físicamente en la posición O tras la inserción de columnas (Excel arrastra el contenido de las celdas al insertar columnas), pero nadie los actualiza desde entonces.

**Arreglo:** escribir ACTIVO en la **columna O** — la posición correcta según la cabecera, y además fuera del rango E:N que limpia Acción 3, eliminando la colisión de raíz.

### Bug 2 — Formato condicional de columnas D y O con rango fijo, sin actualizar desde hace tiempo
El usuario notó que el resaltado de duplicados (columna D) y de "NO activos" (columna O) no llegaba hasta el final de la tabla. Revisando las reglas de formato condicional reales del `.xlsx`: ambas estaban fijadas a `D8:D30070` / `O8:O30070` — un rango que se quedó congelado desde que alguien lo configuró manualmente en Excel hace tiempo, sin que ningún código Python lo gestionara. Con los datos ya en la fila 30181, las últimas ~111 filas quedaban sin resaltar.

**Arreglo:** en `_actualizar_listado_dali()`, tras calcular la última fila real, se amplían automáticamente ambos rangos de formato condicional, **reutilizando las mismas reglas** (mismo `dxfId`, mismo color) en vez de crear unas nuevas — confirmado con una prueba aislada que el número de estilos `dxf` del `.xlsx` no cambia al hacerlo (sin duplicar formato).

### Validación
Prueba funcional de extremo a extremo ejecutando `_actualizar_listado_dali()` completo (con hilo forzado a síncrono y diálogos de confirmación simulados) sobre los ficheros reales del usuario (`ARTICULOS_DALI.xlsx` + `LISTADO_CODIGOS_DALI_-_SAP.xlsx`), confirmando `Última fila (A,O): 30174 SI` y ambos rangos de formato condicional ampliados a `D8:D30181` / `O8:O30181`. 69/69 tests generales en verde. Confirmado también en real por el usuario: el log de producción muestra `Formato condicional ampliado: D (D8:D30070 → D8:D30181), O (O8:O30070 → O8:O30181)` y `30174 artículos actualizados (A, B, C, O)`.

Fusionado a `master` en commit `4114199`.

---

## 17. Roadmap punto 3 — rediseño de fondo — CERRADO Y FUSIONADO A MASTER

Sesión 2026-08-17. Cierre del punto 3 que quedó a medias en la sección 12 (limpieza conservadora): el rediseño de fondo que agrupa las ~15 estructuras compartidas del bucle de `_actualizar_h3_heredados()` (619 líneas) en un solo objeto de estado, permitiendo por fin trocearlo con seguridad.

### Diseño
- **`_EstadoIndicesH3`** (dataclass, 30 líneas, a nivel de módulo, antes de `class GeneradorDesdeExistente:`) — agrupa los 5 índices por clave PIR (`periodos_h3`, `precio_heredado_h3`, `i_heredada_h3`, `j_heredada_h3`, `fila_mas_reciente_h3`), los 5 contadores (`n_solo_fecha`, `n_precio_y_fecha`, `n_solapamiento`, `n_reduccion`, `n_sin_fecha_heredada`), la lista de filas a eliminar, el dict de retorno (`absorbidas_con_tarifa`) y los 4 dicts de warnings — todos con `field(default_factory=...)` para no necesitar inicialización explícita.
- **`_construir_indices_h3_heredados()`** (144 líneas) — antigua PRIORIDAD 2: recorre Hoja 3 una vez y construye el estado.
- **`_aplicar_actualizaciones_h3_heredados()`** (468 líneas) — antigua PRIORIDAD 1: recorre Hoja 3 una segunda vez, muta el estado, escribe en Hoja 3, borra las filas absorbidas y vuelca los WARNs agrupados; devuelve `estado.absorbidas_con_tarifa`.
- **`_actualizar_h3_heredados()`** — pasa a ser un orquestador de 48 líneas: prepara `heredados_con_match`/fechas, llama a construir, llama a aplicar, devuelve el resultado.

### Validación — la más exhaustiva del roadmap
Dado que este es el código de los 2 bugs reales más serios del proyecto, la validación fue deliberadamente más rigurosa que en el resto de puntos:

1. **Comparación directa original vs. rediseñado.** Se ejecutaron ambas versiones del método (la subida por el usuario, sin tocar, y la rediseñada) sobre **el mismo workbook sintético**, con **8 escenarios** cubriendo cada rama del código: mismo precio, precio distinto sin solape (gap), solape parcial, absorción total en 2 centros, material con 2 períodos históricos distintos, material marcado en baja (AJ=X en Hoja1), material con tarifa activa que nunca tuvo ninguna fila en Hoja 3, y material sin tarifa (no debe tocarse). Se comparó **cada línea de log, cada celda de Hoja 3, el diccionario de retorno completo y el estado de auditoría** — resultado idéntico byte a byte entre ambas versiones.
2. **4 bugs de la propia extracción, detectados y corregidos antes de entregar nada** — ninguno llegó a salir del sandbox: indentación heredada del bloque `if` original (4 espacios de más), la dataclass insertada por error dentro del cuerpo de `GeneradorDesdeExistente` (cerraba esa clase antes de tiempo), un `import dataclasses` olvidado tras reiniciar el ensamblaje, y los alias `_clave_pir_h3`/`_mat_de_clave` (a `sap_helpers.py`) que se usan en **ambas** fases pero solo quedaron definidos en la primera — este último lo detectó la propia prueba comparativa con un `NameError`.
3. 69/69 tests generales en verde.
4. **Validado en real** (proveedor Ahumados Canarios, tarifa LP, generación desde Purchasing existente): proceso completo sin errores, `Archivo generado` al final.

### Hallazgo de investigación, no relacionado con el rediseño
El usuario notó que el log mostraba "995 solo fin-validez" mientras que el RINFO final solo tenía 320 filas en Hoja 3. Investigado a fondo con los datos reales del Purchasing: cada material trae entradas para **los 10 centros de Princess** (0110 a 0910), con una media de **9,35 claves PIR por material** — casi idéntico a 995÷107≈9,3. `_actualizar_h3_heredados()` procesa Hoja 3 mientras todavía contiene las filas de los 10 centros (copiadas del Purchasing en pasos anteriores del pipeline); el filtrado al centro seleccionado ocurre en un paso posterior, ya fuera del método rediseñado. Confirmado que es una característica preexistente del pipeline de varios pasos, completamente ajena a los cambios de esta sesión — no requiere corrección.

Fusionado a `master` en commit `64162db`. **Con esto, los puntos 1, 2 y 3 del roadmap de limpieza quedan CERRADOS.**

---

## 18. `tarifa_app.py` — bug real en la cabecera del LISTADO: contadores sobrescribían etiquetas de texto — CERRADO Y FUSIONADO A MASTER

Sesión 2026-08-19. El usuario notó que la fila 4 de la cabecera de `LISTADO CODIGOS DALI - SAP.xlsx` (columnas H/I) no cuadraba: preguntó si, al insertar las 3 columnas nuevas (sección 14), debería haberse ajustado también esta zona, ya que sospechaba que en vez de la columna H ahora debería ser la I.

**Investigación:** la cabecera (filas 1-5, fuera de la tabla principal que empieza en fila 8) tiene un pequeño panel de resumen. `_cruzar_y_escribir_listado_sap()` escribía los 4 contadores del resumen (`h1`, `h2`, `h3`, `filas_prod`) directamente en las celdas **H1, H2, H3, H5** — pero al inspeccionar el `.xlsx` real subido por el usuario, esas celdas en realidad contenían **etiquetas de texto** ("MATERIALES SAP / DALI ACTIVOS", "MATERIALES SAP / DALI NO ACTIVOS", "MATERIALES PENDIENTES DE CREAR EN SAP", "TOTAL MATERIALES SAP") que el código llevaba sobrescribiendo con números en cada ejecución, sin que nadie lo notara hasta ahora. La fórmula real del dashboard vive en **I4** (`=I5-I1-I2`, etiquetada por **H4** = "MATERIALES SAP NO CORRECTOS") y esperaba los valores en **I1, I2, I3, I5** — celdas que el código nunca tocaba, por lo que quedaban con datos congelados de mucho tiempo atrás.

**Decisión del usuario:** escribir directamente en I1/I2/I3/I5, dejando H sin usar (no que ambos clusters se mantengan en paralelo).

**Fix:** cambiadas las 4 asignaciones de `ws_dest["H1"]`/`H2`/`H3`/`H5` a `ws_dest["I1"]`/`I2`/`I3`/`I5`; el mensaje de log también pasa de `H1=... H2=... H3=... H5=...` a `I1=... I2=... I3=... I5=...`. De paso se limpiaron el comentario y los nombres de variable (`l_es_si`/`l_es_no` → `o_es_si`/`o_es_no`) que aún hacían referencia a la columna L, ya movida a O en la sección 16.

**Validación:** prueba funcional real confirmando que H1-H5 quedan vacías (ya no se escriben) e I1/I2/I3/I5 se rellenan correctamente. 69/69 tests en verde. Confirmado en real por el usuario con el `.xlsx` de producción: las etiquetas de H1-H5 se conservan visibles por primera vez en mucho tiempo, e I1=6140, I2=565, I3=587, I5=8301 llevan los valores correctos — la fórmula I4 (`=I5-I1-I2`) ya calcula sobre datos frescos en vez de sobre el snapshot congelado anterior.

Fusionado a `master` en commit `072e3eb`.

---

## 19. Catálogo DALI completo — nueva funcionalidad — CERRADO Y FUSIONADO A MASTER

Sesión 2026-09-16. Nuevo fichero `RECURSOS/catalogo-dali-completo.xlsx` que combina en 20 columnas: `ARTICULOS_DALI.xlsx` (Acción 1, cols A/C-I/T) + `LISTADO CODIGOS DALI - SAP.xlsx` (Acciones 2/3, cols J-S) + una columna B (Código Proveedor) propia.

**Diseño acordado tras varias iteraciones con el usuario:**
- **Regeneración automática, no botón manual** (se descartó un botón "Acción 8" inicial): `_regenerar_catalogo_dali_completo()` corre al final de Acción 2 y al final de Acción 3, cada vez que cualquiera de las dos actualiza el LISTADO. Protegida en su propio `try/except` en ambos puntos de enganche, para que un fallo aquí nunca rompa el éxito de Acción 2/3.
- **Columna B (Código Proveedor):** se conserva del catálogo anterior SOLO si la columna I (Proveedor Asignado - DALI) de ese código no cambió; si cambió, se vacía y se genera un `.txt` en RECURSOS/ con el detalle (código, proveedor anterior, proveedor nuevo), etiquetado explícitamente como **"CASO IRREGULAR"** (normalmente ambas bases de proveedores deberían coincidir exactamente) — con aviso también en log (detalle línea a línea) y ventana emergente.
- **Acción 4** (`_actualizar_codigo_proveedor_catalogo()`) actualiza las 20 columnas de los códigos DALI de cada tarifa completada (no solo B), cruzando en el momento con `ARTICULOS_DALI.xlsx` y el LISTADO — pero SOLO si el Proveedor Asignado coincide con el proveedor de esa tarifa (celda C2); si no coincide, no toca ninguna columna de esa fila (ni siquiera se añade si es nueva) y avisa igual (log + ventana, mismo etiquetado "CASO IRREGULAR").
- **Diseño visual exacto** replicado del fichero de referencia del usuario: título 14pt negrita, subtítulo 9pt cursiva gris (`FF666666`), cabecera fila 4 con relleno beige (`FFE5DECB`), anchos de columna exactos por contenido, autofiltro `A4:T4`.

**Bug propio detectado y corregido antes de fusionar:** faltaba `self._barra.completado()` al final de `_regenerar_catalogo_dali_completo()`. Como Acción 3 llama a su propio `completado()` ANTES de invocar la regeneración del catálogo (a diferencia de Acción 2, que lo llama después), la barra de progreso quedaba congelada indefinidamente en "Actualizando catálogo DALI..." — el usuario lo detectó horas después con una captura mostrando la barra parada. Corregido añadiendo la llamada final.

**Validación:** confirmado en real repetidas veces por el usuario — ciclo completo 1→2→regenerar→3→regenerar sin errores, y los casos irregulares reales detectados correctamente (ORTHIDAL SL→ALCRUZ CANARIAS SL, MESA SOCORRO→vacío, OASISFISH→vacío, más varios casos de renombrado cosmético tipo "CRUZBE"→"CRUZBE SL" que el usuario confirmó que son desfases puntuales esperados, no un fallo del sistema). 69/69 tests.

Fusionado a `master` en commit `0198445`.

---

## 20. Auditoría del punto 4 del roadmap — 3 bugs reales encontrados antes de empezar el troceo — CERRADO Y FUSIONADO A MASTER

Sesión 2026-09-17. Al empezar a auditar `_ejecutar_reinscripcion_sap_desde_inc()` (Acción 4, 458 líneas, punto 4 del roadmap de limpieza) para planificar su troceo, se encontraron 3 bugs reales sin relación directa con la reducción de líneas — el troceo en sí todavía no se ha empezado.

1. **Columna UNIDAD SAP desincronizada.** El propio método leía la columna F del LISTADO pensando que era "UNIDAD SAP" (así lo documentaba su docstring), pero desde el resync de columnas de la sección 14 (3 columnas nuevas insertadas), F pasó a ser "Texto UM (expl.)" y la UNIDAD SAP real se desplazó a la G. Esta ruta de código no formaba parte de ninguna de las Acciones resincronizadas en su momento (2, 3, `crear_registro_info.py`), así que se quedó desactualizada sin que nadie lo notara. Impacto: la columna DESCRIPCION de los archivos INC de tarifa (`desc_sap + " " + unidad_sap`) pegaba el texto expandido ("Pieza", "Kilogramo") en vez del código corto real ("PI", "KG").
2. **Tipo de dato inconsistente en "Código SAP" del catálogo.** Al comparar dos generaciones del catálogo DALI contra el fichero de referencia del usuario, 6.744 de 6.745 filas "distintas" resultaron ser solo una diferencia de tipo (`10200458` int vs `'10200458'` str) heredada de cómo estuviera tipada la columna D del LISTADO en cada momento — el propio código nunca normalizaba. Corregido forzando siempre texto en los dos puntos donde se lee (`_regenerar_catalogo_dali_completo` y `_actualizar_codigo_proveedor_catalogo`).
3. **Espacios sobrantes en campos de texto del catálogo.** La única diferencia de contenido real (no de tipo) en esa misma comparación era un espacio en blanco de más al principio de una descripción, heredado directamente de `Productos.xlsx` sin limpiar. Corregido añadiendo un helper `_lim()` (recorta solo si es texto, deja números/None intactos) aplicado a todos los campos de texto leídos de `ARTICULOS_DALI.xlsx` y el LISTADO, en ambos métodos del catálogo.

**Validación:** los 3 fixes en la misma rama, cada uno con prueba funcional real aislada. Confirmación final: tras regenerar el catálogo con datos reales (Acción 2→3) y comparar contra el fichero de referencia del usuario, **0 diferencias de 30.194 filas** — coincidencia exacta al 100%. 69/69 tests.

Fusionado a `master` en commit `21cf512`.

---

## 21. Roadmap punto 4, primera pasada de troceo — CERRADO Y FUSIONADO A MASTER

Sesión 2026-09-18. Alcance conservador acordado con el usuario: solo los pasos de menor riesgo de `_ejecutar_reinscripcion_sap_desde_inc` (Acción 4, 463 líneas), dejando el resto para otra sesión.

- **`_leer_tabla_cruce_listado_sap()`** (80 líneas) — pasos 1+2: localizar el LISTADO y construir la tabla de cruce DALI→SAP. Devuelve `(None, None)` en fallo, con su propio log/messagebox/barra.reset ya hecho — el llamador solo comprueba `if ruta_listado is None: return`.
- **`_log_resumen_reinscripcion()`** (45 líneas) — paso 6: resumen final en el log, barra completada, botón habilitado.
- Orquestadora baja de 463 a 374 líneas.

**Bug propio corregido antes de entregar:** faltaba `from openpyxl import load_workbook as _lw` dentro del método extraído (antes vivía en el scope de `proceso()`, invisible tras la extracción) — detectado por la propia prueba funcional con un `NameError`.

**Validación:** pruebas aisladas de ambos métodos (éxito y fallo) + prueba de integración completa de la orquestadora con datos sintéticos, y en real con un caso de producción genuino: CADELSA, primera ejecución con 8 SAP+8 DALI duplicados detectados correctamente (sin renombrar), segunda ejecución tras corregir en Excel sin duplicados y con renombrado correcto `[INCOMPLETO]`→final. 69/69 tests.

Quedan pendientes para otra sesión: pasos 3+4+5 (reaplicar SAP/DESC + duplicados + guardar, ~121 líneas) y paso 7 (diálogo `VentanaSAPFaltantes` + renombrado final, ~190 líneas — el más delicado, con closures anidadas y los dos puntos donde engancha el catálogo DALI).

Fusionado a `master` en commit `06d06e2`.

---

## 22. Bug real de producción: materiales de baja con UM obsoleta rechazados por SAP — CERRADO Y FUSIONADO A MASTER

Sesión 2026-09-18. El usuario reportó que los materiales marcados para borrado en el Purchasing existente no actualizaban su info y daban error real al cargar en SAP. Caso concreto: registro `5300002798`, material `10102980`, error SAP `"Por favor, adecúe la UM-base y el factor de conversión"` en los 10 centros — investigado hasta el final, con tres capas de causa distintas.

### Causa 1 — exclusión deliberada pero empíricamente incorrecta
`_sincronizar_um_hoja1()` (Hoja1, cols N/P) y `_sincronizar_um_hoja3()` (Hoja3, cols R/S) excluían explícitamente las filas marcadas para borrado (`AJ='X'` en Hoja1, `Y='X'` en Hoja3), bajo el razonamiento documentado *"no tiene sentido corregir la UM de un registro que se va a dar de baja"*. El log real de SAP demostró que ese razonamiento era incorrecto: SAP sí valida la UM-base incluso al procesar una baja. Quitada la exclusión en **ambos** métodos, de forma simétrica — el primer fix (solo Hoja1) no bastó: SAP seguía rechazando el registro comparando Hoja1 (ya corregida a CJ) contra Hoja3 (aún en BOL), con el aviso *"La conversión entre CJ y BOL se efectúa con factor 1"*.

### Causa 2 — 1.595 materiales SAP sin cruce DALI en el LISTADO
El material 10102980 no aparecía en ninguna columna del LISTADO CODIGOS DALI-SAP — ni él ni otros 1.594 materiales de `Productos.xlsx` (19% del total), típicamente materiales descatalogados y sustituidos, o duplicados nunca vinculados a un código DALI. Sin esa entrada, la sincronización de UM no tenía nada con qué comparar y se saltaba silenciosamente, dejando la UM obsoleta intacta. Añadido un **fallback** en `_cargar_listado_dali()`: tras cargar el LISTADO normalmente, se completa con `Productos.xlsx` (nueva función de módulo `_localizar_ruta_productos()`, mismo patrón que `_localizar_ruta_listado_dali()`) cualquier código SAP que no haya aparecido en el LISTADO. El LISTADO manda siempre que tenga la entrada; sólo se puede rellenar desde `Productos.xlsx` descripción/unidad/grupo de productos — IGIC y grupo de compras quedan vacíos igual que antes, porque dependen de la SUBFAMILIA del lado DALI, inexistente para un material sin cruce.

El usuario pidió explícitamente el fallback automático tras confirmar que el caso es recurrente (materiales que se descatalogan o duplican con cierta frecuencia), descartando la alternativa de ir corrigiendo el cruce a mano cada vez.

### Causa 3 — dato ya existente en SAP, resuelta fuera del código
Con ambos fixes de código aplicados y el `.xlsx` verificado correcto en las 4 hojas (incluida `Escalas p.cond.precio PIR`, vacía para este material), la carga a SAP **seguía fallando con el mismo error**. Diagnóstico final: el registro 5300002798 ya existía en SAP con `BOL` guardado desde su creación original; SAP no reconcilia un cambio de unidad en un registro **existente** vía carga masiva, compara lo nuevo contra lo ya almacenado y lo rechaza — un comportamiento deliberado de SAP para evitar inconsistencias con pedidos/precios ya calculados con la unidad antigua. El usuario lo resolvió eliminando el registro en SAP y dejando que la recarga lo creara de nuevo con `CJ`; confirmó que la carga entró limpia.

**Validación:** ambos fixes de código probados en real repetidas veces, con logs de SAP genuinos en cada iteración — incluyendo un caso donde el usuario subió un log más antiguo por error (detectado por Claude comparando timestamps) antes de dar con la causa 3. 69/69 tests.

Fusionado a `master` en commit `4330ae9`. App recompilada y confirmada como **v10.2.8**.

---

## 23. Integración con «Catálogo Asignaciones» (dali-sap-articulos-app) — descarga automática del catálogo — CERRADO Y FUSIONADO A MASTER

Sesión 2026-09-20. El usuario pidió que, antes de que las Acciones 2/3/4 modifiquen `catalogo-dali-completo.xlsx`, la aplicación avise de la fecha de última actualización de ese archivo y recomiende refrescarlo, ya que también se mantiene desde una app web separada (`dali-sap-articulos-app`, «Catálogo Asignaciones») donde otras personas asignan Código Proveedor manualmente. Lo que empezó como un simple aviso informativo acabó en una integración real entre las dos aplicaciones, a petición del propio usuario.

### Diseño final
- **`_confirmar_catalogo_actualizado()`** — diálogo de 3 vías (Sí/No/Cancelar) mostrado al principio de Acciones 2, 3 y 4: **Sí** descarga una copia fresca y continúa, **No** continúa con la copia actual, **Cancelar** aborta la acción.
- **Descarga automática real**, no solo un enlace manual: se investigó el repositorio de la app web (subido por el usuario) para entender su arquitectura — autenticación por sesión/cookie (no token), exportación síncrona (`GET /admin/export-completo` hace streaming directo del Excel pese a generarse en un worker interno, sin necesidad de sondear ningún job), y despliegue tras dos Cloudflare Workers como proxy inverso (`dali-proxy` → frontend, `dali-proxy-api` → backend) para esquivar bloqueos de antivirus corporativos a `*.onrender.com`. La URL del worker de API no estaba escrita en ningún sitio del repo (solo el patrón); confirmada por el usuario desde el propio dashboard de Cloudflare: `https://dali-proxy-api.centralcompras1-canarias.workers.dev`.
- **`_obtener_credenciales_dali_app()`** — diálogo modal para email+contraseña, pedido solo la primera vez; se guardan en texto plano en `tarifa_config.json`, decisión explícita del usuario (prioriza comodidad sobre no tener que teclearlas cada vez).
- **`_descargar_catalogo_dali_web(email, password)`** — login + descarga vía `urllib`/`http.cookiejar` (sin añadir dependencias nuevas al proyecto), con manejo claro de credenciales incorrectas, verificación en dos pasos pendiente (2FA por inactividad prolongada) y permisos insuficientes.

### Dos bugs propios encontrados en pruebas reales del usuario, corregidos en la misma rama
1. **Credenciales nunca se limpiaban en HTTP 403** — el código solo lo hacía en 401, pero el error real del usuario fue 403 ("Forbidden"), así que la contraseña equivocada se quedaba guardada para siempre y la app la reutilizaba en silencio sin volver a preguntar. Corregido a limpiar en ambos casos. De paso, se añadió un `User-Agent` de navegador real a las peticiones — el de `urllib` por defecto ("Python-urllib/3.x") es una firma reconocible de script que Cloudflare (WAF/Bot Fight Mode) puede bloquear con 403 independientemente de si las credenciales son correctas.
2. **Bug más serio: la interfaz se congelaba por completo** ("No responde" de Windows) — la descarga corría directamente en el hilo principal de Tkinter, bloqueándolo mientras esperaba respuesta de red (hasta 20s+60s de timeout). Corregido separando `_descargar_catalogo_dali_web(email, password)` (lógica de red pura, segura de llamar desde un hilo de fondo) de un nuevo wrapper `_descargar_catalogo_dali_web_bloqueante()`, que la ejecuta en un hilo aparte con una ventana modal de espera, usando el patrón estándar de Tkinter (`after()` con sondeo + `wait_window()`) — nunca un `join()` bloqueante ni una llamada de red directa en el hilo que espera. Importante: la obtención de credenciales (que crea una ventana Tkinter) se mantiene siempre en el hilo principal, nunca dentro del hilo de fondo.

### Validación
Pruebas funcionales reales con un servidor HTTP local que imita exactamente los dos endpoints reales (login + export-completo), incluyendo cookie de sesión mantenida entre peticiones, credenciales incorrectas limpiando el config, y una simulación del mecanismo de espera de Tk con latencia deliberada confirmando que la red corre genuinamente en segundo plano. Confirmado en real por el usuario: descarga real de 3.701.159 bytes completada sin congelar la interfaz, archivo guardado correctamente en `RECURSOS/catalogo-dali-completo.xlsx`. 69/69 tests.

Fusionado a `master` en commit `b2fc5c2`.

---

## 24. Refinamiento del enganche del catálogo + reversión del punto 22 (UM en materiales de baja) — CERRADO Y FUSIONADO A MASTER

Sesión 2026-09-21/22. Cuatro cambios encadenados, todos surgidos de uso real tras cerrar la integración con Catálogo Asignaciones (sección 23).

### 1. Memoria de sesión (Acción 2→3) y quitar el aviso en Acción 4
A petición del usuario: una vez resuelto el aviso de catálogo desactualizado en Acción 2 o 3, no se repregunta el resto de esa misma sesión de trabajo (flag `self._catalogo_verificado_sesion`, vive solo en memoria, se reinicia al reabrir la app; **Cancelar** no lo marca, para poder reintentar). Acción 4 deja de llamar a `_confirmar_catalogo_actualizado()` por completo — es un proceso de *carga* de datos, no de *actualización* de plantillas, y descargar ahí borraría lo que Acciones 2/3 acababan de dejar hecho.

### 2. Fix: la barra de progreso marcaba "completado" antes de tiempo en Acción 4
El usuario detectó que la barra se ponía en verde ("Reinscripción SAP completada") varios segundos antes de que el catálogo terminara de actualizarse de verdad — porque `_log_resumen_reinscripcion()` (paso 6, extraído en la sección 21) marcaba completado ANTES del paso 7 (renombrado + enganche del catálogo). Corregido moviendo el "completado" real a los puntos donde el proceso de verdad termina, en los dos caminos posibles (con y sin diálogo `VentanaSAPFaltantes`), más el caso de duplicados (que antes no cerraba la barra en absoluto) y una red de seguridad en el `finally` exterior (flag `_barra_finalizada`) que garantiza un estado final pase lo que pase, cubriendo también casos no explícitamente revisados (diálogo cancelado, "abrir en Excel"). Validado con una prueba de integración real confirmando el orden correcto de eventos.

### 3. Enganche del catálogo también en `_ejecutar_transformacion` (tarifa limpia a la primera)
Hallazgo real del usuario: cuando una tarifa sale limpia a la primera (sin duplicados, sin SAP faltantes — ej. ARGAL), el flujo nunca pasa por Acción 4 (`_ejecutar_reinscripcion_sap_desde_inc`, donde vivía el único enganche existente) — se guarda directamente desde `transformar_tarifa.py::transformar()`, un camino de código totalmente distinto que nunca actualizaba el catálogo. Corregido añadiendo el mismo enganche en `_ejecutar_transformacion()` (tarifa_app.py), con la MISMA condición exacta que usa `_finalizar_y_guardar()` para decidir si el archivo lleva `[INCOMPLETO]` — así el catálogo solo se actualiza cuando la tarifa queda de verdad lista. `transformar_tarifa.py` no necesitó ningún cambio: ya devolvía toda la info necesaria (`dup_sap`/`dup_dali`/`sin_desc_und`) en su contrato existente; el enganche va en `tarifa_app.py`, que es quien tiene acceso a `self`. De paso se corrigió aquí el mismo problema de "completado prematuro" del punto 2. Validado con una prueba funcional real usando la clase `App` real, y confirmado en producción con ARGAL (`"Catálogo DALI actualizado desde la tarifa: 18 fila(s) modificadas"`).

### 4. Reversión del punto 22: NO tocar ningún valor heredado en materiales de baja
El fix de la sección 22 (quitar la exclusión de materiales marcados para borrado en `_sincronizar_um_hoja1`/`_sincronizar_um_hoja3`) resultó **incompleto en su diagnóstico**: solo "funcionó" en el caso real (registro 5300002798) porque el usuario borró el registro en SAP a mano antes de recargar — SAP no reconcilia un cambio de UM en un registro que **ya existe**, se borre o no después. Sin ese borrado manual previo, cualquier otro material de baja con UM heredada distinta volvería a dar el mismo error. Revertido a la exclusión original: para materiales marcados para borrado, no se toca ningún valor heredado de ningún tipo (alcance acotado explícitamente por el usuario a la UM — teléfono y fechas fijas de Hoja1, que siempre se aplican a todos los materiales sin excepción, se dejan como estaban, ya que nunca dieron problema). Documentado en el propio código el porqué de la reversión, para no repetir el mismo error de diagnóstico. Validado con una prueba funcional real replicando el caso exacto (10102980, BOL, marcado para borrado) confirmando que ya no se toca, junto a un material normal que sigue corrigiéndose con normalidad.

**Validación conjunta:** 69/69 tests en todos los pasos. App recompilada y confirmada como **v10.2.12**.

Fusionado a `master`.

---

## 25. Punto 4 del roadmap CERRADO — paso 7 (VentanaSAPFaltantes + renombrado final)

Sesión 2026-09-22. Última pasada del troceo de `_ejecutar_reinscripcion_sap_desde_inc` (Acción 4): el paso 7 (~205 líneas), el más delicado por sus closures anidadas (`_mostrar_dialogo_inc`, `_avisar_sap_sin_desc`) que capturan variables de la llamada original.

**Enfoque deliberado:** a diferencia de los pasos anteriores, este NO se dividió en piezas más pequeñas — se extrajo el bloque completo tal cual a un único método, `_gestionar_faltantes_o_finalizar_inc()` (237 líneas), convirtiendo cada variable capturada por las closures en un parámetro explícito del método (`ruta_inc`, `_ruta_listado`, `_actualizados`, `_sin_sap`, `_sap_sin_desc`, `_dup_sap`, `_dup_dali`, `_barra_finalizada` por referencia). Partirlo en trozos más pequeños habría arriesgado una fuga de alguna captura, como ya pasó (y se corrigió) en pasadas anteriores de este mismo punto.

**Resultado final del troceo completo del punto 4** (varias sesiones): `_ejecutar_reinscripcion_sap_desde_inc` pasa de **463 a 102 líneas**, con cuatro métodos de contrato claro: `_leer_tabla_cruce_listado_sap()`, `_reaplicar_sap_y_guardar_inc()`, `_log_resumen_reinscripcion()`, `_gestionar_faltantes_o_finalizar_inc()`.

**Validación:** dos pruebas de integración reales (diálogo `VentanaSAPFaltantes` mockeado completando todos los faltantes, y camino limpio sin faltantes) y, en real por el usuario, un caso que ejercitó las tres ramas posibles en una sola sesión de trabajo continua (SABORES Y GUSTOS INUR SL): `VentanaSAPFaltantes` usado para completar 3 códigos → archivo se queda `[INCOMPLETO]` correctamente por duplicados restantes → corregido en Excel → Acción 4 final con SAP aplicado, renombrado y catálogo actualizado, sin ningún error. 69/69 tests.

Fusionado a `master` en commit `7826e31`. App recompilada y confirmada como **v10.2.14**.

**Con esto, el roadmap de limpieza tiene los puntos 1, 2, 3 y 4 CERRADOS.** Quedan pendientes el punto 5 (`crear_registro_info.py::_rellenar_h2_nuevas_filas`, 480 líneas, sin empezar formalmente) y el punto 6 (Fase B, baja prioridad: clase base común para las ventanas tkinter duplicadas).

---

## 26. Punto 5 del roadmap CERRADO — `_rellenar_h2_nuevas_filas` troceado en una sola pasada

Sesión 2026-09-23. A diferencia de los puntos anteriores, este método tenía una estructura mucho más limpia — 6 pasos numerados y prácticamente autocontenidos, sin las closures anidadas que complicaron el punto 4 — así que se hizo el troceo completo en una única pasada, tras confirmarlo con el usuario.

**Seis métodos nuevos:**
- `_marcar_borrado_logico_y_escanear_h2()` (110 líneas) — elimina físicamente filas de centro descartado, sanea la col M, construye los sets de deduplicación.
- `_propagar_borrado_h2_a_h3()` (43 líneas) — elimina en Hoja3 lo eliminado en Hoja2.
- `_leer_h1_para_expansion_h2()` (39 líneas) — lee materiales de Hoja1.
- `_escribir_h2_bloques_nuevos()` (116 líneas) — escribe filas nuevas por bloques de centro, con deduplicación; devuelve el resultado final del método original.
- `_sincronizar_igic_h2()` (100 líneas) — sincroniza IGIC en toda Hoja2.
- `_normalizar_obligatorios_h2_completo()` (102 líneas) — rellena AF/AJ/AH obligatorios.

`_rellenar_h2_nuevas_filas` pasa de **495 a 74 líneas**.

**Dos bugs propios detectados y corregidos antes de entregar:** (1) error de indentación — a diferencia de los métodos de Acción 4 (que vivían dentro de un `def proceso():` anidado con doble indentación), este es un método directo de la clase, así que el dedent de 8 espacios aplicado por costumbre rompía la sintaxis; corregido extrayendo los bloques sin dedentar, ya que su indentación original YA era la correcta para un método nuevo. (2) el `return filas_h2_escritas` final del método original (que pertenece a la orquestadora) quedó incluido por error dentro de `_normalizar_obligatorios_h2_completo()` al extraer el último paso — detectado por la propia prueba funcional con un `NameError`, nunca salió del sandbox.

**Validación:** prueba de integración real cubriendo los 5 comportamientos documentados a la vez (centro descartado con eliminación física, deduplicación, expansión de material nuevo, expansión de material parcialmente heredado, sincronización de IGIC, saneado de cantidad heredada) — todos correctos. 69/69 tests. Confirmado en real con AHUMADOS CANARIOS SA (proveedor de referencia habitual del proyecto): 1 fila eliminada por centro descartado, deduplicación funcionando sobre 107 materiales × 10 centros, IGIC y obligatorios sin discrepancias, archivo generado sin errores.

Fusionado a `master` en commit `3647dd2`. **Con esto, el roadmap de limpieza tiene los puntos 1, 2, 3, 4 y 5 CERRADOS** — solo queda el punto 6 (Fase B, boilerplate tkinter, baja prioridad, sin urgencia).

---

## 27. Bug real encontrado tras el troceo del punto 5 (sin relación con él): crash en «Crear Registro Info»

Misma sesión. El usuario reportó un crash real de la app (`_tkinter.TclError: invalid command name`) al cerrar la ventana de «Crear Registro Info».

**Causa:** `_abrir_ri()` en `tarifa_app.py` llamaba a `iniciar_crear_registro_info()` (que bloquea con `app.wait_window(ventana)`) **antes** de registrar el log `"Iniciando: Crear Registro Info…"` — es decir, el log en realidad se ejecutaba al **cerrar** la ventana, no al abrirla, con el nombre completamente al revés de lo que hacía. Si en ese momento el widget de consola ya no estaba disponible, crasheaba toda la aplicación.

**Fix, dos partes:**
1. Reordenado: el log ahora va antes de la llamada bloqueante, donde tiene sentido.
2. Protección defensiva centralizada en `log_consola()` (try/except `tk.TclError`) — cubre cualquier llamada futura de toda la app, no solo esta, por si se repite un escenario similar en otro punto.

**Validación:** prueba funcional simulando un widget destruido (confirma que el error se absorbe sin crashear) y en real por el usuario — dos aperturas/cierres consecutivos de «Crear Registro Info» sin ningún crash. 69/69 tests.

Fusionado a `master` en commit `cdad3ee`. App recompilada y confirmada como **v10.2.16**.

---

## 28. Tres mejoras en Acción 4: antigüedad del catálogo, recuperación de código proveedor, ventanas centradas

Sesión 2026-09-23. Tres cambios encadenados, todos sobre el flujo de Acción 4 y el enganche del catálogo DALI.

### 1. Comprobación de antigüedad específica de Acción 4 (30 minutos)
Acción 4 dejó de mostrar el aviso de catálogo desactualizado en la sesión anterior (para no arriesgar sobrescribir lo que ella misma acaba de escribir) — pero eso dejaba un hueco real: si se genera una tarifa sin haber corrido antes Acciones 1-3, se corre el riesgo de usar un catálogo desactualizado sin ningún aviso. Nuevo método `_catalogo_reciente_o_confirmar()`, enganchado en los 4 puntos donde Acción 4 escribe en el catálogo (camino limpio de `_ejecutar_transformacion` + los 3 caminos de `_gestionar_faltantes_o_finalizar_inc`): comprueba la antigüedad de `catalogo-dali-completo.xlsx` **y también** de sus dos ficheros de origen (`ARTICULOS_DALI.xlsx`, `LISTADO CODIGOS DALI - SAP.xlsx`). Si todo tiene menos de 30 minutos, se actualiza sin más, sin ningún aviso; si algo es más antiguo, muestra el mismo diálogo de 3 vías que Acciones 2/3 (indicando qué archivo y desde cuándo), con Cancelar limitado a saltar solo el paso del catálogo (la tarifa en sí ya quedó completada igualmente). Validado con pruebas funcionales de los 4 escenarios (catálogo inexistente, todo reciente, desactualizado+No, desactualizado+Cancelar).

### 2. Recuperación de código proveedor desde el catálogo cuando la tarifa lo trae vacío
Hallazgo del usuario: `_actualizar_codigo_proveedor_catalogo()` sobrescribía siempre la col B del catálogo con el código proveedor de la tarifa, **incluso cuando venía vacío** — borrando sin darse cuenta un valor ya guardado de una carga anterior. Corregido: cuando la tarifa trae el código vacío para un material que **ya existe** en el catálogo, se conserva el valor existente en el catálogo (no se borra) **y además** se escribe de vuelta en la propia tarifa (col C), para que el Registro Info que se genere después salga más completo. Aviso visual (ventana + log) listando qué materiales estaban incompletos y con qué código se completaron, mismo estilo que el aviso ya existente de "Proveedor Asignado no coincide". Solo aplica a materiales que ya tienen fila en el catálogo (los nuevos no tienen nada previo que recuperar). Validado con prueba funcional real de extremo a extremo, y confirmado en producción con COMESA CANARIAS SL: `DALI 2384 → «GALLETAS»` recuperado y completado correctamente en la tarifa.

### 3. Ventanas de descarga/credenciales centradas
El usuario reportó que la ventana "Descargando catálogo..." (y, se comprobó, también la de credenciales) aparecía sin centrar, arriba a la izquierda de la pantalla -- ninguna de las dos tenía código de centrado. Corregido aplicando el mismo patrón de centrado ya usado en el resto de ventanas del proyecto (`winfo_screenwidth`/`winfo_screenheight`), calculado sobre el tamaño natural del contenido (`winfo_reqwidth`/`reqheight`) para no cortar nada. Confirmado visualmente por el usuario tras la prueba.

**Validación conjunta:** 69/69 tests. Fusionado directamente a `master` en commit `d54711e`. App recompilada y confirmada como **v10.2.18**.

---

## 29. Enganche del catálogo completado en los 2 puntos que faltaban (tarifa_app.py)

Sesión 2026-09-23 (trabajo realizado en una sesión de trabajo paralela, verificado y validado aquí). Se detectó que, de los cuatro caminos de código que pueden dar una tarifa por terminada, solo dos llamaban a `_actualizar_codigo_proveedor_catalogo()`. Faltaban:

1. Dentro del diálogo `VentanaSAPFaltantes` en `_ejecutar_transformacion()` (~línea 3726) — el camino en el que la transformación inicial detecta SAP faltantes y se completan ahí mismo, en el momento, sin llegar a pasar por Acción 4.
2. `_ejecutar_transformacion_inc()`, "Caso 2" (~línea 4747) — cuando se reabre un `.xlsx` ya procesado por Acción 3 (no un `.xls` sin transformar) para completar SAP pendientes.

En ambos se añadió el mismo bloque que ya usaban los otros dos caminos, respetando el guard `_catalogo_reciente_o_confirmar()` (umbral de 30 minutos, sección 28):
```python
if self._catalogo_reciente_o_confirmar():
    self._actualizar_codigo_proveedor_catalogo(_ruta_nueva)
```
Con esto, los **6** puntos de enganche (contando los 2 ya existentes en cada uno de los 2 caminos originales) quedan cubiertos. Verificado: sintaxis correcta, los 6 puntos localizados y confirmados con `grep`, 69/69 tests. Validado en real solo para el camino ya existente (caso limpio de HNOS RODRIGUEZ TRUJILLO SL, sin faltantes ni duplicados) — los 2 puntos nuevos quedan pendientes de una validación real específica (un caso que complete SAP faltantes en la transformación inicial, o que reabra un `.xlsx` "Caso 2").

---

## 30. Bug crítico real: pérdida silenciosa de precio/período en Hoja3 con materiales de varios períodos históricos

Sesión 2026-09-23. El usuario reportó que el RINFO de FRAN Y CHEMI SL no guardaba correctamente precios y períodos para varios artículos, ejemplo CALABACIN — un bug real y grave, no cosmético: el material se quedaba **sin ningún precio ni período para la tarifa nueva, en silencio, sin ningún error visible**.

**Diagnóstico:** el material CALABACIN (10101584) tenía en el Purchasing original **dos períodos históricos** en Hoja3 por centro: uno antiguo ya cerrado (01.11.2025-30.11.2025) y uno vigente que conectaba perfectamente con la tarifa nueva (01.08.2026-31.08.2026 → tarifa desde 01.09.2026, sin ningún hueco real). El log mostraba un aviso `⚠ GAP SIN COBERTURA — 54 registro(s)` con 274 días cada uno — un número que no cuadraba con la realidad de los datos.

Dos bugs relacionados en `_aplicar_actualizaciones_h3_heredados()`:
1. **El real (pérdida de datos):** cuando ningún período histórico solapa con la tarifa nueva, el código exigía `fila_num == fila_mas_rec` para actualizar la fila del período más reciente -- pero `fila_num` es la PRIMERA fila física que el bucle exterior encuentra para esa clave (por el guard de "procesar cada clave una sola vez"), casi nunca la más reciente cuando hay 2+ períodos. La condición era prácticamente siempre falsa, así que la actualización **jamás se aplicaba**: el material se quedaba sin precio ni período para la tarifa nueva.
2. **El del aviso falso:** el cálculo de "días sin cobertura" comparaba contra CADA período histórico (incluidos los ya cerrados hace tiempo), en vez de solo contra el más reciente — generando avisos enormes y falsos (274 días) cuando el período realmente vigente conectaba sin ningún hueco.

**Fix:** escribir siempre en la fila del período más reciente directamente (sin depender de en qué fila estuviera el bucle exterior en ese momento), y comprobar el gap solo contra ese mismo período reciente.

**Validación:** pruebas funcionales reales usando el propio método `_actualizar_h3_heredados()`, replicando exactamente el caso CALABACIN (confirmado: período viejo intacto, período reciente actualizado con precio y fechas correctas de la tarifa nueva, sin aviso falso) y un caso de gap genuino con un solo período (confirmado: sigue detectándose correctamente, sin falsos negativos). 69/69 tests.

**Nota sobre cobertura de tests:** los tests existentes de "gap" (`TC09_GapVigencia`) prueban una función distinta (`_detectar_gaps`, validación de huecos dentro de la propia tarifa) — nunca hubo cobertura automatizada para este escenario concreto (reconciliación de Hoja3 heredada con múltiples períodos históricos contra la tarifa nueva).

Fusionado en la misma sesión. App recompilada y confirmada como **v10.2.20**.

---

## 31. Bloqueo global del catálogo + fix de crash real en la ventana de descarga

Sesión 2026-09-24. Continuación de la validación de la sección 29: al probar el primer punto nuevo del enganche del catálogo, el usuario reprodujo exactamente el escenario sospechado — lanzar Acción 3 justo tras el renombrado de una tarifa, interrumpiendo el paso del catálogo todavía en marcha.

### 1. Bloqueo global de `catalogo-dali-completo.xlsx`
Confirmado como riesgo real: sin protección, dos operaciones que escriben el mismo archivo a la vez (descarga desde Catálogo Asignaciones, regeneración tras Acciones 2/3, actualización de código proveedor al completar una tarifa) pueden chocar y perder datos en silencio. Añadido `self._catalogo_lock` (`threading.Lock`), envolviendo los tres métodos que escriben el archivo — cada uno renombrado a `_impl` con un wrapper fino (`try/finally`) que adquiere el lock al entrar y lo libera siempre al salir. Si una segunda operación choca con otra en marcha, se avisa en el log *"Esperando a que termine otra operación sobre el catálogo..."* y espera su turno en vez de chocar. Verificado que no hay llamadas anidadas entre los tres métodos (evita interbloqueo) y validado con una prueba funcional real de dos hilos concurrentes.

### 2. Crash real: `bad screen distance "28 4"`
Al probar la protección anterior, el usuario topó con un crash real y distinto en la ventana "Descargando catálogo": `tk.Label(..., pady=(28, 4))` — una tupla como opción del **constructor** de un widget Tkinter es inválida; las tuplas de padding asimétrico solo funcionan dentro de `.pack()`/`.grid()`, no como argumento directo al crear el widget. Bug arrastrado desde la sesión del centrado de ventanas (2026-09-23), nunca disparado hasta entonces porque no se había probado con una ventana Tk real. Corregido moviendo el padding a `.pack(pady=(28, 4))`. Validado con Tkinter real en el sandbox (Xvfb): reproducido el error exacto con el código original, confirmado el fix, y probada la ventana completa sin fallos — y confirmado también en real por el usuario, con la secuencia completa (transformación → descarga → actualización del catálogo) sin ningún crash.

**Validación conjunta:** 69/69 tests. Fusionado a `master` en commit `3abb514`.

**Estado de validación de los 2 puntos nuevos del punto 29:** uno confirmado en real (camino limpio con descarga, HNOS RODRIGUEZ TRUJILLO); el otro (completar SAP faltantes en la transformación inicial, con posterior actualización del catálogo) validado en real de forma indirecta al confirmar el fix del crash, pero sin un log dedicado que aísle específicamente ese camino. El bloqueo global en sí queda validado solo con prueba funcional, aún pendiente de un caso real donde dos acciones choquen deliberadamente y se vea el aviso de espera en el log.

---

## 32. Punto 6 del roadmap — primera pasada (boilerplate de ventanas Toplevel)

Sesión 2026-09-24. Primer avance real sobre el último punto del roadmap de limpieza (baja prioridad, sin urgencia). Auditoría previa: 15 clases `tk.Toplevel` repartidas entre `crear_registro_info.py` (10) y `tarifa_dialogs.py` (5), con dos familias de centrado distintas — centrado sobre pantalla (`_ajustar_geometria`, familia de `crear_registro_info.py`) y centrado sobre la ventana padre (`_centrar(parent)`, familia de `tarifa_dialogs.py`).

**Alcance deliberadamente conservador:** solo se extrajo lo que era **idéntico byte a byte** entre varias clases, dejando intactas las geometrías con fórmulas de ancho/alto propias (6 de las 9 métodos `_ajustar_geometria`), para no forzar una firma común que metiera riesgo real de romper la disposición visual de alguna ventana por un beneficio menor.

**Tres funciones nuevas en `tarifa_core.py`** (ya importado por ambos ficheros):
- `cargar_icono_ventana(win)` — sustituye 13 copias idénticas del try/except del icono (assets/logo.ico).
- `centrar_sobre_padre(win, parent)` — sustituye los 4 métodos `_centrar()` byte a byte idénticos de `tarifa_dialogs.py` (Selección XLS, Abrir tarifa edición, Crear proyecto, Selección tarifa RI). El 5º (`VentanaSAPFaltantes`) tiene una geometría genuinamente distinta y se dejó tal cual.
- `ajustar_geometria_pantalla(win, ancho)` — sustituye los 3 métodos `_ajustar_geometria()` idénticos salvo el ancho, en `VentanaTelefonoRINFO`/`VentanaFechasRINFO`/`VentanaCentrosRINFO` (flujo "crear RINFO desde cero").

**Bug propio detectado y corregido antes de entregar** (nunca salió del sandbox): al eliminar los bloques de método y sustituir sus llamadas en el mismo script, procesar primero todas las eliminaciones y luego todas las sustituciones descuadraba las líneas — eliminar un bloque desplaza hacia arriba todo lo que viene después, invalidando los números de línea ya calculados para las sustituciones pendientes. Corregido procesando cada par (eliminación + sustitución) de abajo hacia arriba, completo antes de pasar al siguiente.

**Validación:** con Tkinter real (Xvfb en el sandbox) — reproducidas las fórmulas ORIGINALES de cada método eliminado y comparadas byte a byte contra el resultado de las nuevas funciones compartidas, coincidiendo exactamente (`480x98+400+463` en ambos casos, mismo offset de centrado sobre el padre). 69/69 tests. Validado en real con dos flujos completos "crear RINFO desde cero" sin ningún error (PANADERIA ARTESANAL AMARO S.L. y AMD MAYORISTAS EN FRUTAS Y VERDURAS — fechas, centros y teléfono escritos correctamente en ambos, confirmado directamente en los `.xlsx` generados).

Fusionado a `master` en commit `6825dad`.

**Con esto, el roadmap de limpieza tiene los puntos 1-5 CERRADOS, y el punto 6 con su primera pasada hecha.** Quedan las 6 geometrías con fórmulas propias sin tocar, deliberadamente, por si algún día se quiere ir más allá (bajo mismo criterio: solo merece la pena si en algún momento vuelven a ser realmente idénticas entre sí).

---

## 33. Fix crítico de estabilidad: Tkinter no es thread-safe (posiblemente el fix más importante del proyecto hasta la fecha)

Sesión 2026-09-24 (continuación). El usuario reportó la aplicación **completamente congelada**, sin responder a nada, obligando a matarla desde el Administrador de tareas de Windows — justo después de descargar el catálogo (sección 31) y mientras se actualizaba el código proveedor con un aviso "CASO IRREGULAR" largo.

**Causa raíz real:** `_catalogo_reciente_o_confirmar()` (Acción 4) y `_confirmar_catalogo_actualizado()` (Acciones 2/3) creaban ventanas Tkinter **directamente** (`messagebox`, `tk.Toplevel`, `wait_window`) — pero ambas se llaman desde dentro de `proceso()`, que **siempre** corre en un hilo de fondo (`threading.Thread`). Tkinter no es *thread-safe*: tocar widgets desde un hilo que no es el principal puede colgar la aplicación entera, pero de forma **intermitente** — funciona la mayoría de las veces por pura suerte de sincronización con lo que esté haciendo el hilo principal en ese instante, hasta que un día no funciona y se congela del todo. Esto explica por qué llevaba semanas "funcionando bien" en la mayoría de las pruebas de hoy y de sesiones anteriores, y por qué esta vez no.

**Fix:** nuevo helper `_ejecutar_en_hilo_principal_y_esperar(self, funcion)` — si quien llama ya es el hilo principal, ejecuta `funcion` directamente (evita un interbloqueo real: bloquear el hilo principal esperando un `after()` que solo el propio hilo principal puede procesar sería un punto muerto inmediato). Si quien llama es un hilo de fondo, programa `funcion` con `self.after(0, ...)` para que corra en el hilo principal, y bloquea de forma segura (con `threading.Event`) al hilo que llama hasta que termine, devolviendo su resultado o relanzando su excepción si la hubo. Ambos métodos afectados se envolvieron con el patrón ya establecido `_impl` + wrapper fino.

**Segundo hallazgo relacionado**, a raíz de que el usuario notó que "la barra de progreso se pone en verde antes de terminar de actualizar el catálogo completo": `log_consola()` **también** tocaba el widget de texto directamente desde hilos de fondo — con diferencia el punto que más veces toca Tkinter fuera del hilo principal en todo el proyecto (una vez por cada línea de log). Corregido con marshaling **no bloqueante** (fire-and-forget vía `self.after(0, ...)`, sin esperar el resultado — bloquear en cada línea de log sería carísimo) para que el log se procese en el mismo orden relativo que otras actualizaciones de UI ya correctas, como la barra de progreso.

**Validación:** con un `mainloop()` de Tk real (Xvfb en el sandbox) — confirmado que un hilo de fondo tocando Tkinter de verdad (creando un widget) se marshala correctamente al hilo principal y recibe el resultado; confirmada la propagación de excepciones; confirmado que 5 líneas de log seguidas de una actualización de barra, todas lanzadas desde el mismo hilo de fondo, se procesan en el **orden exacto** en que se llamaron. 69/69 tests. Validado en real repitiendo el caso exacto que se había congelado (LIBRERIA Y PAPELERIA ANDROMEDA SL, mismo aviso CASO IRREGULAR de 38 líneas) sin ningún problema, seguido de un flujo completo "crear RINFO desde cero" sin errores.

Fusionado a `master` en commit `b815fd4`. App recompilada y confirmada como **v10.2.22**.

---

## 34. Detección de códigos SAP duplicados (dos capas) + sincronización del catálogo + reintentos de descarga

Sesión 2026-09-25. El usuario pidió, tras un caso real de SAP duplicados, un control que evite que dos códigos DALI distintos acaben compartiendo el mismo código SAP.

### 1. Detección en el momento de teclear (`VentanaSAPFaltantes`)
Antes de guardar, comprueba si algún SAP recién introducido ya está asignado a otro DALI — en el LISTADO existente o dentro del propio lote que se está guardando. Si hay conflicto, nueva ventana `VentanaResolverSAPDuplicados`: una casilla por cada DALI candidato ("recién introducido" o "ya en el listado, fila N"); el usuario marca cuál(es) borrar, debe quedar al menos uno sin marcar por grupo. Se muestra de forma *thread-safe* (mismo patrón `self.after()` + `threading.Event` que el fix de la sección 33, ya que `_tarea_guardar()` corre en un hilo de fondo).

### 2. Auditoría completa (Acción 3)
A petición explícita del usuario ("el problema no es el código DALI, es que existan códigos SAP iguales para diferentes códigos DALI"): nuevo método `_auditar_sap_duplicados_listado()`, enganchado en `_cruzar_y_escribir_listado_sap` justo antes de proteger/guardar — hace un barrido de **todo** el LISTADO ya actualizado, no solo lo que se teclea en ese momento, reutilizando la misma ventana de resolución.

### 3. Bug real encontrado tras la sección 1: el catálogo quedaba desincronizado
El usuario detectó que, tras resolver un duplicado, `LISTADO CODIGOS DALI - SAP.xlsx` quedaba correcto pero `catalogo-dali-completo.xlsx` seguía mostrando el SAP viejo para el DALI ajeno a la tarifa en curso — el enganche normal del catálogo (`_actualizar_codigo_proveedor_catalogo`) solo toca los DALI de la tarifa activa, nunca tocaría un DALI resuelto por casualidad en la misma sesión pero de otro proveedor. Nuevo método ligero `_refrescar_sap_catalogo_para_dali()`, pasado como callback a `VentanaSAPFaltantes`, que actualiza solo las columnas J-S (datos de origen SAP) para los DALI concretos cuyo SAP se acaba de borrar — sin regenerar el catálogo entero. (La vía de la Acción 3 ya se cubre sola, porque la regeneración automática del catálogo corre justo después, leyendo el LISTADO ya corregido).

**Dos bugs propios detectados y corregidos durante la validación** (ninguno llegó a producción):
- `UnboundLocalError: _dali_limpiados` — la variable solo se creaba dentro del `if conflictos:`; si no había ningún duplicado, el código que la referenciaba (el nuevo callback) crasheaba. Corregido inicializándola siempre a `[]` antes del condicional. **Este sí llegó a producción una vez** (ver captura del usuario) y se corrigió en el momento.
- `ws.cell(row, col, value=None)` de openpyxl **no borra** una celda existente (ya nos había pasado antes en `_sincronizar_um_hoja3`) — hace falta `.cell(row, col).value = None` en dos pasos. Detectado con trazas paso a paso cuando la prueba de extremo a extremo mostraba el valor viejo persistiendo pese al log de "refrescado correctamente".

### 4. Reintentos automáticos en la descarga del catálogo
Error real (`[WinError 5] Acceso denegado`) al guardar el catálogo descargado — bloqueo transitorio de Windows (OneDrive, antivirus, o el archivo abierto en Excel). `os.replace()` ahora reintenta hasta 5 veces con 1s de espera antes de dar el guardado por fallido.

**Validación:** todo probado con Xvfb de extremo a extremo (Tkinter real, no solo lógica aislada) — conflicto detectado y resuelto en ambas ramas posibles, catálogo sincronizado correctamente tras confirmarlo con una traza paso a paso, y los reintentos de descarga probados simulando fallos transitorios y permanentes. 69/69 tests en todos los pasos. Fusionado a `master` en commit `32e4a71`.

---

## 35. Bug crítico: solapamiento PARCIAL con precio distinto acortaba el periodo viejo pero nunca creaba el nuevo

Sesión 2026-09-28. El usuario reportó que en el RINFO de MAHOU, S.A. el DALI 29612 (SAP 10200147, CERVEZA SAN MIGUEL BARRIL 30 LITROS) salía a 19,36 € cuando la tarifa decía 19,83 €, y que tampoco aparecían las fechas introducidas (01.02.2026 → 31.01.2027).

**Causa raíz:** en `_aplicar_actualizaciones_h3_heredados()`, rama "precio distinto", cuando las fechas nuevas solapan PARCIALMENTE con un período heredado (aquí 01.09.2025 → 31.08.2026 @ 19,36), el código acortaba el viejo (J = I_nueva − 1 = 31.01.2026, correcto) pero **nunca creaba el período nuevo**. La creación de filas nuevas para PIR ya existentes depende de `estado.absorbidas_con_tarifa` (paso 9b → `_generar_filas_absorbidas_con_tarifa` → paso 10 `_rellenar_h3_nuevas_filas`), y solo la absorción TOTAL registraba la clave ahí. Hoja 2 no aportaba filas nuevas (los 210 cruces ya existían) y `_rellenar_h3_nuevas_filas` solo escribe filas nuevas de Hoja 2. Sin ningún error: el propio log decía "SAP recibirá la fila nueva en la siguiente importación", promesa que nadie cumplía. Firma en el log: `SOLAPAMIENTO — N registro(s): 0 absorbidos, N parciales` + `0 precio+fechas`, sin líneas `Recuperados/generados` ni `Hoja3 — N fila(s) nueva(s)`.

**Alcance real (reproducido con `_ejecutar()` sobre los ficheros reales del usuario):** antes del fix solo 70 de 200 combinaciones material×centro tenían el precio de tarifa cubriendo el rango nuevo (las de mismo precio); el otro 65% —130 registros, 13 materiales— conservaba el precio viejo.

**Fix:** registrar la clave en `estado.absorbidas_con_tarifa` también en el solapamiento parcial (mismo criterio que CASO A: solo si el material está en `heredados_con_match`). `_generar_filas_absorbidas_con_tarifa` ya servía tal cual: no duplica la fila de Hoja 2 si existe y devuelve la tupla para escribir la de Hoja 3. Texto del log del paso 9b ajustado.

**Validación:**
- Reproducción con los ficheros reales (Purchasing + tarifa MAHOU): tras el fix, 200/200 combinaciones con el precio de tarifa cubriendo todo el rango; 10200147 → 10 filas viejas acortadas a 31.01.2026 @ 19,36 + 10 nuevas 01.02.2026 → 31.01.2027 @ 19,83, sin huecos ni solapes.
- Diff celda a celda ANTES vs. DESPUÉS: Hojas 1, 2 y Escalas idénticas; Hoja 3 con las 201 filas previas idénticas y exactamente 130 filas añadidas, todas con las fechas introducidas. Cero efectos secundarios.
- Nuevo `TC21_SolapamientoParcialPrecioDistinto` (4 tests) que ejecuta el CÓDIGO REAL encadenando 6c → 9b → 10 (se salta solo sin tkinter). Con el código original fallan los 2 tests que cubren el bug; los otros 2 protegen comportamientos vecinos (sin solape: reescribe la fila más reciente, fix de CALABACIN; mismo precio: solo extiende J). Suite: **73/73** en la máquina del usuario.
- Confirmado en real: el log de la regeneración de MAHOU muestra `Recuperados/generados: 130 fila(s)` y `Hoja3 — 130 fila(s) nueva(s) añadida(s), K=201–330, I=01.02.2026 / J=31.01.2027`.

**Impacto en RINFO ya generados:** el fix no corrige lo que ya se hubiera cargado en SAP. Escaneo de los logs históricos disponibles por la firma del bug: FRIT RAVICH SL (25-09, 680 parciales, sin filas nuevas), EMICELA (13-08, 60, sin filas nuevas), MAHOU (28-09, 130, ya regenerado), y dudosos MILO (24-09), AHUMADOS CANARIOS y EMICELA (10-08). Pendiente que el usuario compare precios RINFO vs. tarifa en esos proveedores y confirme si se llegaron a cargar en SAP.

**Caso residual no tratado:** periodo nuevo que empieza en o antes que el viejo pero termina antes de que acabe el viejo → sigue saliendo "cierre inválido evitado" y no se crea fila (arreglarlo exigiría mover el inicio del periodo viejo). Raro; sin cambios.

**Lección de proceso:** los tests originales replican la lógica en vez de importar el módulo real, por eso ningún test detectó este bug ni el de CALABACIN (sección 30). TC21 es el primero que ejercita `GeneradorDesdeExistente` real; conviene ampliar en esa línea cuando aparezcan más bugs de Hoja 3.

Fusionado a `master` en commit `f43866f`. Se decide **mantener la versión v10.2.26** (sin subir número) al compilar también el fix de la sección 36.

---

## 36. Bug: hueco de Hoja 3 por CENTRO — material con precio en unos centros activos y sin precio en otros

Sesión 2026-09-28. El usuario preguntó por qué, al pedir el RINFO para los centros 0810 y 0910, no salían todos los materiales en ambos centros (ejemplo: 10101531 de FRAN Y CHEMI SL, "pero existen más").

**Causa raíz:** Hoja 2 añade una fila por cada material de la tarifa × centro activo, pero Hoja 3 (el precio) dependía de una red de seguridad que razonaba **solo por material**: `_materiales_sin_ninguna_fila_h3 = heredados_con_match − {materiales presentes en periodos_h3}`. Un material con condiciones de precio en 0810 pero ninguna en 0910 sí aparecía en `periodos_h3`, así que no se detectaba y el RINFO salía con la fila de Hoja 2 de 0910 pero sin su condición de precio. El hueco venía ya en el Purchasing original (6 de 268 combinaciones material×centro con Hoja 2 sin Hoja 3); la generación lo heredaba sin rellenarlo.

**Alcance:** FRAN Y CHEMI SL: 6 combinaciones (0810: 10101161, 10101525, 10102134; 0910: 10101462, 10101531, 10102227), todas con precio en la tarifa. MAHOU: 10 combinaciones más (10800192 y 10800193), que además quedaron fuera de mi validación anterior ("200/200") porque el denominador eran solo las combinaciones que ya tenían Hoja 3 (200 de 210).

**Fix:** en el cálculo de estado de Hoja 3 heredada, la comprobación pasa a ser por (material, centro activo): si un material con match en tarifa no tiene ninguna fila en Hoja 3 para un centro activo, se registra `("", "IC00", material, centro)` en `estado.absorbidas_con_tarifa`, con el mismo mecanismo que las secciones 22/35 (paso 9b → paso 10). No duplica Hoja 2 si ya existe y el paso 10 deduplica por clave completa, así que es seguro solaparlo con otros orígenes. Aviso en el log: `N combinación(es) material×centro con tarifa activa sin ninguna fila en Hoja 3 (hueco ya presente en el Purchasing original…) — generando fila nueva: …`.

**Validación:**
- Reproducción con los ficheros reales (Purchasing + tarifa FRAN Y CHEMI): antes 6 combinaciones sin Hoja 3; después 0. Diff celda a celda: Hojas 1, 2 y Escalas idénticas; Hoja 3 con las 447 filas previas idénticas y exactamente 6 nuevas, con precio igual al de la tarifa (p. ej. 10101531 @ 0910 = 3,20, 01.09.2026 → 30.09.2026). MAHOU regenerado con el código nuevo: 140 filas nuevas (130 del fix de la sección 35 + 10 de este), 0 combinaciones sin Hoja 3.
- Tests: `TC21` ampliado con 2 tests (`test_hueco_por_centro_…`), que fallan sin el fix. El helper `_ejecutar_cadena` admite `centros_con_h3` para simular Hoja 2 sin Hoja 3 en un centro. Suite: **75/75** en la máquina del usuario.
- Confirmado en real: el log de la regeneración de FRAN Y CHEMI muestra el aviso de las 6 combinaciones y `Hoja3 — 6 fila(s) nueva(s) añadida(s), K=499–504`. Validación sobre el RINFO real subido, con Hoja 2 × centros como denominador: 262/262 combinaciones tarifa×centro presentes en Hoja 2, 0 sin Hoja 3, 262/262 con periodo que cubre 01.09→30.09.2026 al precio de tarifa, 0 solapes y 0 huecos entre periodos consecutivos.

**Lección de proceso:** al validar precios, el denominador debe ser Hoja 2 × centros (lo que debería existir), no las filas que ya están en Hoja 3.

**Escaneo de RINFO ya generados** (combinaciones en Hoja 2 sin Hoja 3 en la salida; sobreestima porque algunos materiales legítimamente no están en la tarifa de ese fichero): AHUMADOS 27 y 28-07 (85), AHUMADOS TF-GC 06-08 (84), AHUMADOS FV 06-08 (6), AHUMADOS LP 06/07/17-08 (2), EMICELA 13-08 (10); sin huecos: AMD, CAFES CANDELAS, LORMACAN, AMARO, PRO A PRO. Para separar huecos reales de legítimos hace falta la tarifa de cada uno.

Fusionado a `master` en commit `8d4b3b5`. Versión compilada: **v10.2.26** (se mantiene el mismo número; incluye los fixes de las secciones 35 y 36).

---

## 37. Catálogo DALI completo pasa a 22 columnas: "Texto UM (expl.)" y "Precio €"

Sesión 2026-09-28 (continuación). El usuario avisó de que `catalogo-dali-completo.xlsx` cambió de formato en la web Catálogo Asignaciones: se insertaron dos columnas nuevas, M "Texto UM (expl.)" y N "Precio €", entre "Unidad SAP" y "Grupo productos SAP" (que pasan de M–T a O–V). Pidió mantener este formato exacto, con cada valor en su columna.

**Riesgo real:** todo el código que tocaba el catálogo (regeneración, actualización por tarifa, refresco tras resolver duplicados SAP) escribía y leía por POSICIÓN fija, no por nombre de columna. Probado contra el catálogo nuevo con el código de la sección 36: el Grupo de productos acababa en "Texto UM", "AGUAS" en "Precio €", y todo lo demás desplazado.

**Fix — definición canónica en `sap_helpers.py`:**
- `CATALOGO_COLUMNAS`: lista ordenada (clave interna, cabecera exacta, ancho) — única fuente de verdad del layout completo (22 columnas).
- `mapa_columnas_catalogo(cabeceras)`: construye `{clave: nº columna}` leyendo la fila 4 real del Excel — funciona igual con el formato nuevo (22 cols) que con el viejo (20, sin Texto UM ni Precio €) o con columnas añadidas en el futuro.
- `datos_sap_desde_fila_listado(row)`: lector único de una fila del LISTADO (antes estaba copiado en 3 sitios); Texto UM sale de la columna F del LISTADO.
- `es_precio_valido(v)`: número > 0, no bool, finito.

Los tres métodos que tocan el catálogo (regeneración, `_actualizar_codigo_proveedor_catalogo_impl`, `_refrescar_sap_catalogo_para_dali`) pasan a localizar cada columna por cabecera antes de leer o escribir.

**"Texto UM (expl.)":** se copia siempre del LISTADO col F durante la regeneración y el refresco SAP.

**"Precio €" (decisión del usuario, mismo turno):** funciona exactamente como "Código Proveedor" — en `_actualizar_codigo_proveedor_catalogo_impl` se pega el valor de la col F ("PRECIO") de la tarifa final ya preparada para RINFO, y SUSTITUYE el que hubiera. Solo se acepta un número > 0 (`es_precio_valido`): vacío, 0, texto o negativo no borran un precio bueno, y se avisa con los DALI afectados. Solo se toca en las filas cuyo Proveedor Asignado coincide con la tarifa (igual regla que el código proveedor). En la regeneración se conserva solo si el Proveedor Asignado no cambió respecto al catálogo anterior; si cambió, se vacía junto al código proveedor y queda anotado en el `.txt` de casos irregulares (columna nueva "Precio € vaciado") y en el aviso emergente. Catálogo todavía en formato viejo (sin columna Precio €): aviso, no se guarda nada — se incorpora en la siguiente regeneración.

**Validación con los ficheros reales del usuario:**
- Cabeceras del catálogo subido por el usuario == definición canónica, exactamente.
- Regeneración: 30.179 filas × 22 columnas sin ninguna celda fuera de sitio frente a ARTICULOS_DALI/LISTADO; Texto UM rellenado en 6.818 filas; precios sembrados conservados.
- Tarifa real GRUPO KALISE SA (106 DALI): 105 coinciden con el Proveedor Asignado → precio del catálogo == columna F de la tarifa en las 105; un precio previo distinto se sustituyó, uno igual quedó igual, uno en una fila omitida por proveedor quedó intacto; 0 filas ajenas a la tarifa cambiaron.
- Casos límite en columna F (12,5 / vacío / 0 / texto / 3 / −1): solo 12,5 y 3 se aplicaron, el resto conservó el precio anterior con aviso.
- Regeneración desde catálogo viejo de 20 columnas: sube a 22 sin contaminar "Precio €", códigos proveedor conservados correctamente.
- Cambio de Proveedor Asignado en la regeneración: código y precio se vacían juntos, reflejado en el `.txt` y en el log.
- TC22 (7 tests del layout) + test de `es_precio_valido`: **83/83**.
- Confirmado en real por el usuario (log 2026-09-28 20:33): `Precio €: 105 precio(s) pegado(s) desde la tarifa (0 sustituyen un valor anterior distinto)` — el DALI 26997 se omitió correctamente por no tener Proveedor Asignado.

**Pendiente:** la web Catálogo Asignaciones sigue exportando Texto UM y Precio € vacíos (30.196 filas), y su descarga sobrescribe el fichero local completo — pendiente de que el usuario confirme si el precio se guardará también ahí, o si vive solo en la app de escritorio (de eso depende si hay que proteger el precio local frente a una descarga).

Fusionado a `master`. Versión compilada: **v10.2.28**.

---

## 38. SAT amplía el Purchasing Info Records con una columna nueva (col AA "Unidad medida pedido")

Sesión 2026-09-29. El usuario avisó de que SAT cambió/amplió el Purchasing Info Records y pidió revisar antes de poner el fichero nuevo como base para los RINFO "sin Purchasing" (RECURSOS/Purchasing Info Records.xlsx, usado por `CreadorRINFO`/`copiar_template_registro_info()`).

**Comparación de los dos ficheros:** único cambio estructural real — Hoja 3 (`Det.precios p.registros info`) pasa de 26 a 27 columnas, con la nueva **AA "Unidad medida pedido"** añadida al final, sin desplazar ninguna otra. Hojas 1, 2 y Escalas quedan idénticas byte a byte en sus cabeceras.

**Bug encontrado:** la mayor parte del código lee el número de columnas de forma dinámica (`ws.max_column`), pero `ESTRUCTURA_HOJAS` (dict hardcoded usado por `aplicar_formato_fila()` para saber cuántas columnas formatear en las filas nuevas) tenía la lista de cabeceras de Hoja 3 fija en 26 elementos. Con el Purchasing nuevo, cualquier fila genuinamente nueva de Hoja 3 (más allá de las que ya trajera el propio Purchasing original) se habría quedado sin ningún formato en la columna AA — sin el relleno azul claro, sin negrita, sin borde — mientras el resto de la fila sí lo tendría. Afecta a las dos rutas de generación (con y sin Purchasing existente), ya que ambas comparten `aplicar_formato_fila()`.

**Fix:** añadida "Unidad medida pedido" como cabecera 27 en `ESTRUCTURA_HOJAS["Det.precios p.registros info"]["cabeceras"]`.

**Validación:**
- Prueba aislada y decisiva sobre `aplicar_formato_fila()` con una fila destino genuinamente en blanco: sin el fix, columna AA sale con `fill=00000000`, sin negrita, sin borde; con el fix, idéntica al resto de la fila.
- RINFO completo generado desde cero (`CreadorRINFO`) con el Purchasing nuevo como plantilla: 27 columnas, sin errores. 83/83 tests (sin cambios en el número, es el mismo fix de un dato interno).
- Confirmado en real por el usuario: RINFO ANI ARMAS RIVERO SL, generado sin Purchasing tras instalar el fix — las 11 filas de artículos en Hoja 3 tienen la columna AA con el mismo fondo, negrita y borde que el resto de columnas (valor vacío, correcto, ese dato no lo trae la tarifa).

Fusionado a `master`. Versión compilada: **v10.2.30**.

---

## Archivos de referencia usados en las pruebas reales (proveedor "Ahumados Canarios")
- `AHUMADOS_CANARIOS__S_A___21-07-2026_.xlsx` — tarifa de proveedor de prueba.

## Caso de referencia para el bug de la sección 30 (proveedor "FRAN Y CHEMI SL")
- Material CALABACIN, código SAP `10101584`, registro info `5300000640`, centros 0810/0910 — dos períodos históricos en Hoja3 (01.11.2025-30.11.2025 y 01.08.2026-31.08.2026), tarifa nueva desde 01.09.2026. Caso mínimo para reproducir: cualquier material con 2+ períodos en Hoja3 donde ninguno solape con la tarifa activa.

## 39. Materiales con registro info marcado para borrado en el Purchasing: dos intentos fallidos y la solución real (manual, fuera de la app)

Sesión 2026-09-30. El usuario detectó que un material (MELOCOTON, SAP 10101624, registro 5300006192, proveedor FRAN Y CHEMI SL) con precio correcto en Hoja 3 (1,60 €, 01.08→30.09.2026) no aparecía disponible al crear un pedido en SAP hoy. Investigado con los ficheros reales: el Purchasing Info Records traía ese registro marcado para borrado (H1 col AJ "Registro info marcado para borrado" = X, H2 col AN "Organización de compras marcada para borrado" = X en ambos centros 0810/0910) — SAP excluye de la determinación de precios un registro en ese estado, aunque su condición de precio siga siendo técnicamente correcta.

**Dos intentos probados en real, con carga real a SAP, y AMBOS fallaron:**
1. Limpiar AJ/AN (dejar en blanco) y marcar Nueva (AK en H1) conservando el número real de SAP en col A → SAP respondió **"Ya existe el registro Info de compras NNNNNNNNNN"** (intenta CREAR con un ID que ya existe).
2. Renumerar el registro ENTERO como si fuera genuinamente nuevo — correlativo provisional en H1+H2+H3 (mismo mecanismo que un material insertado por primera vez), con AK/AO/Z marcados Nueva en las tres hojas → **el mismo error**, con el número REAL antiguo, aunque el Excel ya no lo contuviera en ningún sitio. Esto fue lo decisivo: confirma que SAP compara por **proveedor + material + organización de compras**, no por el número que lleve la fila del Excel — es una regla de unicidad de sus datos maestros que ninguna carga masiva puede sortear con ningún truco de numeración.

**Confirmado también por el propio usuario en SAP:** al intentar crear manualmente un registro nuevo desde la web (Gestión de registros info de compras → Crear) para el mismo proveedor+material+centro, el formulario pedía un campo desconocido ("Clave de clasificación") y no se pudo completar — así que ni siquiera la vía manual de "crear nuevo" es sencilla.

**La solución real la encontró el usuario**: transacción SAP **"Marcar reg.info de compra para borrar"** (pantalla de acceso: Registro info + Proveedor + Material + Centro, con dos casillas "Regtro-Info completo" e "Ind.borr.dat.org.co." bajo "Peticiones de borrado") — desmarcando esas dos casillas para cada centro afectado, reactiva el registro de verdad. Confirmado con una descarga posterior del Purchasing: la cabecera y el centro 0810 quedaron limpios tras usarla.

**Fix final en la app** (revertidos los dos intentos anteriores): `_avisar_materiales_reactivados_bloqueados()` solo limpia AJ/AN (sin marcar Nueva en ningún sitio — probado que no ayuda y si se hace en H1 genera el error) y avisa con claridad. El aviso ahora tiene dos capas:
- Log: `⚠ N material(es) con tarifa activa que el Purchasing traía marcado(s) para borrado (Registro info completo): <lista con código y registro antiguo>. Confirmado en real que SAP rechaza tanto reactivar el número existente como crear uno nuevo...`
- **Ventana nueva `VentanaMaterialesParaReactivar`**, mostrada justo después de la confirmación de "Archivo generado correctamente": lista de materiales/registros afectados (código, descripción, registro antiguo), los 4 pasos para reactivarlos en SAP, y una captura de pantalla de la propia transacción SAP embebida en base64 (mismo mecanismo que el logo SAP del header) como referencia visual — aportada por el propio usuario.

**Validación:** reproducción completa con los ficheros reales en las tres fases (limpiar+marcar H1, renumerar completo, versión final de solo aviso) confirmando en cada una el estado exacto de H1/H2/H3 (col A, AJ/AN, AK/AO/Z). Ventana nueva renderizada y capturada con Xvfb, confirmando que el texto, la lista y la imagen embebida se muestran correctamente. 83 tests (sin cambios de número — el fix final es de UI/aviso, no de lógica de datos).

**Efecto colateral menor detectado y corregido en la misma sesión:** SAT amplió también el Purchasing Info Records con una columna nueva en Hoja 3, **AA "Unidad medida pedido"**, al final, sin desplazar nada más. `ESTRUCTURA_HOJAS` (usada por `aplicar_formato_fila()` para saber cuántas columnas formatear en filas nuevas) seguía con 26 cabeceras — las filas nuevas se habrían quedado sin formato en esa columna. Añadida como cabecera 27. Validado con una prueba aislada y decisiva (fila destino en blanco, con y sin el fix) y confirmado en real con el RINFO de ANI ARMAS RIVERO SL (generado sin Purchasing): la columna AA de las filas nuevas sale con el mismo formato que el resto.

Fusionado directamente a `master` (commits `96efb65` y `feabf9a`). Versión compilada: **v10.2.32**.

---

## Caso de referencia para la sección 31 (proveedor "HNOS RODRIGUEZ TRUJILLO SL")
- Tarifa `365.xls`, cód. proveedor 365. Usado repetidamente para probar el enganche del catálogo en la transformación inicial (con y sin SAP faltantes) y el bloqueo global.
- `Purchasing_Info_Records__2_.xlsx` — Purchasing Info Records base (⚠️ requiere el parche `_parchear_column_dimension()` / `_restaurar_column_dimension()` del propio proyecto para poder abrirse con openpyxl "a pelo" — tiene una peculiaridad OOXML de exportación SAP que openpyxl no soporta sin ese parche).
- Material de prueba clave: código SAP `10101839`, precio 11,95€/kg, sirvió para destapar el bug del punto 4.
- **Segunda tanda (confirmación del troceo del paso 6c):** `log_20260727_142600.txt` + `RINFO AHUMADOS CANARIOS, S.A. [27-07-2026].xlsx`, mismo Purchasing base. Confirmó que `_actualizar_h3_heredados()` reproduce exactamente el comportamiento validado antes: 10101839 con 10 filas nuevas en H3 (11,95€); 10101583/10100148/10101136 con 20 filas en H3 cada uno (10 históricas sin tocar + 10 nuevas por absorción total) y 10 filas en H2 sin duplicar.

## Caso de referencia para la sección 35 (proveedor "MAHOU, S.A.")
- Material 10200147 (DALI 29612), registro info 5300000186, 10 centros, período heredado 01.09.2025→31.08.2026 @ 19,36; tarifa 28-09-2026 a 19,83 con fechas 01.02.2026→31.01.2027. Caso mínimo: cualquier material con precio distinto cuyo período nuevo solape parcialmente con el heredado (TC21).

## Caso de referencia para la sección 36 (proveedor "FRAN Y CHEMI SL")
- Material 10101531, registro info 5300000621: Hoja 2 en 0810 y 0910, Hoja 3 solo en 0810 en el Purchasing original; tarifa 3,20 € para 01.09.2026→30.09.2026. Caso mínimo: TC21 `test_hueco_por_centro_…` (Hoja 2 en dos centros, Hoja 3 en uno).

## 40. VentanaMaterialesParaReactivar: ventana más grande + centros concretos afectados

Sesión 2026-09-30 (continuación). Primera vez que `VentanaMaterialesParaReactivar` (sección 39) apareció en real, en la carga de AMD MAYORISTAS EN FRUTAS Y VERDURAS: mensaje correcto (MANDARINAS 10102056 y PERA DEVOER 10102104), pero el usuario pidió dos mejoras al verla en pantalla:
1. Ventana demasiado pequeña — texto apretado, imagen de referencia cortada por el scroll.
2. Solo mostraba material + registro, sin decir en qué centro(s) concretos seguía marcado para borrado — un mismo registro puede tener el problema en unos centros y no en otros.

**Fix — tamaño:** de 740×~480 (mínimo 640×480) a 920×760 (mínimo 800×680); `wraplength` de las etiquetas y el ancho de la imagen de referencia escalados a juego (680→850, 640→760).

**Fix — centros:** nuevo diccionario `self._materiales_centros_borrado = {material: [centro1, centro2, ...]}`, poblado en el mismo punto donde `_limpiar_marca_aj_heredados()` ya limpiaba la marca AN de Hoja 2 (se captura el centro de esa misma fila antes de continuar). Se muestran tanto en el log (`centro(s) 0110, 0120`) como en cada línea de la ventana (`(registro 5300000953)   centro(s): 0110, 0120`).

**Validación:** renderizado con Xvfb reproduciendo el caso real (MANDARINAS solo en 0110, PERA DEVOER en 0110 y 0120) — captura confirma el tamaño nuevo, el texto legible, la imagen completa sin necesidad de scroll, y los centros correctos por material. 83 tests (sin cambios en el número, es una mejora de UI).

Fusionado directamente a `master`. Versión compilada: **v10.2.34**.

---

## 41. Octubre a precio viejo por un estado heredado incoherente + aviso explícito cuando un cambio de precio se pierde por la UM

Sesión 2026-10-06. El usuario detectó que el RINFO de FRAN Y CHEMI SL (octubre completo, 01.10→31.10.2026) no traspasaba bien los precios: ESCAROLA 1,80 en tarifa, 1,50 en el RINFO.

**Causa 1 (código — `_aplicar_actualizaciones_h3_heredados`):** la decisión "MISMO PRECIO" comparaba la tarifa solo con el precio del período de MAYOR INICIO de la clave. El Purchasing de ese día traía ~59 materiales × 2 centros con un estado incoherente de SAP (`01.10→30.10 @precio viejo` + `31.10→31.10 @precio de tarifa`): el "más reciente" (31.10) igualaba la tarifa, se daba el precio por bueno ("260 solo fin-validez") y la otra fila, a precio viejo, cubría 30 de los 31 días de octubre sin ningún aviso. El origen más probable del estado roto fue haber subido el mismo RINFO 4-5 veces el 30-09 (pruebas del melocotón): los logs de SAP de esa fecha muestran que cada carga creó una condición nueva de la escarola solapada con la anterior (147241 → 147484 → 147600 → 147718) y que SAP recorta fechas por su cuenta.

**Fix 1:** nuevo `_solapa_con_precio_distinto()` — "mismo precio" exige además que NINGÚN período heredado que solape con [I_nueva, J_nueva] tenga otro precio; si lo hay se trata como cambio de precio (los períodos solapados quedan absorbidos y se genera la fila correcta) y se emite un WARN "períodos heredados SOLAPADOS…". Con datos sanos la salida es idéntica byte a byte. Tests `TC23` (4).

**Lección de validación:** comprobar solapes/huecos NO bastaba — el RINFO roto tenía 0 solapes y 0 huecos. La validación correcta es día a día: el precio vigente de cada día del periodo debe ser el de la tarifa.

**Validación en real:** RINFO regenerado con el fix → aviso con los 59 materiales, 236 filas absorbidas, 118 filas nuevas (`01.10→31.10`, Z=X, precio = tarifa) y 262/262 combinaciones correctas. Tras cargarlo UNA vez en SAP y descargar un Purchasing nuevo: **260/262 correctas**; la escarola quedó perfecta (nov-25, ago 1,40, sep 1,50, **oct 01→31 a 1,80**) — **confirmado que SAP recorta y elimina las condiciones antiguas que la fila nueva cubre por completo.**

**Causa 2 (hallada en esa validación — COCOS 10101648, los 2 fallos restantes):** el material quedó SIN precio de octubre en SAP (seguía `01.11.2025→30.11.2025 @1,20` tras cargar `01.10→31.10 @1,35`). Su condición antigua está en KG mientras la cabecera del registro y el LISTADO dicen PI (pieza). `_sincronizar_um_hoja3` corrige la UM en la misma fila que `_marcar_modificado` ya había marcado Z=X por cambio de fecha/precio; como Z=X + cambio de UM en la misma fila hace que SAP rechace la carga ("adecúe la UM-base…"), prioriza que la carga no reviente y fuerza Z a vacío. Su docstring afirmaba que "SAP recoge igualmente el cambio de fecha y de UM al modificar el registro existente": **cierto para la UM, FALSO para fecha y precio** — SAP ignora en silencio una modificación de validez/precio sin Z=X. El único rastro era la línea de resumen "2 con Z forzada a vacío por conflicto fecha+UM".

**Fix 2 (solo aviso, sin cambiar la lógica):** `_sincronizar_um_hoja3` registra cada fila afectada y emite un WARN agregado explícito (material, centros, fechas y precio que NO se registrarán, y que hay que crear la condición a mano en SAP o corregir la UM antes); mensaje por fila y docstring corregidos. Tests `TC24` (3; el principal falla sin el aviso). No se ha cambiado el comportamiento porque el error de unidad base de SAP es real y no se puede probar sin una carga real.

**Resolución de COCOS:** el usuario creó un artículo nuevo para el coco (SAP 10103986 "COCO ENTERO - PI", DALI 1965, a 1,35, unidad PI) y el RINFO regenerado lo inserta como material genuinamente nuevo (registro provisional, Nueva=X en H1/H2/H3, unidad PI/PI coherente en cabecera y condición); el coco viejo (10101648) va de baja. Validado día a día: 262/262. No existía ya en el Purchasing, así que no hay riesgo de "ya existe".

**Causa 3 (hallada al revisar ese log — MELOCOTON 10101624, dado de baja por decisión del usuario por no estar en la tarifa de octubre):** el RINFO llevaba un período con **fin anterior al inicio** (`01.10.2026 → 30.09.2026`). En SAP el material tenía dos períodos (`01.08→30.09` y `01.10→30.10`); `_propagar_borrado_h1_a_h2_h3` cierra la fecha fin del período más reciente al día anterior al inicio de la nueva tarifa (30.09) si está vigente, **sin comprobar que ese período empezara antes** de esa fecha.

**Fix 3:** nueva función pura `debe_cerrar_fin_validez()` en `sap_helpers.py` — solo se cierra si el período está vigente Y empieza el día del cierre o antes; si empieza después no se toca (de todas formas va con Y=X). Reproducido con los ficheros reales simulando la decisión de baja: sin el fix `01.10.2026→30.09.2026`, con el fix `01.10.2026→30.10.2026` y "0 col J cerrada(s)"; solo cambian esas 2 filas, el resto del RINFO es idéntico. Tests `TC25` (5: 3 de la función y 2 con los métodos reales `_limpiar_marca_aj_heredados` + `_propagar_borrado_h1_a_h2_h3`, con fechas lejanas para no depender del día). Suite: **95**.

**Observaciones sin cambiar (decisión del usuario):**
- La baja automática marca para borrado cualquier material sin tarifa cuyo último período ya venció (hoy: PARAGUAYO 10102184, CAQUIS, CARAMBOLA, COCOS) y el usuario puede dar de baja los vigentes (hoy: MELOCOTON). En productos de temporada (melocotón, paraguayo) eso crea exactamente la trampa de la sección 39: cuando vuelvan a la tarifa estarán marcados para borrado en SAP y habrá que reactivarlos a mano con "Marcar reg.info de compra para borrar". Alternativa a valorar: dejar caducar la validez sin marcar borrado en los productos que se espera que vuelvan.
- `_limpiar_marca_aj_heredados` decide "vigente o vencido" con el fin de validez de la PRIMERA fila de Hoja 3 de cada material (no la máxima); como SAP no exporta en orden cronológico, un material con una fila antigua primero y otra vigente después podría darse de baja automáticamente sin preguntar. Hoy no afectó a ninguna baja. Corregirlo (usar la fecha fin máxima) solo haría que se pregunte en más casos.

**Pendiente:** (1) decidir si se implementa la solución de fondo para filas con cambio de precio/fecha + cambio de UM (fila nueva aparte, hipótesis sin verificar en SAP; hoy solo hay aviso explícito); (2) valorar los dos puntos de "Observaciones"; (3) cargar el RINFO una sola vez y validar el Purchasing resultante día a día.

Fusionado a `master` (commit `d0ad26b` para el fix 1). Versión compilada: **v10.2.36** (incluye fix 1, aviso del fix 2 y fix 3).

---

## Caso de referencia para la sección 39 (proveedor "FRAN Y CHEMI SL")
- Material 10101624 (MELOCOTON), registro 5300006192, centros 0810/0910: Purchasing lo traía marcado para borrado pese a tener precio correcto en la tarifa. Solo se resuelve con la transacción SAP "Marcar reg.info de compra para borrar" (manual, fuera de la app). Otros dos casos idénticos en la misma tarifa: 10101585 (registro 5300000641) y 10102184 (registro 5300006196).

## Caso de referencia para la sección 41 (proveedor "FRAN Y CHEMI SL")
- ESCAROLA 10101544 (registro 5300000627), tarifa octubre a 1,80: el Purchasing traía `31.10→31.10 @1,80` + `01.10→30.10 @1,50`. COCOS 10101648 (registro 5300000654): condición antigua en KG, cabecera/LISTADO en PI, tarifa octubre 1,35 por pieza. MELOCOTON 10101624 (registro 5300006192): períodos `01.08→30.09` y `01.10→30.10` en SAP, dado de baja por decisión del usuario. Tests mínimos: TC23 (escarola), TC24 (COCOS) y TC25 (MELOCOTON).

## Técnica usada para probar métodos de `GeneradorDesdeExistente` fuera de la UI (sin lanzar la app completa)
```python
import crear_registro_info as cri
gen = object.__new__(cri.GeneradorDesdeExistente)  # salta __init__ (que lanza un hilo real)

class ParentFake:
    def after(self, ms, fn): pass
    def withdraw(self): pass

gen._parent = ParentFake()
gen._app_root = None
gen._ruta_tarifa = ruta_tarifa
gen._nombre_archivo = os.path.basename(ruta_tarifa)
gen._nombre_sap = "..."
gen._codigo_sap = "..."
gen._purchasing = {"ruta": ruta_purchasing, "archivo": "..."}
gen._carpeta_tarifas = carpeta_tarifas
gen._telefono_proveedor = "..."
gen._fecha_inicio = "..."
gen._fecha_fin = "..."
gen._centros_activos = [(cod, cod) for cod in [...]]

gen._ejecutar()   # síncrono, sin hilo — los errores salen al momento como traceback normal
```
(`self._audit` y `self._fila_h1_por_codigo` se autoinicializan dentro de `_ejecutar`, no hace falta prepararlos.)

---

## 42. Repositorio en GitHub: seguridad y etapa 1 de la reorganización (9 nov 2026) — PR en GitHub

El proyecto se subió a GitHub (`controlpedidosprincesscanarias-coder/consola-registro-info`, un solo commit «Add files via upload»). Se sigue el mismo método que con el Organizador Princess: por etapas, una rama y un PR por etapa, sin push a `main` sin confirmación de Víctor.

### Seguridad (PR #1, fusionado)
- `tarifa_config.json` se subió por error con el correo y la contraseña de Catálogo Asignaciones en claro (el `.gitignore` local no se subió). Se quita del repositorio y se añade `.gitignore` (mismo criterio que el local, más `TARIFA_Finder.spec`, `Thumbs.db` y `__pycache__`). Víctor confirma que esa contraseña no sirve para nada.
- Nuevo `credenciales.py` (sin tkinter): la contraseña se guarda cifrada con DPAPI de Windows (`dali_app_password_cifrada`, prefijo `dpapi1:`), como el Organizador. Una antigua en claro se cifra y se borra al arrancar; si no se puede cifrar, no se guarda. `test_credenciales.py` (8 pruebas).

### Etapa 1 (base del repositorio)
- **Bug de pruebas** (commit propio): TC25 no tenía `@unittest.skipIf(_CRI is None, …)` y en un Python sin tkinter daba 2 errores en vez de saltarse.
- **Importaciones sin usar** (pyflakes, unas 70) quitadas. pyflakes no encuentra ningún nombre sin definir (solo una anotación de tipo en texto en `sap_helpers.py`). Variables asignadas y sin usar revisadas: son restos inofensivos (`val_col5`, `mejor_tipo`, `df1`, `COL_*`…), sin errores escondidos.
- **Archivos muertos fuera:** `organizar_dali_metodo.py` (fragmento de integración de la antigua Acción 1, ya integrada) y `TARIFA_Finder.spec` (lo genera PyInstaller con rutas de G:).
- **Documentación nueva:** `CLAUDE.md` (reglas para Claude Code), `AUTORIA.md` y README reescrito (el anterior era de «TARIFA Finder v2» y tenía los acentos rotos).
- **Pruebas:** 103 en verde con tkinter (`xvfb-run` en Linux); las 18 que importan `crear_registro_info` se saltan sin tkinter.

---

## 43. Etapa 2: código en carpetas por temas (10 nov 2026) — PR en GitHub

- **Saltos de línea** (commit propio): los archivos mezclaban CRLF y LF. Ahora `.gitattributes` (`* text=auto`, `.bat` con CRLF): el repositorio guarda LF y Git pone CRLF al sacarlos en Windows. Solo cambian los saltos de línea.
- **Movimiento** (commit de solo renombrados, para que `git log --follow` siga funcionando): `app/comun` (`tarifa_constants`, `tarifa_core`, `credenciales`, `sap_helpers`), `app/ui` (`tarifa_dialogs`), `app/rinfo` (`crear_registro_info`), `app/tarifas` (`transformar_tarifa`), `app/acciones`, y las pruebas a `tests/`. `tarifa_app.py` se queda en la raíz (lo usan el `.bat`, `_bat_helper.py` y `build_exe.py`).
- **Importaciones** con la ruta completa (`from app.comun.tarifa_core import …`).
- **Rutas:** `resource_path()` y `get_config_path()` usaban la carpeta del propio módulo (`__file__`). Ahora usan `RAIZ_PROYECTO` (`tarifa_constants.py`), así que `assets/` y `tarifa_config.json` siguen en la raíz, junto a `tarifa_app.py`.
- **`build_exe.py`:** ya no añade los `.py` como datos. Incluye el paquete con `--collect-submodules app` y solo comprueba que estén todos los módulos.
- **Comprobado:** 103 pruebas (18 se saltan sin tkinter); autotest `python -m app.rinfo.crear_registro_info --test` (9 OK); la app arranca con su logo y lee la configuración de la raíz. Además, compilado con PyInstaller en Linux con el mismo `build_exe.py`: el ejecutable incluye todo `app/` y arranca.
- **Visto de paso (ya pasaba antes de la etapa):** en el ejecutable de Linux la cabecera muestra «P» en vez del logo (lo mismo compilando `main`). Hay que confirmar que en el `.exe` de Windows sale el logo.

---

## 44. Etapa 3: limpieza a fondo y arranque más rápido (11 nov 2026) — PR en GitHub

Petición de Víctor: «limpiar bien la aplicación, hacerla más ágil, reorganizar y quitar los archivos que no correspondan».
- **Código muerto (~1.090 líneas)**, comprobado con `vulture` y con `grep` (también en cadenas):
  - en `tarifa_dialogs.py`, `VentanaSeleccionTarifaRI` (~670 líneas) y `VentanaAbrirTarifaEdicion` (~150), que no se abrían desde ningún sitio;
  - `App._ejecutar_transformacion_inc` (~220 líneas, sin llamadas);
  - `escribir_fila_con_formato`, `_sanear_nombre_archivo` y `get_subcarpetas`.
- **Autotest PIR fuera del programa:** el bloque «TESTS EDGE CASES SAP PIR» (~300 líneas, T01–T09, `--test`) vivía al final de `crear_registro_info.py` y se cargaba en cada arranque. Pasa sin cambios a `tests/test_casos_limite_pir.py`. T07–T09 no estaban cubiertos en otro sitio.
- **Imágenes fuera del código (~95 KB de base64):**
  - la captura de `VentanaMaterialesParaReactivar` pasa a `assets/marcar_borrado_sap.png`;
  - el logo SAP, a `assets/logo_sap.png`;
  - el logo de la tarifa en Excel era idéntico a `assets/logo_dali_articulos.png` y ahora se lee de ahí;
  - sobraba el respaldo en base64 del logo de la cabecera.
- **Assets que no se usaban** (~4,4 MB): `corona.png`, «logo princess.png», «logo blanco horizontal.png» y el `.jpg` del logo negativo. `build_exe.py` mete todos los assets necesarios y se para si falta alguno.
- **Bug real** (commit propio): al `.exe` le faltaba el *hidden import* `PIL._tkinter_finder`. ImageTk fallaba dentro del ejecutable y la cabecera mostraba «P» en vez del logo. Visto compilando con PyInstaller en Linux, también con `main`. **Confirmar en Windows** que ahora sale el logo (seguramente ya pasaba en el `.exe` de Windows).
- **Arranque:** pandas solo se carga al usar la Acción 1. Importar `tarifa_app.py` pasa de 0,29 s a 0,14 s.
- **Comprobado:** 104 pruebas; el logo se inserta en una hoja de tarifa; la app arranca; el ejecutable compilado muestra el logo.
- **Propuesta, sin hacer:** el `.exe` es `--onefile` (48 MB) y en cada arranque descomprime todo en una carpeta temporal, que es lo que más tarda. Con `--onedir` arrancaría bastante más rápido, pero se distribuiría una carpeta en vez de un solo archivo. Lo decide Víctor.

---

## 45. Etapa 4 (primera parte): crear_registro_info.py repartido en módulos (11 nov 2026) — PR en GitHub

`crear_registro_info.py` (~9.000 líneas) pasa a cinco módulos de `app/rinfo/`, sin cambiar el código:

| Módulo | Contenido | Líneas |
|---|---|---|
| `comun_rinfo.py` | constantes, `_log`, `_LogAgrupado`, `_safe_after` y funciones auxiliares (plantilla Purchasing, formato de filas, similitud de nombres, búsqueda de tarifas, Purchasing y listados) | ~1.100 |
| `ventanas_rinfo.py` | 9 ventanas que no dependen de otras | ~1.500 |
| `creador_rinfo.py` | `CreadorRINFO` | ~1.100 |
| `generador_existente.py` | `GeneradorDesdeExistente` y `_EstadoIndicesH3` | ~4.300 |
| `crear_registro_info.py` | las 3 ventanas del flujo, `iniciar_crear_registro_info` y la reexportación de todos los nombres | ~900 |

- **Hecho con un script** (cada definición con sus comentarios de encima). Las 55 definiciones se han comparado por AST con las originales: son idénticas. No hay importaciones circulares (`comun_rinfo` ← `ventanas_rinfo` ← `creador_rinfo` / `generador_existente` ← `crear_registro_info`) ni variables `global`.
- **Pruebas:** parcheaban `_log`, `_safe_after` y `_ventana_sin_tarifa` en `crear_registro_info`. Ahora cada módulo tiene su propia referencia, así que se parchean en todos con `_parchear_rinfo()`. Sin ese cambio, 4 pruebas abrían la ventana real de «materiales sin tarifa». 104 pruebas OK y la app arranca.
- **Siguiente:** trocear `tarifa_app.py` (clase `App`, ~5.200 líneas) por acciones, con clases *mixin* (`App(AccionDali, AccionSap, …)`), siguiendo el mismo método.

---

## 46. Etapa 4 (segunda parte): la clase App de tarifa_app.py repartida en mixins (11 nov 2026) — PR en GitHub

`tarifa_app.py` (~5.300 líneas) queda en ~200: la clase `App` hereda de seis *mixins* de `app/principal/`, sin cambiar el código.

| Mixin (módulo) | Contenido | Métodos | Líneas |
|---|---|---|---|
| `InterfazApp` (`interfaz.py`) | ventana, consola, botones, acciones sencillas | 23 | ~880 |
| `AccionOrganizarDali` | Acción 1 | 1 | ~380 |
| `AccionListadoDali` | Acción 2 | 1 | ~350 |
| `AccionListadoSap` | Acción 3 | 6 | ~830 |
| `CatalogoDali` | catálogo DALI completo y Catálogo Asignaciones (`DALI_APP_API_BASE`) | 15 | ~1.420 |
| `AccionesTarifa` | Acciones 4 y 6 | 11 | ~1.600 |

- En `App` quedan `__init__`, `_build_header` y `_build_footer`, porque usan `APP_VERSION`, que sigue en `tarifa_app.py` (lo leen `_bat_helper.py` y el `.bat`). Comprobado que `_bat_helper.py leer` da `v 10.2.36`.
- Los 60 métodos se han comparado por AST con los originales: son idénticos, y ninguno está repetido entre mixins (si lo estuviera, ganaría el primero en silencio).
- **Limpieza de la parte anterior** (commit propio): en `app/rinfo` quedaban importaciones que no se usaban, porque la de `sap_helpers` llevaba `# noqa: E402` y autoflake no la tocó.
- **Comprobado:** 104 pruebas; la app arranca; compilado con PyInstaller, el ejecutable incluye los 6 módulos y muestra el logo.

---

## 47. Compilación en GitHub con versión automática (11 nov 2026) — PR en GitHub

Petición de Víctor: compilar desde GitHub, igual que el Organizador, y que la versión suba sola en cada versión nueva, sin el `.bat`.
- **Workflow** `.github/workflows/compilar.yml` («Compilar Consola Registro Info»), solo a mano, en Windows:
  1. calcula la versión;
  2. instala `compilacion/requisitos.txt`;
  3. compila y lanza todas las pruebas (`unittest discover`);
  4. ejecuta `build_exe.py` y genera `TARIFA_Finder.exe` y su `.sha256`.

  Tiene dos destinos: «solo-descarga» (deja el `.exe` como artefacto) y «borrador» (crea una Release en borrador con `--generate-notes`).
- **Versión automática** (`compilacion/version.py`, 8 pruebas en `tests/test_version.py`):
  - toma la mayor entre `APP_VERSION` y las Releases, +2 (siempre par): la primera será **10.2.38**;
  - se escribe en la copia que se compila y no se sube al repositorio;
  - el campo «version» permite forzar otra, que tiene que ser mayor y par.
- `COMPILAR_TARIFA.bat` se queda para compilar en local, pero ya no hace falta para la versión.
- **Pendiente:** probar la primera ejecución en GitHub, que solo se puede lanzar cuando el workflow esté en `main`. Después, las actualizaciones automáticas a los compañeros (como el Organizador, con `releases/latest`); antes hay que decidir si el repositorio vuelve a ser público.

---

## 48. Actualizaciones automáticas para los compañeros (11 nov 2026) — PR en GitHub

**Validación previa** de la v 10.2.38 compilada en GitHub, con datos reales. Víctor pasó las mismas tarifas con la versión anterior y con la nueva, y se compararon los logs y los Excel celda a celda:
- ORION: tarifa, INC, reinscripción y RINFO nuevo con 2 centros.
- Acciones 1–3 y catálogo DALI.
- OFIPAPEL: alta de SAP faltantes, resolución de duplicados, Acción 3 a mitad de proceso y RINFO nuevo con 3 centros.
- ANDROMEDA: RINFO desde Purchasing existente, con 92 materiales heredados y 3 nuevos.

En todos los casos, **0 diferencias** en los datos. Solo cambian el mensaje del logo (paso 14b), las fechas «Generado el…», la celda seleccionada al abrir y el orden interno de la tabla de estilos.

**Actualización automática.** Cada compañero tiene su copia del `.exe` (respuesta de Víctor).
- `app/comun/actualizaciones.py` (sin tkinter): consulta `releases/latest` sin iniciar sesión (un 404, sin Releases publicadas, no es error), descarga junto al `.exe`, comprueba el SHA-256 (sin `.sha256` no instala), sustituye el `.exe` renombrando el actual a `.anterior.exe` y vuelve atrás si falla, y abre el nuevo con `PYINSTALLER_RESET_ENVIRONMENT=1` y sin las variables `_PYI_*`.
- `app/ui/ventana_actualizacion.py`: aviso con las novedades de GitHub en castellano, «Actualizar ahora» y «Más tarde».
- `InterfazApp._comprobar_actualizaciones` se lanza 3 s después de arrancar. En la consola aparece «Versión al día» o «Nueva versión disponible».
- **Pregunta de Víctor: ¿se verifica el borrado de la versión anterior?** Ahora sí: al arrancar la versión nueva se borra el `.anterior.exe` y la consola lo indica («Versión anterior borrada»). Si la vieja aún se está cerrando, se reintenta cada 10 s, hasta 6 veces; si no, se avisa y se reintenta en el próximo arranque. En el Organizador, el instalador se instala encima, en la misma carpeta y con el mismo identificador, pero el instalador descargado se queda en `%TEMP%\OrganizadorPrincess_actualizacion` hasta la actualización siguiente (propuesto borrarlo).
- 13 pruebas nuevas en `tests/test_actualizaciones.py` (125 en total).
- **Por probar en Windows:** publicar la v10.2.38, compilar y publicar después otra (10.2.40) y abrir la 10.2.38: debe avisar, actualizarse, abrirse sola y borrar el `.anterior.exe`.
- ~~**Pendiente:** el aviso de GitHub «Node.js 20 is deprecated» en el workflow~~ resuelto (sección 58).

---

## 49. Manual de usuario, botón «Manual» y «Consultar el manual» (11 nov 2026) — PR en GitHub

Petición de Víctor: manual de usuario con las mismas pautas que los de Control Pedidos y el Organizador, un botón para abrirlo y una celda de consulta que busque las respuestas en el manual (elegidos: búsqueda en el manual y botón en «Accesos rápidos»).
- **Manual** (`docs/Manual_Usuario_ConsolaRegistroInfo.docx` / `.pdf`, 22 páginas, 9 capítulos): qué es y flujo, ponerla en marcha (un solo `.exe`, sin instalar; `tarifa_config.json` junto a él), ventana principal, preparar el proyecto (qué archivo va en TARIFAS y en RECURSOS), Acciones 1–3 y catálogo, Acciones 4–6 (SAP faltantes y duplicados, [INCOMPLETO]), Registro Info (todas sus ventanas), consulta del manual y actualizaciones, y tablas de problemas frecuentes. Cada dato se comprobó en el código.
- **Fuente** en `docs/manuales_fuente/` (ver su README): 21 capturas reales de la consola en Xvfb con un proyecto de datos inventados y sin red (`capturas.py`), y maquetación con `docx` de Node (`lib.js`, adaptado del Organizador). `lib.js` genera además `manual_indice.json` (apartados con capítulo y texto) y `paginar_indice.py` les pone la página del PDF. Corregido en `lib.js`: el párrafo separador tras una tabla, si caía al final de una página, dejaba una página en blanco antes del capítulo siguiente.
- **En la app:** tarjeta «📘 Manual de usuario» en Accesos rápidos (abre el PDF del propio `.exe`, copiado a `%TEMP%\ConsolaRegistroInfo`) y barra «🔎 Consultar el manual» encima de la consola: con Intro abre `VentanaConsultaManual` con los apartados que mejor encajan (capítulo, página y fragmento), «Abrir el manual» y «Guardar una copia…». La búsqueda (`app/comun/manual.py`) normaliza acentos, quita palabras vacías, usa raíces sencillas («creo»/«crear», «tarifas»/«tarifa») y puntúa más el título y las palabras poco frecuentes. Sin Internet ni claves.
- `build_exe.py`: el PDF y el índice pasan a `ASSETS_NECESARIOS`.
- **Error corregido** (commit propio): las etiquetas «Email» y «Contraseña» de la ventana de Catálogo Asignaciones salían en un color casi invisible sobre el fondo oscuro.
- **Icono de la barra de tareas** (pregunta de Víctor: salía la pluma de Tk): `poner_icono_aplicacion()` (`tarifa_core.py`) pone el `.ico` como icono por defecto de todas las ventanas y sus tamaños grandes con `iconphoto`, que es el que usa la barra de tareas. **Por probar en Windows.**
- **Pregunta de Víctor: ¿el `.exe` tiene que estar siempre en el mismo sitio?** Sí, conviene: la actualización lo sustituye en la misma carpeta y con el mismo nombre, así que el acceso directo del Organizador sigue valiendo. Explicado en el capítulo 2 del manual.
- 6 pruebas nuevas en `tests/test_manual.py` (131 en total), incluida una contra el índice real.

---

## 50. Icono propio de la Consola en la barra de tareas (11 nov 2026) — PR en GitHub

Con la v10.2.42, Víctor vio en la barra de tareas la corona gris, borrosa y sin recuadro, casi igual que la del Organizador. Pidió el icono con el recuadro para distinguirlas.
- `assets/logo.ico` nuevo: corona dorada sobre un cuadrado azul marino de esquinas redondeadas con borde dorado, en 256, 128, 64, 48, 32, 24 y 16 px. Cada tamaño se dibuja aparte a 8x y se reduce; en los pequeños el borde es más grueso y la corona algo mayor, para que se vea nítido.
- Se genera con `compilacion/generar_icono.py` (la corona de partida está en `compilacion/corona_icono.png`, sacada del icono anterior).
- `poner_icono_aplicacion()` pasa también los tamaños 48 y 24 a la barra de tareas, que los usa con el zoom de Windows al 125–150 %.
- **Por probar en Windows:** la barra de tareas, el `.exe` en el Explorador y las ventanas. Windows guarda en caché los iconos: si se sigue viendo el antiguo, quitar el anclado y volver a anclar.
- La actualización de la 10.2.40 a la 10.2.42 funcionó (confirmado por Víctor).

---

## 51. Novedades de la versión al pulsar sobre ella (11 nov 2026) — PR en GitHub

Petición de Víctor: en todas las aplicaciones se pueden ver las notas de la versión pulsando sobre ella (Organizador v4.18.44, móvil v1.0.6, Control Pedidos v12.34.20 y aquí).
- El número de versión de la cabecera (subrayado) y el del pie se pueden pulsar y abren «Novedades de la versión» (`app/ui/ventana_notas_version.py`), con el mismo formato que el aviso de versión nueva (`_notas_legibles`) y un botón «Todas las versiones».
- Las notas son las de la Release de esa versión en GitHub (generadas a partir de los PR). Se piden en segundo plano y sin iniciar sesión con `actualizaciones.notas_de_version()` (`/releases/tags/v…`). Sin conexión, o desde el código (sin Release, 404), se dice y se ofrece la página de versiones.
- Manual: la leyenda 11 de la ventana principal lo explica; PDF e índice regenerados (también en `assets/`).
- 2 pruebas nuevas en `tests/test_actualizaciones.py` (133 en total). Ventana probada en una pantalla virtual con notas simuladas.
- **Por probar en Windows:** pulsar la versión en el `.exe` publicado (la de la Release que se compile).

---

## 52. Etapa 6: la Consola envía sus cambios a Catálogo Asignaciones (12 nov 2026) — PR en GitHub

Petición de Víctor: que el LISTADO DALI–SAP y el catálogo completo dejen de ser Excel compartidos y la Consola escriba directamente en la web, **solo con las modificaciones** («que la comunicación sea solo la necesaria»). La parte de la web es Catálogo Asignaciones v1.20.63 (`POST /admin/catalogo/lote`, `GET /admin/catalogo/cambios`, permiso «Escribir desde la Consola Registro Info»). Esto es la fase 1 (doble escritura).
- `app/comun/catalogo_web.py` (sin tkinter): conexión con la web (misma sesión por cookie que la descarga del catálogo), copia de la web en el PC (`%LOCALAPPDATA%\ConsolaRegistroInfo\catalogo_web.json`: la primera vez se descarga entera, después solo lo cambiado), lectura de `catalogo-dali-completo.xlsx` por cabecera, normalización igual que la web (también el redondeo del precio), diferencias por grupo (datos DALI, datos SAP, precio y código de proveedor por proveedor) y envío por lotes de 400.
- `app/principal/publicar_catalogo_web.py`: se lanza tras `_regenerar_catalogo_dali_completo` (Acciones 2 y 3), `_actualizar_codigo_proveedor_catalogo` (Acción 4) y `_refrescar_sap_catalogo_para_dali` (SAP duplicado), en su propio hilo. Si ya hay un envío en marcha, se repite una sola vez al acabar (dos tarifas seguidas = una ventana más, no dos a la vez). Las credenciales son las de siempre (DPAPI).
- `app/ui/ventana_publicar_catalogo.py`: resumen tras simular. Lo editado en la web desde el último envío queda fuera por defecto (casilla para incluirlo) y sigue protegido en los envíos siguientes; los irregulares que la web no aplica (proveedor asignado distinto, código de proveedor distinto) se listan para corregirlos (decisión 5). «Ahora no» no envía nada: la diferencia sigue ahí y vuelve a salir.
- Orden de aplicación: altas y datos DALI → datos SAP → precio y código de proveedor (un artículo nuevo lleva todo en el mismo envío). Un precio o código de proveedor vacío en local no borra el de la web. Nunca se borra nada en la web.
- Decisión de Víctor (12 nov 2026, durante el trabajo): el código DALI y su descripción solo se cargan desde la Consola; en la web solo se edita lo de SAP. Por eso la descripción local siempre se envía, aunque el artículo se haya tocado en la web. (La web se adapta en su v1.20.64.)
- Al terminar se vuelve a comparar y la consola dice si la web y el catálogo local coinciden, o cuántos quedan distintos y por qué (comparación de la fase 1).
- Un 403 explica que falta el permiso; un 401 borra las credenciales guardadas, como la descarga.
- Pruebas: `tests/test_catalogo_web.py` (12) y `tests/test_publicar_catalogo_web.py` (7), con una web falsa; 152 en total. Ventana vista en pantalla virtual.
- **Por probar en Windows con datos reales:** la primera sincronización (~30.000 artículos, una sola vez por equipo), una Acción 2 o 3 y una tarifa reales con un usuario con el permiso, y comprobar en la web (Administración) el registro de envíos.
- **Manual:** pendiente de regenerar con la próxima versión (ventana de resumen y textos nuevos de la consola).

---

## 53. Inicio de sesión de Catálogo Asignaciones al abrir la Consola (13 nov 2026) — PR en GitHub

Petición de Víctor: «incluir el login de Catálogo Asignaciones también en la Consola; solo usuario con login administrador y activo para escribir consola tendrá permiso. No hace falta anotar nada: si el login no se corresponde, mensaje de usuario no autorizado».
- La web ya devuelve lo necesario al iniciar sesión (`POST /auth/login`): solo deja entrar a cuentas activas y trae `rol`, `esAdminPrincipal` y `permisos`. No hace falta tocar la web.
- `catalogo_web.usuario_autorizado` / `comprobar_acceso` (sin tkinter): entra quien tiene rol `admin` y el permiso `admin-catalogo-escribir`, o es administrador principal (como en la web). Estados: correcto, no autorizado, email o contraseña incorrectos, verificación por email pendiente, sin conexión y otro error.
- `app/ui/ventana_login.py`: la ventana; la comprobación va en un hilo, para que no se congele. La ventana principal arranca oculta y solo se enseña si se entra; «Salir» o cerrar la ventana cierra la Consola. Las credenciales guardadas (cifradas) se prueban solas; si son incorrectas o la cuenta ya no está autorizada, se borran. Sin conexión no se borran y se puede reintentar. No se anota nada.
- Tras entrar, las credenciales de la sesión sirven para todo lo demás (descarga del catálogo y envío de cambios), aunque no se hayan podido cifrar.
- La comprobación de actualizaciones va **antes** del inicio de sesión, y su ventana sale aunque la principal esté oculta: si algún día el inicio de sesión fallara (web caída, un error), la versión que lo arregle llegaría igual.
- Pruebas: 2 nuevas en `tests/test_catalogo_web.py` (quién está autorizado y cada respuesta de la web) y `tests/test_login_consola.py` (5, con tkinter: credenciales guardadas, no autorizado, sin conexión y reintento, entrar y salir). 159 en total. La Consola completa probada en pantalla virtual con una web simulada (no autorizado: no se abre; autorizado: se abre).
- **Por probar en Windows:** entrar con una cuenta autorizada, con una de administrador sin el permiso y con una de hotel; y que la descarga del catálogo y el envío de cambios ya no vuelvan a pedir la contraseña.
- **Manual:** pendiente de regenerar con la próxima versión (pantalla de inicio de sesión).
- **Ojo al publicar:** hasta que cada comprador tenga el permiso «Escribir desde la Consola Registro Info», no podrá abrir la Consola.

---

## 54. Un solo archivo de artículos: adiós al LISTADO CODIGOS DALI - SAP (13 nov 2026) — PR en GitHub

Petición de Víctor: «¿Podríamos prescindir de uno de los Excel y mantener únicamente el catálogo completo que se guarda en RECURSOS? Creo que este listado realmente es el completo y reúne todo lo necesario». Confirmado: el catálogo tiene todas las columnas del LISTADO (A-C datos DALI, D código SAP, E-N datos SAP, O activo) y además proveedor, clasificación, código de proveedor y precio. Respuestas: nadie abre el LISTADO fuera de la Consola; primero esto y después la acción «Asignar SAP pendientes».

**Riesgo que se corrige:** la Acción 2 escribía código, descripción, unidad y activo en el LISTADO por **posición de fila** (fila 8+n, en el orden de ARTICULOS_DALI) y el código SAP de la columna D se quedaba en su fila. Si DALI quitaba un código de en medio o intercalaba uno, todos los SAP de debajo quedaban emparejados con el artículo de al lado. Ahora todo va por código DALI (prueba `test_sap_por_codigo_aunque_dali_quite_e_intercale_codigos`).

**Cómo (para cambiar lo menos posible la lógica que ya funcionaba):** `app/comun/catalogo_local.py` ofrece el catálogo con la misma forma que el LISTADO (filas desde la 8, columnas A..O).
- Solo leían (cambia la línea que abría el archivo): Acción 4 e INC (`_leer_tabla_cruce_listado_sap`, búsqueda del archivo para la ventana de faltantes), paso 27b de la transformación de la tarifa, y RINFO (`_localizar_ruta_listado_dali` → `_localizar_ruta_catalogo`, con el nombre antiguo como alias; 3 lecturas en `creador_rinfo` y 1 en `generador_existente`).
- Acción 3: el cruce de siempre sobre un libro en memoria sacado del catálogo; al final, `tabla_sap_desde_libro` y regenerar el catálogo con esos datos. La auditoría de SAP duplicados sigue igual (va antes).
- Acción 2: regenera el catálogo (DALI de ARTICULOS_DALI; SAP, código de proveedor y precio del catálogo anterior, por código DALI).
- Ventana de SAP faltantes: escribe el código SAP en el catálogo, con el bloqueo del catálogo (se suelta pase lo que pase) y, al guardar, envía los cambios a la web (`_tras_cambiar_sap_en_catalogo`). Se quita `_refrescar_sap_catalogo_para_dali`, que copiaba del LISTADO al catálogo. «Modificar directamente en el Excel» abre ahora el catálogo.
- Paso único: `_pasar_listado_al_catalogo()` al iniciar sesión y al regenerar. Si hay LISTADO, sus datos SAP pasan al catálogo (manda el LISTADO, como en cada Acción 2/3 hasta ahora), queda un informe `RECURSOS/paso_listado_a_catalogo_<fecha>.txt` con las diferencias y el LISTADO se renombra a «LISTADO CODIGOS DALI - SAP (antiguo, ya no se usa).xlsx» (no se borra). Lo hace el primer PC que abra esta versión.
- Nuevo riesgo cubierto: «descargar el catálogo de Catálogo Asignaciones» sustituye el archivo local. Sin LISTADO, un código SAP escrito en la Consola y aún no enviado («Ahora no», sin conexión) solo está en el catálogo local: al descargar, si la web no tiene SAP para ese artículo, se conserva el local (`conservar_sap_local`); si los dos tienen y no coinciden, manda la web.
- Se quita `_proteger_hoja_destino` y la contraseña de protección de hojas que estaba escrita en el código (4 sitios; el repositorio es público).
- Fallo encontrado y corregido en su propio commit: el paso 27b de la tarifa leía la «unidad SAP» de la columna F del LISTADO, que desde agosto es «Texto UM»: la descripción salía «… - Kilogramo» en vez de «… - KG» (como el de la sección 496, que se quedó sin corregir aquí).
- También en su propio commit: se recupera el encabezado «Próximos pasos» de este historial, borrado por error al añadir la sección 53.
- Pruebas: `tests/test_catalogo_local.py` (10) y `tests/test_catalogo_unico.py` (5, de punta a punta con las clases reales: regeneración, paso del LISTADO, primera vez sin catálogo, Acción 3 y la ventana de SAP faltantes). 174 en total.
- **Por probar en Windows con datos reales:** el paso del LISTADO al abrir (revisar el informe de diferencias), una Acción 2 y una 3 completas, una tarifa con SAP faltantes y un RINFO (nuevo y desde Purchasing existente).
- **Ojo al publicar:** las copias antiguas de la Consola siguen buscando el LISTADO. Cuando el primer PC lo haya pasado al catálogo, una Consola sin actualizar dirá «Listado no encontrado» en las Acciones 2, 3 y 4 (sin estropear nada): conviene que todos actualicen.
- **Manual:** pendiente de regenerar con la próxima versión (la sección 5 habla del LISTADO).

---

## 55. Acción 8: Asignar SAP pendientes (13 nov 2026) — PR en GitHub

Petición de Víctor: «una nueva acción, parecida a cuando se abre la ventana para introducir los códigos SAP faltantes en una tarifa, pero con todos los códigos DALI activos sin SAP, para que el usuario vaya grabando los SAP en el orden que quiera». Decisiones: desde la Consola (no desde la web: el código SAP y sus datos salen de los listados de SAP que solo tiene la Consola); botón nuevo (8); un código que no esté en `Productos.xlsx` se anota igualmente y se informa de que hasta que no se descarguen de nuevo `Productos.xlsx` y `Listas de productos.xlsx` en RECURSOS y se pulse otra vez la Acción 3 no quedará todo cumplimentado.
- `app/comun/asignar_sap.py` (sin tkinter): pendientes (activos sin SAP), filtros sin acentos ni mayúsculas, validación del código, consulta en `Productos.xlsx`, códigos repetidos (con la misma forma que usa `VentanaResolverSAPDuplicados`) y aplicar al catálogo.
- `app/ui/ventana_asignar_sap.py`: tabla con buscador y filtros por proveedor y familia (pueden ser cientos o miles: no una casilla por fila como en la ventana de faltantes). Se elige una fila, se escribe el código y con Intro se asigna y pasa al siguiente sin asignar; al lado, lo que dice `Productos.xlsx` (verde si está, naranja si no). «Guardar en el catálogo» en un hilo; las filas guardadas desaparecen y se sigue. Al cerrar con códigos sin guardar, pregunta. La pregunta de códigos repetidos y el final del guardado los atiende la ventana desde el hilo principal (sondeo), sin tocar Tk desde el hilo de fondo.
- `app/principal/accion_asignar_sap.py`: guarda bajo el bloqueo del catálogo y después completa los datos SAP con el mismo cruce de la Acción 3 (que regenera el catálogo y envía a Catálogo Asignaciones). Si falta `Productos.xlsx`, `Listas de productos.xlsx` o `ARTICULOS_DALI.xlsx`, solo se envían los códigos.
- La lectura de `Productos.xlsx` se movió sin cambios a `sap_helpers.leer_productos_xlsx` (commit propio) para compartirla.
- Pruebas: `tests/test_asignar_sap.py` (6) y `tests/test_accion_asignar_sap.py` (5: guardado con repetidos resueltos o cancelados, completar los datos con el cruce de la Acción 3 con archivos inventados, y la ventana en pantalla virtual). 185 en total. Ventana vista en pantalla virtual con datos inventados.
- **Por probar en Windows con datos reales:** abrir la acción con el catálogo real (cuántos pendientes salen y si la tabla va fluida), asignar algunos códigos (uno que esté y otro que no esté en `Productos.xlsx`) y comprobar el catálogo, el resumen de la Acción 3 y la ventana de cambios para Catálogo Asignaciones.
- **Manual:** pendiente de regenerar con la próxima versión (botón 8).

---

## 56. Manual de usuario al día (13 nov 2026) — PR en GitHub

Petición de Víctor: «revisar los manuales de Catálogo Asignaciones y Consola y actualizarlos».
- **Capturas** (`docs/manuales_fuente/capturas.py`): la consola se arranca sin el inicio de sesión real y su ventana se fotografía aparte; el proyecto de ejemplo ya no lleva LISTADO sino un catálogo inventado; capturas nuevas `c_login`, `c_publicar` (cambios para Catálogo Asignaciones) y `c_asignar_sap` (Acción 8); `c_sap_faltantes` con «Guardar en el catálogo»; la principal con el botón 8 y la versión del `.exe` que llevará el manual (v 10.2.52); fuera `c_credenciales` (ya lo hace el inicio de sesión).
- **Texto** (`manual_usuario.js`): inicio de sesión (capítulo 2), archivos del proyecto sin LISTADO (4), capítulo 5 «El catálogo DALI / SAP» (Acciones 1, 2, 3 y 8, descarga del catálogo y nota del paso del LISTADO), códigos SAP faltantes en el catálogo (6), capítulo 8 nuevo «Cambios en Catálogo Asignaciones» y tabla de problemas frecuentes al día (10). 27 páginas, 31 apartados en el índice de la consulta.
- Regenerados `docs/Manual_Usuario_ConsolaRegistroInfo.docx` / `.pdf` y copiados el PDF y `manual_indice.json` a `assets/` (van dentro del `.exe`). La consulta del manual encuentra los apartados nuevos («usuario no autorizado», «asignar sap pendientes», «enviar cambios a la web»).
- En commits propios: (1) los comentarios y la documentación decían v10.2.52 / v10.2.54 / v10.2.56 para lo de hoy, pero todo salió junto en la **v10.2.50** (borrador compilado del commit 86d157c); (2) textos visibles que aún nombraban el LISTADO (aviso de la tarifa «sin DESCRIPCION/UNIDAD SAP», ventana de duplicados, auditoría de la Acción 3 y mensajes de RINFO).
- **Para que llegue a los compañeros:** compilar otra vez (saldrá la v10.2.52, con el manual nuevo dentro) y publicar esa en lugar del borrador v10.2.50.

---

## 57. Acción 8: copiar los datos del artículo para crear el material en SAP (13 nov 2026) — PR en GitHub

Petición de Víctor (con una captura de la ventana real, 560 pendientes): «en esta ventana se deberá poder copiar la información para pegarla en la web de SAP al crear el nuevo material, ya lo hacemos en la ventana de códigos DALI sin SAP en la Acción 4».
- La tabla (Treeview) no deja seleccionar texto. Ahora el artículo elegido sale abajo en casillas de **solo lectura** (código DALI con sus ceros, descripción y unidad), seleccionables como en la ventana de faltantes de la Acción 4, cada una con **📋 Copiar**. Debajo, proveedor y familia; al copiar se confirma «✔ Copiado (…)».
- Clic derecho en la tabla: copiar código DALI, descripción, unidad, proveedor o la fila entera (separada por tabuladores). Ctrl+C con la tabla activa copia la descripción.
- «Código SAP», «Asignar» y «Quitar» siguen juntos en su propia fila.
- Prueba nueva en `tests/test_accion_asignar_sap.py` (portapapeles real en pantalla virtual). 186 en total.
- Manual: captura y leyenda de la Acción 8 al día (marca 6) y paso «copia su descripción y su unidad para pegarlas en SAP»; PDF e índice copiados a `assets/`.
- **Por probar en Windows:** copiar y pegar en la web de SAP.

---

## 58. Compilación sin el aviso de Node.js 20 y pendiente del «Precio €» cerrado (13 nov 2026) — PR en GitHub

Víctor dio el visto bueno a dos cosas que ofrecí al repasar los pendientes:
- **Workflow** (`.github/workflows/compilar.yml`): GitHub avisaba «Node.js 20 is deprecated». Se suben las acciones a la primera versión que ya usa Node 24, comprobado en su `action.yml`: `actions/checkout` v4 → v5, `actions/setup-python` v5 → v6 y `actions/upload-artifact` v4 → v6 (la v5 todavía usaba Node 20). Las opciones que usa el workflow (`python-version`, `name`, `path`, `retention-days`) no cambian. Lo mismo en el Organizador y en la app del móvil (con `actions/setup-java` v4 → v5).
- **Pendiente de la sección 37** («¿guardará la web el "Precio €"?»): resuelto con la etapa 6 (secciones 52 y 54): la web guarda el precio y el código de proveedor de cada tarifa (Acción 4) y una descarga del catálogo ya no pierde los códigos SAP locales.
- **Por probar:** la próxima compilación desde Actions (debería salir sin el aviso).

---

## 59. La actualización automática falla si el `.exe` está en una unidad de red (13 nov 2026) — PR en GitHub

Captura de Víctor al actualizar de la 10.2.50 a la 10.2.54: «Actualización no completada: No se pudo renombrar el programa actual: [WinError 32] El proceso no tiene acceso al archivo porque está siendo utilizado por otro proceso: 'G:\\…\\CODIGO\\TARIFA_Finder.exe' -> '…TARIFA_Finder.anterior.exe'».
- **Causa:** la actualización (sección 48) renombra el `.exe` abierto a `.anterior.exe`. En un disco local Windows lo deja; en una carpeta compartida de red (G:), no. Ya estaba así desde la sección 48: las pruebas en Windows (sección 50) se hicieron con el `.exe` en disco local.
- **Arreglo** (`app/comun/actualizaciones.py`): si ese renombrado falla, `sustituir_exe` lanza `ExeEnUso` (hija de `RuntimeError`) sin haber tocado nada, y `sustituir_al_cerrar` deja en `%TEMP%` un `actualizar_consola_<pid>.cmd` y lo lanza oculto. El script espera a que termine la Consola (hasta 90 s), mueve `.nuevo.exe` sobre el `.exe` (hasta 60 reintentos, uno por segundo: el proceso de arranque de PyInstaller lo suelta un instante después), abre el `.exe` (el nuevo o, si no se pudo mover, el de siempre) y se borra solo. Espera con `ping` y no con `timeout`, que sin consola falla al momento. La ventana (`app/ui/ventana_actualizacion.py`) lo explica y cierra la Consola, que se vuelve a abrir sola. En disco local todo sigue igual.
- **Pruebas** (`tests/test_actualizaciones.py`, 3 nuevas, 189 en total): renombrado con WinError 32 → `ExeEnUso` y el `.exe` intacto; texto del `.cmd` (rutas con espacios, tilde y `%`, espera al PID, sin `timeout`, sin bucles sin fin, se borra solo); y que se lanza oculto con el entorno de PyInstaller limpio.
- **Para la actualización que ya ha fallado:** quien actualiza es la versión instalada (10.2.50), que no lleva el arreglo, así que esta vez hay que terminar a mano (cerrar la Consola; en `CODIGO`, renombrar `TARIFA_Finder.exe` a `TARIFA_Finder.anterior.exe` y `TARIFA_Finder.nuevo.exe` a `TARIFA_Finder.exe`; abrirla). A partir de la versión que lleve este arreglo, las siguientes se instalan solas también en G:.
- **Por probar en Windows:** actualizar con el `.exe` en G: (de la versión con el arreglo a la siguiente).

---

## 60. Acción 8 con tabla editable, guía de los listados de SAP y sin preguntas de catálogo (13 nov 2026) — PR en GitHub

Peticiones de Víctor del 13 nov 2026:
- **«Me gusta más este sistema para la Acción 8, pero mantener los filtros»** (con una captura de «Códigos SAP faltantes» de la Acción 4). `app/ui/ventana_asignar_sap.py` se rehízo con esa tabla: una fila por artículo con el código DALI, la descripción, el formato, el proveedor y la familia en celdas de solo lectura seleccionables (Ctrl+C copia la celda; clic derecho → copiar cualquier dato o la fila entera, para pegarlo en SAP al crear el material) y la columna «CÓDIGO SAP (introducir)»; Intro (o flecha abajo) pasa a la fila siguiente. Al lado, «EN PRODUCTOS.XLSX»: verde si está, naranja si no está o si el código está repetido en otra fila, rojo si no es válido. Se mantienen el buscador y los filtros por proveedor y familia, y lo escrito se conserva al filtrar (una `StringVar` por DALI). Como mucho 250 filas a la vez (con miles, Tk se vuelve lento): para el resto, filtrar. El guardado (hilo, duplicados, cruce de la Acción 3) no cambia.
- **Avisar de que faltan los listados de SAP actualizados.** Tras guardar códigos SAP en la Acción 8 o en «Códigos SAP faltantes» de la Acción 4, si alguno no está en `Productos.xlsx`, `App._avisar_listados_sap` abre `app/ui/ventana_listados_sap.py`: explica que los códigos ya van a la web pero sus datos SAP no, y guía con imágenes la descarga de `Productos.xlsx` (aplicación de SAP «Gestionar datos maestros de producto») y `Listas de productos.xlsx` («Lista de productos»): «Ir», botón ⌄ de exportar (el último de la barra, «Abrir menú») → «Exportar como…» → nombre y formato Excel → «Exportar» → guardar en RECURSOS. La ventana enseña si cada archivo ya se ha descargado de nuevo (fecha de modificación). «Ya están en RECURSOS» lanza sola la Acción 3 (si alguno no se ha renovado, pregunta antes); «Más tarde» avisa de que la web queda parcialmente actualizada (y lo deja en la consola). Las imágenes son esquemas dibujados a partir de las capturas de Víctor (`compilacion/generar_guia_sap.py` → `assets/guia_sap_*.png`, en `ASSETS_NECESARIOS`); las capturas originales no llegaron como archivo a la sesión: si se prefieren, basta con dejarlas en `assets/` con el mismo nombre. Lógica sin tkinter en `app/comun/listados_sap.py`. El mensaje de la tarifa «DALI sin DESCRIPCION/UNIDAD SAP» ya no habla de la transacción ME3N, sino de esas dos aplicaciones, y la ventana de faltantes ya no lleva el «Recuerda: actualiza Productos…».
- **«Sigue saliendo "Catálogo DALI completo — puede estar desactualizado"; ya no debería porque la conexión es directa con la web»** (con dos registros). Se quitaron esa pregunta (Acción 4, `_catalogo_reciente_o_confirmar`, que ahora devuelve siempre True) y la equivalente de las Acciones 2 y 3 («comprobar actualización», `_confirmar_catalogo_actualizado`). Con la etapa 6 el envío a la web solo manda lo cambiado y deja protegido lo editado allí, así que no hace falta descargar el catálogo antes. En el registro se veía además que la descarga tardaba más de 60 s y fallaba («The read operation timed out»). Las funciones de descarga (`_descargar_catalogo_dali_web*`) quedan sin uso, por si se quiere un botón.
- Esta rama incluye también el arreglo de la actualización en carpetas de red (sección 59, PR propia), porque los dos tocan el historial, CLAUDE.md y el manual.
- **Manual:** capturas nuevas de la Acción 8 y de la guía, apartado «Descargar los listados de SAP», tabla de archivos con la aplicación de SAP de cada listado, nota de la carpeta de red en las actualizaciones y filas nuevas en «Consulta rápida». PDF e índice copiados a `assets/`.
- **Pruebas:** 196 (las de la ventana de la Acción 8 rehechas para la tabla, y `tests/test_listados_sap.py` nuevo: lógica, ventana y enganche con la Acción 3).
- **Por probar en Windows:** la Acción 8 con muchos pendientes (rapidez al filtrar) y el clic derecho; la guía con los listados reales (que «Ya están en RECURSOS» lance la Acción 3 y que esta complete los datos).

---

## 61. Acción 6: «Abrir en Excel» a la primera (13 nov 2026) — PR en GitHub

Captura de Víctor: «desde las últimas actualizaciones, debo pulsar 2 veces en abrir para que se abra». La consola registraba las dos pulsaciones («Acción 6 — Abriendo en Excel» dos veces), así que el clic llegaba y el archivo se mandaba abrir las dos veces.
- **Causa probable:** la ventana de la Acción 6 es modal (`grab_set`) y Windows no deja que un programa recién abierto le quite el primer plano. Excel se abría **detrás** de la Consola y la segunda pulsación solo lo traía delante (con el archivo ya abierto, Excel se activa). No se ha podido reproducir sin Windows.
- **Arreglo:** `abrir_con_su_programa(ruta)` (`app/comun/tarifa_core.py`) llama antes a `AllowSetForegroundWindow(ASFW_ANY)` para que Excel pueda ponerse delante; fuera de Windows, `xdg-open`. En la Acción 6, tras abrir el archivo, la ventana suelta el modo modal y se cierra (lo siguiente se hace en Excel), y la consola dice «Edítalo en Excel, guárdalo y ciérralo; después, Acción 4». «Abrir resultado» (Acción 5) usa la misma función.
- **Pruebas:** `tests/test_abrir_con_su_programa.py` (3): el permiso va antes de abrir, fuera de Windows no falla, y en la ventana de la Acción 6 un solo clic abre el archivo una vez y cierra la ventana. 199 en total.
- **Por probar en Windows:** que Excel salga delante a la primera, en la Acción 6 y en la Acción 5.

---

## 62. ⚠️ Listas de productos.xlsx con otras columnas vació la Categoría valoración (13 nov 2026) — PR en GitHub

Registro de Víctor del 13 nov 2026 (12:09–12:34, v10.2.56): tras asignar 4 códigos SAP en «Códigos SAP faltantes», descargó los listados de SAP y pulsó «Ya están en RECURSOS». La Acción 3 que se lanzó sola dijo «7466 código(s) con Categoría valoración distinta según el centro» (con nombres de centro como «GUAYARMINA PRINCESS (0110)»), «0 materiales con Categoría valoración desde Listas de productos.xlsx» y «7159 código(s) SAP sin coincidencia», y envió a Catálogo Asignaciones «Datos SAP: 7159 cambiado(s)».
- **Causa:** `_leer_categoria_valoracion` leía por posición: A = código SAP y F = Categoría valoración. La exportación nueva traía en la A «QUESO BRIE (10100050)» (descripción y código, como SAP exporta los campos con clave) y en la F el centro. Ningún código casó, así que la Acción 3 limpió la columna K del catálogo (como hace siempre antes de rellenar) y la dejó vacía, y el envío a la web mandó esos 7.159 cambios. Ya estaba así antes; salió ahora porque la exportación tenía otras columnas.
- **Arreglo** (`app/comun/sap_helpers.py`): `leer_listas_productos` busca las columnas por su cabecera («Producto» y «Categoría valoración», sin importar acentos) y saca el código de «DESCRIPCIÓN (código)» (`clave_entre_parentesis`). Sin esas cabeceras usa A y F como siempre, **sin** sacar nada de paréntesis, para no colar el código del centro como categoría. `comprobar_listas_productos` para la Acción 3 (y el cruce de la Acción 8) **antes de escribir** si el listado no casa con Productos.xlsx (menos de la mitad de sus códigos) o no trae ninguna categoría, con un mensaje que dice cómo volver a descargarlo. Los conflictos por centro se resumen en 20 líneas (el registro tenía 7.500).
- ⚠️ **Datos en producción:** el catálogo local y la web tienen ahora vacía la Categoría valoración de esos 7.159 materiales. **Se recupera** pasando la Acción 3 con la versión que lleva este arreglo y un Listas de productos.xlsx que tenga las columnas «Producto» y «Categoría valoración»: rellena la columna K y la Consola envía a la web solo esas diferencias. Es idempotente (repetirla no cambia nada más). Hasta entonces, no pasar la Acción 3.
- **Pruebas:** `tests/test_listas_productos.py` (6): formato de siempre, la exportación con «DESCRIPCIÓN (código)» y otras columnas, el archivo sin cabeceras conocidas (se para), vacío (se para), los paréntesis, y que la Acción 3 lanza el error y lo dice en la consola. 205 en total (incluye la sección 61, que va en la misma rama).
- **Por probar con datos reales:** el Listas de productos.xlsx que se descargó ese día (si sus cabeceras no son «Producto» y «Categoría valoración», la Acción 3 se parará y habrá que ver sus nombres).

---

## 63. Barra de progreso hasta el final, tarifa abierta en Excel y categoría de valoración recuperada (13 nov 2026) — PR en GitHub

Registro de Víctor de las 12:59–13:03 (v10.2.58) y sus archivos (catálogo, Productos.xlsx y Listas de productos.xlsx):
- **Recuperada la Categoría valoración** (sección 62): la Acción 3 con la v10.2.58 leyó «7159 materiales con Categoría valoración desde Listas de productos.xlsx» y envió a la web los 7.159 cambios. El Listas de productos.xlsx nuevo trae «Producto» en la columna A («QUESO BRIE (10100050)») y «Categoría valoración» en la **D**, con «Centro» en la F: confirmado que la lectura por cabecera era necesaria.
- **«No hay sincronización entre el log y la barra de tareas, termina en verde siempre sin haber terminado».** La barra se completaba al acabar la parte local, pero el envío a Catálogo Asignaciones seguía en segundo plano (en el registro, de 13:01:09 a 13:02:06). `CanvasProgressBar.ocupar()/liberar()` (`tarifa_core.py`): `_publicar_catalogo_web` la ocupa con «Enviando cambios a Catálogo Asignaciones…» y la libera al acabar el último envío; mientras tanto, `completado()` y `reset()` quedan pendientes y se aplican al terminar. Pruebas en `tests/test_barra_progreso.py` (4).
- **El «[INCOMPLETO]» reaparecía.** A las 12:43 y a las 13:00 la Acción 4 renombró la tarifa de PRO A PRO (sin [INCOMPLETO]) y poco después volvía a encontrarla con [INCOMPLETO]. Lo más probable: seguía abierta en Excel desde la Acción 6, y al guardar en Excel se creaba otra vez con el nombre antiguo. Antes de renombrar se comprueba el archivo «~$…» que deja Excel (`sap_helpers.abierto_en_excel` / `comprobar_cerrado_en_excel`, en los tres sitios de `accion_tarifa.py`); si está abierta, no se renombra y se pide guardarla y cerrarla. El mensaje final ya no dice «TARIFA PREPARADA» si no se pudo quitar [INCOMPLETO]. Pruebas en `tests/test_abierto_en_excel.py`. **Conviene** revisar en TARIFAS si quedan las dos copias (con y sin [INCOMPLETO]) y quedarse con la buena.
- 210 pruebas.
- **Por probar en Windows:** que la barra siga en movimiento hasta el final del envío, y que con la tarifa abierta en Excel la Acción 4 avise y no la renombre.

---

## 64. Calendario en «Fechas de validez del Registro Info» (13 nov 2026) — PR en GitHub

Petición de Víctor (con captura): «establecer calendario para no tener que escribir las fechas manualmente con su desplegable; la norma de pegado de las fechas en formato no debe cambiar para el RINFO».
- Botón 📅 junto a «Inicio» y «Fin» en `VentanaFechasRINFO` (`app/rinfo/ventanas_rinfo.py`). Abre `app/comun/selector_fecha.py`: calendario propio (sin dependencias nuevas, solo tkinter y `calendar`), en el mes de la fecha que haya en la casilla, con desplegables de mes y año, ◀ ▶, «Hoy» y semanas de lunes a domingo.
- Al pulsar un día se escribe en la casilla como **dd.mm.aaaa**, el formato de siempre: la validación (`_parsear_fecha`) y lo que se pega en la hoja 3 del RINFO no cambian. La casilla se sigue pudiendo escribir a mano.
- **Pruebas:** `tests/test_selector_fecha.py` (3): formato y lectura, semanas que empiezan en lunes, y en la ventana real, elegir en el calendario y confirmar devuelve `("01.10.2026", "31.03.2027")`. 213 en total.
- **Por probar en Windows:** el aspecto del calendario y el icono 📅.

---

## 65. Fechas de validez del Registro Info según el periodo de la tarifa (13 nov 2026) — PR en GitHub

Registro de Víctor de las 13:53–13:58 (v10.2.60): la tarifa de PRO A PRO dice «4º TRIMESTRE 2026», pero el Registro Info se creó con `inicio=01.01.2026 fin=31.12.2026`, que es justo lo que la ventana proponía por defecto (año en curso). Resultado en el RINFO: «absorción TOTAL» de periodos anteriores, 430 periodos cerrados y 530 vigencias recortadas (las que llegaban al 31.03.2027). Propuesta aceptada por Víctor («sí, perfecto»):
- **Fechas propuestas** (`app/comun/fechas_tarifa.py`, sin tkinter): del periodo de la tarifa (celda C3, «Periodo:», la escribe la Acción 4): trimestres («4º TRIMESTRE 2026», «T4», «4T», «IV TRIMESTRE»), semestres, meses o rangos de meses («OCT-DIC 2026», «Octubre 2026 - Marzo 2027», también cruzando de año), dos fechas («01/10/2026 - 31/03/2027») o el año suelto. Si no se entiende, las últimas fechas usadas en este PC (`%LOCALAPPDATA%\ConsolaRegistroInfo\ultimas_fechas_rinfo.json`); si no hay, el año en curso, como antes. La ventana dice de dónde salen.
- **Aviso al confirmar:** si las fechas no coinciden con el periodo de la tarifa, o si el inicio es anterior a hoy y la tarifa no indica periodo, se pide confirmación con el motivo; «No» deja corregirlas. Si coinciden con la tarifa no pregunta nada (aunque el inicio sea anterior a hoy: es la retroactividad de la propia tarifa).
- El formato que se pega en el RINFO (dd.mm.aaaa) no cambia. `VentanaFechasRINFO(parent, on_confirm, periodo="")`: los dos sitios que la abren (`crear_registro_info.py`) le pasan el periodo de la tarifa elegida.
- **Pruebas:** `tests/test_fechas_tarifa.py` (6): periodos de ejemplo, propuesta, avisos del caso del registro, últimas fechas, lectura de C3 y la ventana real (propone 01.10.2026 – 31.12.2026, pregunta si se cambia a 01.01.2026 y no pregunta con las del periodo). 219 en total.
- ⚠️ **El RINFO de PRO A PRO del 13 nov (13:58)** se generó con 01.01.2026 – 31.12.2026: no importarlo en SAP si esas no eran las fechas buenas; volver a crearlo.

---

## Próximos pasos, en orden de prioridad
- **PENDIENTE (petición de Víctor, 13 nov 2026): auditar la sincronización con la web, unir las Acciones 1 y 2 y guiar la Acción 3.**
  - Petición: «auditar que todos los pasos se están sincronizando directamente con la web. Me gustaría que al terminar la acción 1, después de limpiar los Excel 1 y 2 y crear el limpio que se guarda en recursos, esta información ya se vuelque directamente en la web, ya que todas las columnas se integran luego en la acción dos al Excel "catálogo completo" y a la web; de esta manera podríamos eliminar creo el paso 2. Antes de subirlo a la web, se le pregunta al usuario si desea revisar el Excel limpio antes del volcado final en la web, dos botones sí - no; si dice que sí, se abre el Excel para revisión del usuario; una vez que el usuario cierre y guarde los cambios en Excel, en la misma ventana anterior se debe abrir un nuevo botón que indique subir datos a la web o actualizar catálogo asignaciones; este botón deberá estar oculto hasta que el usuario no indique sí o no a la primera pregunta».
  - **Auditoría:** repasar cada acción (1, 2, 3, 4, 6, 8, SAP faltantes y SAP duplicados) y anotar si escribe en `catalogo-dali-completo.xlsx`, si después llama a `_publicar_catalogo_web()` y qué campos llegan a la web. Hoy publican las Acciones 2, 3 y 4, la 8 (por el cruce de la 3) y la resolución de duplicados; la Acción 1 solo crea `RECURSOS/ARTICULOS_DALI.xlsx` (`app/principal/accion_organizar_dali.py`) y no publica nada.
  - **Nuevo final de la Acción 1:** tras crear `ARTICULOS_DALI.xlsx`, una ventana (no `messagebox`) con el resumen actual y la pregunta «¿Desea revisar el Excel limpio antes de actualizar Catálogo Asignaciones?» con **Sí / No**. El botón **«Actualizar Catálogo Asignaciones»** está oculto hasta que se responde. Con «Sí» se abre el Excel (`os.startfile`) y el botón aparece en la misma ventana, para pulsarlo cuando se haya guardado y cerrado el Excel (avisar si sigue abierto/bloqueado o no se ha guardado: comprobar fecha de modificación y que se puede abrir). Con «No» aparece directamente. Al pulsarlo: releer `ARTICULOS_DALI.xlsx` (con los cambios del usuario), hacer lo que hoy hace la Acción 2 (regenerar el catálogo por código DALI conservando SAP, código de proveedor y precio) y `_publicar_catalogo_web()` con su ventana de resumen (simulación primero), como ahora.
  - **Acción 2: se elimina** (decisión de Víctor, 13 nov 2026: «la idea es reducir pasos y que todo sea fluido y actualizado en la web lo antes posible»). Su código pasa a la Acción 1, no se pierde. No hace falta un «volver a enviar»: si no hay `.xls` nuevos, la Acción 1 ya ofrece usar el `ARTICULOS_DALI.xlsx` existente. Decidir al hacerlo si se renumeran los botones (1, 2 = antigua 3…) o se deja el hueco; revisar referencias en el manual, la consulta del manual, mensajes que dicen «Acción 2/3» y `tests/test_catalogo_unico.py`.
  - **Acción 3 (petición de Víctor, 13 nov 2026):** «debería salir el mensaje que ya hemos instaurado con las instrucciones de descarga de estos archivos para indicar al usuario que actualice ambos antes de continuar, también con opción sí o no; en caso de sí, abrir la carpeta RECURSOS para pegarlos, y en caso de no, que aparezca el botón igual que en el paso anterior para actualizar Catálogo Asignaciones en la web».
    - Al pulsar la Acción 3, antes de nada, la guía de `app/ui/ventana_listados_sap.py` (la de «Faltan los listados de SAP actualizados», con un modo «antes de la Acción 3»: otro título y sin la lista de códigos) con la pregunta «¿Necesita descargar de nuevo `Productos.xlsx` y `Listas de productos.xlsx`?» **Sí / No**.
    - **Sí:** abre RECURSOS y aparece el botón «Actualizar Catálogo Asignaciones»; mientras, la línea de estado que ya existe (cada 2 s) dice si cada archivo está descargado de nuevo. **No:** aparece el botón directamente. El botón está oculto hasta responder, igual que en la Acción 1.
    - Al pulsar el botón: la Acción 3 de siempre (con la protección de «Listas de productos» que no casa, que se mantiene) y el envío a la web.
    - **Que la Consola sepa sola si hace falta descargar los listados** (idea de Víctor, 13 nov 2026: «la aplicación podría saber si necesitas actualizar los dos Excel del paso 3? Normalmente se necesita porque se han insertado los códigos SAP pero no tiene el resto de información: descripción, unidad de medida, grupo, etc.»). Antes de preguntar Sí/No:
      1. Buscar en el catálogo los artículos **con código SAP y sin sus datos SAP** (descripción SAP, unidad, grupos, categoría de valoración vacíos).
      2. Mirar si esos códigos están en `Productos.xlsx` y en `Listas de productos.xlsx` de RECURSOS (ya existe `listados_sap.codigos_sin_datos`; falta lo mismo para Listas/categoría).
      3. Resultado: **nada que completar** → sin pregunta, directo al botón (la Acción 3 sirve igual para refrescar cambios hechos en SAP); **lo que falta ya está en los Excel de RECURSOS** → «Los listados de RECURSOS ya tienen los datos de N artículos que faltan» y directo al botón; **faltan códigos en los listados** → la guía de descarga con la lista de esos códigos y el estado que se actualiza solo (✔ cuando el archivo nuevo ya los contiene, no solo cuando cambia su fecha).
      4. Límite: no puede saber si en SAP cambió algo de un material que ya tiene datos (por ejemplo una descripción corregida). Para eso, mostrar la fecha de los listados y avisar si son antiguos (umbral a decidir con Víctor, p. ej. 30 días).
      5. La misma comprobación se puede usar tras las Acciones 1, 4 y 8 para avisar solo cuando de verdad falta algo.
  - **El Excel `catalogo-dali-completo.xlsx` al día en todos los pasos, cambiando solo lo necesario (como en la web):** hoy `_regenerar_catalogo_dali_completo` lo vuelve a escribir entero. Auditar si cada acción lo deja al día y valorar escribir solo las celdas/filas que cambian (comparando con lo que ya tiene, como hace `catalogo_web.py` con la web), sin tocar formato ni lo demás, y decir en la consola cuántas filas cambiaron. Ojo: con el Excel abierto por otra persona en la red no se puede guardar; mantener el aviso actual.
  - Lógica sin tkinter en `app/comun/` con pruebas; reutilizar el código de la Acción 2 tal cual (moverlo, no reescribirlo). Regenerar el manual (capítulo 5) al terminar. Probar en Windows con datos reales: abrir/editar/guardar en Excel y el envío a la web.
- ~~Acción «Asignar SAP pendientes»~~ hecha (sección 55). ~~Regenerar el manual~~ hecho (sección 56).
0. **Reorganización en GitHub** (sección 42), siguientes etapas, cada una en su PR:
   - ~~etapa 2: carpetas por capas~~ hecha (sección 43);
   - ~~etapa 3: limpieza~~ hecha (sección 44); ~~trocear `crear_registro_info.py`~~ hecho (sección 45); ~~trocear `tarifa_app.py`~~ hecho (sección 46);
   - ~~compilación automática en GitHub~~ hecha y probada (sección 47); ~~actualizaciones automáticas~~ hechas (sección 48), probadas en Windows (10.2.40 → 10.2.42, sección 50);
   - ~~manual de usuario con capturas~~ hecho (sección 49), con botón y consulta en la app; pendiente: aviso «Node.js 20 is deprecated» del workflow; si se quiere, seguir troceando `generador_existente.py` (`GeneradorDesdeExistente`, ~4.200 líneas) y `accion_tarifa.py` (~1.600), con el mismo método (sin cambiar código y comparando por AST).
1. Roadmap de limpieza — puntos **1, 2, 3, 4 y 5 CERRADOS**. Punto 6 con su **primera pasada hecha** (sección 32): boilerplate 100% idéntico consolidado; quedan 6 geometrías con fórmulas propias, sin tocar deliberadamente — no hay más trabajo pendiente salvo que en algún momento vuelvan a coincidir entre sí.
2. Recomendación pendiente para el usuario: revisar en SAP los ~50 materiales que aparecieron con el mismo hueco de Hoja3 que 10101839, por si llevan tiempo sin condición de precio activa sin que nadie lo notara.
3. Recomendación pendiente (sección 22): siguen existiendo ~1.594 materiales SAP más (además del 10102980, ya resuelto) sin cruce DALI en el LISTADO — el fallback a `Productos.xlsx` los cubre para la sincronización de UM, pero conviene revisar y completar el cruce DALI real cuando haya ocasión, ya que el fallback no puede aportar IGIC ni grupo de compras.

**URGENTE (sección 35):** revisar los RINFO ya generados con solapamiento parcial y sin filas nuevas en Hoja 3 (FRIT RAVICH SL 25-09, EMICELA 13-08; comprobar MILO, AHUMADOS y EMICELA 10-08): comparar precio del RINFO vs. tarifa y confirmar si se cargaron en SAP; regenerar y recargar los afectados.

**Pendiente (sección 36):** regenerar MAHOU con el código nuevo (140 filas nuevas en Hoja 3, no 130; el RINFO de la pasada anterior sigue incompleto en 10800192 y 10800193) y revisar los RINFO ya cargados con posibles huecos por centro (AHUMADOS TF-GC y de 27/28-07, EMICELA 13-08): hace falta la tarifa de cada uno para separar huecos reales de legítimos.

~~**Pendiente (sección 37):**~~ **Resuelto (sección 58):** con la etapa 6 la web guarda el precio y el código de proveedor (Acción 4) y una descarga ya no pierde los códigos SAP locales. Texto original: confirmar si la web Catálogo Asignaciones va a guardar también "Precio €" (y "Texto UM") antes de que una descarga sobrescriba precios que solo existan en local; decidir si aplicar formato de moneda a la columna "Precio €".

**Pendiente (sección 39):** ninguno de código — la solución es 100% manual en SAP (transacción "Marcar reg.info de compra para borrar"). Si vuelve a aparecer el aviso de `VentanaMaterialesParaReactivar`, seguir esos mismos pasos, centro por centro.

**Confirmado (sección 40):** la ventana ya se vio en real por primera vez el mismo día — funcionó correctamente y se mejoró tamaño y detalle de centros a partir de ese uso real.

**Pendiente (sección 41):** (1) decidir si se implementa la solución de fondo para filas con cambio de precio/fecha + cambio de UM (fila nueva aparte, hipótesis sin verificar en SAP); (2) valorar las dos observaciones sobre las bajas (productos de temporada marcados para borrado; decisión de vigencia por la primera fila de Hoja 3); (3) no volver a subir el mismo RINFO a SAP varias veces — si una carga se repite, regenerar desde un Purchasing recién descargado.

Sin bugs abiertos de código ni troceos pendientes a fecha de la sección 41 salvo lo indicado arriba; queda la revisión de impacto en SAP de las secciones 35/36 y la decisión de la sección 37.

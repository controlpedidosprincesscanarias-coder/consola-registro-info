# Fuente del manual de usuario

> La idea y el desarrollo de la aplicación, incluidas sus pruebas, son de **Víctor Alejandro Martín Andrade**, para el Departamento de Central de Compras de Princess Hotels en Canarias. Ver `AUTORIA.md`.

Genera `Manual_Usuario_ConsolaRegistroInfo.docx` / `.pdf` y `manual_indice.json` (lo que usa «Consultar el manual» dentro de la app), con el mismo estilo que los manuales de Control Pedidos y del Organizador.

## Archivos
| Archivo | Para qué |
|---|---|
| `capturas.sh` / `capturas.py` | Abre la consola real en una pantalla virtual (Xvfb), con un proyecto de **datos inventados** y sin red, y guarda una captura marcada (círculos naranjas numerados) de cada ventana en `cap/` |
| `manual_usuario.js` | El texto del manual |
| `lib.js` | Maquetación con `docx` de Node (portada, índice, capturas, leyendas, recuadros, tablas) y el índice de apartados para la consulta |
| `paginar_indice.py` | Pone a cada apartado del índice su página en el PDF (con `pdftotext`) |
| `generar.sh` | Todo seguido: docx → PDF (LibreOffice) → páginas del índice, en `salida/` |
| `badges/` | Los números de las leyendas |

`cap/` y `salida/` no se suben (`.gitignore`).

## Regenerarlo (Linux)
Requisitos: Python con tkinter y los paquetes de `compilacion/requisitos.txt`, Xvfb, ImageMagick, la fuente Inter, Node con `docx`, LibreOffice y `poppler-utils`.

```bash
PYTHON=/ruta/a/python docs/manuales_fuente/capturas.sh
```
```bash
cd docs/manuales_fuente && ./generar.sh
```
Después, copiar:
- `salida/Manual_Usuario_ConsolaRegistroInfo.docx` y `.pdf` a `docs/`;
- `salida/Manual_Usuario_ConsolaRegistroInfo.pdf` y `salida/manual_indice.json` a `assets/` (van dentro del `.exe`, ver `ASSETS_NECESARIOS` en `build_exe.py`).

`SOLO=c_principal` rehace solo esa captura; `VOLCAR=1` guarda el árbol de widgets de cada ventana (sirve para situar las marcas).

## Cómo se arranca la consola para las capturas
- Sin red: no hay Catálogo Asignaciones ni GitHub. El **inicio de sesión** se salta (`_iniciar_sesion_consola`) y su ventana se fotografía aparte (`c_login`).
- El proyecto de ejemplo lleva un `catalogo-dali-completo.xlsx` inventado (desde la v10.2.50 es el único archivo de artículos; no hay LISTADO), que usa la captura de la Acción 8.
- La versión que se ve en las capturas es la del `.exe` que llevará el manual (variable `VERSION`, por defecto la del manual), no la de `APP_VERSION` del código.

## Reglas
- **Solo datos inventados** en las capturas (proveedores DEMO / EJEMPLO / MODELO, ruta `G:\CENTRAL COMPRAS\PROYECTO`). Nunca tarifas, Purchasing ni listados reales.
- Si cambia una pantalla que sale en el manual, rehacer su captura y regenerar.
- Si cambia el texto, regenerar también `manual_indice.json`: si no, la consulta mostraría páginas o fragmentos antiguos.

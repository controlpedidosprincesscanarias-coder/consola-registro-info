# TARIFA Finder v2
### Creaci\xf3n Registro Info SAP \u2014 Central Compras Princess Canarias
**\xa9 VAMA 2026 \u2014 V\xedctor Alejandro Mart\xedn Andrade**

---

## Distribuci\xf3n

Solo se distribuye el archivo:
```
TARIFA_Finder.exe
```
**No requiere Python, ni librer\xedas, ni instalaci\xf3n.**

---

## Interfaz

| Zona | Contenido |
|---|---|
| **Cabecera** | Logo Princess, t\xedtulo del m\xf3dulo, sub\xedtulo corporativo |
| **Panel izquierdo** | Selecci\xf3n de carpeta, estado del archivo, barra de progreso, botones de acci\xf3n |
| **Panel derecho** | Ventana Inmediata (consola de log tipo VBA): muestra el flujo paso a paso con timestamp y colores por nivel |
| **Pie** | Firma \xa9 VAMA 2026 y denominaci\xf3n corporativa |

### Niveles de log en la consola
- `SISTEMA` \u2014 dorado \u2014 eventos del sistema
- `INFO`    \u2014 azul   \u2014 informaci\xf3n general
- `OK`      \u2014 verde  \u2014 operaciones completadas
- `WARN`    \u2014 amarillo \u2014 avisos
- `ERROR`   \u2014 rojo   \u2014 errores

---

## Compilar el ejecutable

1. Instalar dependencias (una sola vez):
   ```
   pip install pyinstaller pillow
   ```
2. Compilar:
   ```
   python build_exe.py
   ```
3. El ejecutable se genera en `dist/TARIFA_Finder.exe`.

> **Nota sobre Pillow:** se usa solo para renderizar el logo con fondo transparente.
> Si no est\xe1 instalado, la app funciona igualmente mostrando el texto "PRINCESS HOTELS & RESORTS".

---

## Archivos del proyecto

```
tarifa_finder_v2/
\u251c\u2500\u2500 tarifa_app.py       \u2190 C\xf3digo fuente principal
\u251c\u2500\u2500 build_exe.py        \u2190 Script de compilaci\xf3n
\u251c\u2500\u2500 README.md           \u2190 Este archivo
\u2514\u2500\u2500 tarifa_config.json  \u2190 Generado autom\xe1ticamente (carpeta guardada)
```

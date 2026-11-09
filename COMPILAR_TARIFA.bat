@echo off
setlocal EnableDelayedExpansion

:: Forzar UTF-8 en la consola
chcp 65001 >nul

:: Activar colores ANSI (Windows 10+)
>nul reg add "HKCU\Console" /v VirtualTerminalLevel /t REG_DWORD /d 1 /f

:: Colores ANSI
set "C_OK=[32m"       :: Verde
set "C_WARN=[33m"     :: Amarillo
set "C_ERR=[31m"      :: Rojo
set "C_RESET=[0m"     :: Reset

title TARIFA Finder -- Compilador

echo.
echo  ============================================================
echo   TARIFA FINDER -- COMPILADOR AUTOMATICO
echo  ============================================================
echo.

cd /d "%~dp0"
echo  [DIR] %CD%
echo.

:: ──────────────────────────────────────────────────────────────
:: PASO 1: Cerrar instancias previas
:: ──────────────────────────────────────────────────────────────
echo  [1/5] Comprobando instancias en ejecucion...
tasklist /FI "IMAGENAME eq TARIFA_Finder.exe" 2>nul | find /I "TARIFA_Finder.exe" >nul
if %ERRORLEVEL%==0 (
    echo        Instancia detectada. Cerrando...
    taskkill /F /IM "TARIFA_Finder.exe" >nul 2>&1
    timeout /t 2 /nobreak >nul
    echo        Aplicacion cerrada.
) else (
    echo        No hay instancias activas.
)
echo.

:: ──────────────────────────────────────────────────────────────
:: PASO 2: Leer versión actual
:: ──────────────────────────────────────────────────────────────
echo  [2/5] Version actual en tarifa_app.py...
python _bat_helper.py leer > _version_tmp.txt 2>&1
set /p VERSION_ACTUAL=<_version_tmp.txt
del _version_tmp.txt 2>nul
echo        Version actual: !VERSION_ACTUAL!
echo.

echo        (Solo escribe el numero, ejemplo: 8.7.4)
echo        (La "v" se añade automaticamente)
echo.
set "NUEVA_VERSION="
set /p "NUEVA_VERSION=  Nueva version (ENTER para mantener !VERSION_ACTUAL!): v "

if "!NUEVA_VERSION!"=="" (
    echo.
    echo        Sin cambios: !VERSION_ACTUAL!
) else (
    echo.
    python _bat_helper.py escribir "!NUEVA_VERSION!"
    if !ERRORLEVEL! neq 0 (
        echo  %C_ERR%[ERROR] No se pudo modificar tarifa_app.py%C_RESET%
        pause
        exit /b 1
    )
    python _bat_helper.py leer > _version_tmp.txt 2>&1
    set /p VERSION_NUEVA_FINAL=<_version_tmp.txt
    del _version_tmp.txt 2>nul
    echo        Actualizado: !VERSION_ACTUAL! -^> !VERSION_NUEVA_FINAL!
)
echo.

:: ──────────────────────────────────────────────────────────────
:: PASO 3: Limpiar compilaciones anteriores
:: ──────────────────────────────────────────────────────────────
echo  [3/5] Limpiando compilaciones anteriores...
if exist "build\" (
    rd /s /q "build" 2>nul
    echo        build\ eliminada.
)
if exist "dist\" (
    rd /s /q "dist" 2>nul
    echo        dist\ eliminada.
)
if exist "TARIFA_Finder.spec" (
    del /f /q "TARIFA_Finder.spec" 2>nul
    echo        .spec eliminado.
)
echo.

:: ──────────────────────────────────────────────────────────────
:: PASO 4: Verificar entorno Python
:: ──────────────────────────────────────────────────────────────
echo  [4/5] Verificando entorno Python...
python --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo  %C_ERR%[ERROR] Python no encontrado en PATH.%C_RESET%
    pause
    exit /b 1
)
for /f "tokens=*" %%V in ('python --version 2^>^&1') do echo        %%V

python -m PyInstaller --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo  %C_ERR%[ERROR] PyInstaller no instalado. Ejecuta: pip install pyinstaller%C_RESET%
    pause
    exit /b 1
)
echo        PyInstaller OK
echo.

:: ──────────────────────────────────────────────────────────────
:: PASO 4B: Verificar e instalar dependencias
:: ──────────────────────────────────────────────────────────────
echo  [4B] Verificando dependencias de Python...

set "REQ=pandas openpyxl xlrd pillow pywin32 numpy"

for %%M in (%REQ%) do (
    python -c "import %%M" 2>nul
    if !ERRORLEVEL! neq 0 (
        echo        %C_WARN%Falta %%M — instalando...%C_RESET%
        pip install %%M --quiet
        if !ERRORLEVEL! neq 0 (
            echo  %C_ERR%[ERROR] No se pudo instalar %%M%C_RESET%
            pause
            exit /b 1
        )
        echo        %C_OK%%%M instalado correctamente.%C_RESET%
    ) else (
        echo        %C_OK%%%M OK%C_RESET%
    )
)

echo.

:: ──────────────────────────────────────────────────────────────
:: PASO 5: Compilar
:: ──────────────────────────────────────────────────────────────
echo  [5/5] Compilando...
echo  ============================================================
echo.
python build_exe.py
set "BUILD_CODE=%ERRORLEVEL%"
echo.
echo  ============================================================

if %BUILD_CODE% neq 0 (
    echo  %C_ERR%[ERROR] Compilacion fallida ^(codigo %BUILD_CODE%^).%C_RESET%
    pause
    exit /b %BUILD_CODE%
)

if exist "dist\TARIFA_Finder.exe" (
    echo  %C_OK%[OK] Ejecutable generado en: %CD%\dist\TARIFA_Finder.exe%C_RESET%
)

echo.
if "!NUEVA_VERSION!"=="" (
    echo   Version compilada: !VERSION_ACTUAL!
) else (
    echo   Version compilada: !VERSION_NUEVA_FINAL!
)
echo.
echo  ============================================================
echo   %C_OK%COMPILACION COMPLETADA CON EXITO%C_RESET%
echo  ============================================================
echo.
pause
endlocal

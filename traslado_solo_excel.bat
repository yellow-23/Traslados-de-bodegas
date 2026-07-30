@echo off
chcp 65001 >nul
cd /d "%~dp0"

REM Usa la hoja del Excel tal cual esta, sin bajar nada de oygbase.

if not exist "procesar_traslado.py" (
    echo [ERROR] Falta procesar_traslado.py en esta carpeta.
    echo         Copia la carpeta completa, no solo este .bat.
    pause
    exit /b 1
)
if not exist "traslado.py" (
    echo [ERROR] Falta traslado.py en esta carpeta.
    echo         Copia la carpeta completa, no solo este .bat.
    pause
    exit /b 1
)

set PYEXE=

REM 1) Entorno virtual local, si alguien lo creo.
if exist ".venv\Scripts\python.exe" set PYEXE=.venv\Scripts\python.exe

REM 2) El lanzador py: es el mas confiable.
if not defined PYEXE (
    py -c "import sys" >nul 2>nul
    if not errorlevel 1 set PYEXE=py
)

REM 3) python del PATH, pero solo si de verdad ejecuta: en Windows 11 puede ser
REM    el acceso directo a la Microsoft Store, que existe y no corre nada.
if not defined PYEXE (
    python -c "import sys" >nul 2>nul
    if not errorlevel 1 set PYEXE=python
)

if not defined PYEXE (
    echo No se encontro Python en este equipo.
    echo Instalalo desde https://www.python.org/downloads/ marcando "Add Python to PATH".
    pause
    exit /b 1
)

REM Equipo nuevo: se instalan las dependencias una sola vez.
%PYEXE% -c "import pandas, openpyxl" >nul 2>nul
if errorlevel 1 (
    echo Faltan dependencias, instalando por unica vez...
    %PYEXE% -m pip install -r requirements.txt
    if errorlevel 1 (
        echo [ERROR] Fallo la instalacion de dependencias.
        pause
        exit /b 1
    )
    echo.
)

%PYEXE% procesar_traslado.py
set CODIGO=%errorlevel%

echo.
if %CODIGO%==0 (
    echo Proceso terminado.
) else (
    echo Proceso terminado con errores. Revisa el detalle de arriba.
)

pause
exit /b %CODIGO%

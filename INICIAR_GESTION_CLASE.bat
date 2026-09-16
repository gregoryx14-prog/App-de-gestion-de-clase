@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"
chcp 65001 >nul
title Gestión de Clase V5.1 - Inicio

echo.
echo ================================================================
echo       SISTEMA DE GESTION DE CLASE V5.1 - DOBLE CLIC
echo ================================================================
echo.

REM 1) Comprobar espacio libre en la unidad del sistema (minimo 4 GB).
for /f "tokens=3" %%A in ('dir /-c "%SystemDrive%\" ^| findstr /R /C:"[0-9][0-9]* bytes free"') do set FREE=%%A
if defined FREE (
  set FREE=!FREE:,=!
  set FREE=!FREE:.=!
)
if defined FREE if !FREE! LSS 4294967296 (
  echo [ERROR] Hay menos de 4 GB libres en %SystemDrive%.
  echo Libera espacio y vuelve a ejecutar este archivo.
  echo.
  pause
  exit /b 1
)

REM 2) Usar Python 3.12. NO instala Python 3.14.
set PYEXE=
py -3.12 -c "import sys; print(sys.executable)" > "%TEMP%\gestion_py_path.txt" 2>nul
if exist "%TEMP%\gestion_py_path.txt" set /p PYEXE=<"%TEMP%\gestion_py_path.txt"
del "%TEMP%\gestion_py_path.txt" >nul 2>&1

if not defined PYEXE (
  where python >nul 2>&1
  if not errorlevel 1 (
    for /f "delims=" %%P in ('python -c "import sys; print(sys.executable if sys.version_info[:2] == (3,12) else '')" 2^>nul') do set PYEXE=%%P
  )
)

if not defined PYEXE (
  echo [INFO] Python 3.12 no esta instalado.
  where winget >nul 2>&1
  if errorlevel 1 (
    echo [ERROR] No se encontro Python 3.12 ni winget.
    echo Instala Python 3.12 manualmente y vuelve a ejecutar este archivo.
    pause
    exit /b 1
  )
  echo [INFO] Instalando Python 3.12 mediante winget...
  winget install --id Python.Python.3.12 --scope user --silent --accept-package-agreements --accept-source-agreements
  if errorlevel 1 (
    echo [ERROR] No fue posible instalar Python 3.12.
    pause
    exit /b 1
  )
  for /f "delims=" %%P in ('py -3.12 -c "import sys; print(sys.executable)" 2^>nul') do set PYEXE=%%P
)

if not defined PYEXE (
  echo [ERROR] Python 3.12 no esta disponible despues de la instalacion.
  echo Reinicia Windows una vez y vuelve a ejecutar este archivo.
  pause
  exit /b 1
)

set VENV=%~dp0.venv
if not exist "%VENV%\Scripts\python.exe" (
  echo [INFO] Creando entorno local...
  "%PYEXE%" -m venv "%VENV%"
  if errorlevel 1 (echo [ERROR] No se pudo crear el entorno Python.& pause& exit /b 1)
)
set PY=%VENV%\Scripts\python.exe

echo [INFO] Preparando dependencias...
"%PY%" -m pip install --disable-pip-version-check --no-warn-script-location -r "%~dp0backend\requirements.txt"
if errorlevel 1 (echo [ERROR] No se pudieron instalar las dependencias.& pause& exit /b 1)

REM 3) Iniciar unicamente FastAPI. El frontend ya esta incluido; no requiere Node.js.
echo [INFO] Iniciando servidor...
start "Gestion Clase Backend" /min cmd /c "cd /d "%~dp0backend" && "%PY%" -m uvicorn app.main:app --host 127.0.0.1 --port 8000 > "%~dp0gestion_clase.log" 2>&1"

set URL=http://127.0.0.1:8000
for /L %%N in (1,1,30) do (
  timeout /t 1 /nobreak >nul
  powershell -NoProfile -Command "try { (Invoke-WebRequest -UseBasicParsing -Uri '%URL%/api/health' -TimeoutSec 2).StatusCode } catch { 0 }" | findstr "200" >nul && goto READY
)
echo [ERROR] El servidor no respondio. Revisa gestion_clase.log
pause
exit /b 1

:READY
echo [OK] Sistema iniciado correctamente.
start "" "%URL%"
echo.
echo El sistema esta abierto en el navegador.
echo Puede cerrar esta ventana; el servidor seguira ejecutandose.
echo Para detenerlo use DETENER_GESTION_CLASE.bat
exit /b 0

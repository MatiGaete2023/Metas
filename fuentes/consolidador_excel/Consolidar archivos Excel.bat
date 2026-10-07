@echo off
setlocal EnableExtensions DisableDelayedExpansion
title Consolidar Excel - Espera y Cumplimiento
cd /d "%~dp0"

set "CONSOLIDADOR_PYTHON="
where py >nul 2>nul
if not errorlevel 1 (
  py -3 -c "import sys; raise SystemExit(0 if sys.version_info >= (3,10) else 1)" >nul 2>nul
  if not errorlevel 1 set "CONSOLIDADOR_PYTHON=py -3"
)
if not defined CONSOLIDADOR_PYTHON (
  where python >nul 2>nul
  if not errorlevel 1 (
    python -c "import sys; raise SystemExit(0 if sys.version_info >= (3,10) else 1)" >nul 2>nul
    if not errorlevel 1 set "CONSOLIDADOR_PYTHON=python"
  )
)
if not defined CONSOLIDADOR_PYTHON (
  echo No se encontro Python. Instala Python 3.10 o superior.
  pause
  exit /b 1
)

%CONSOLIDADOR_PYTHON% -c "import openpyxl, xlrd" >nul 2>nul
if errorlevel 1 (
  echo Faltan dependencias. Ejecuta primero Instalar dependencias.bat.
  pause
  exit /b 1
)

if "%~1"=="" (
  %CONSOLIDADOR_PYTHON% "%~dp0consolidar_excels.py"
) else (
  %CONSOLIDADOR_PYTHON% "%~dp0consolidar_excels.py" "%~1"
)
set "CONSOLIDADOR_RC=%ERRORLEVEL%"

if "%CONSOLIDADOR_RC%"=="3" (
  echo Proceso cancelado.
  pause
  exit /b 0
)
if not "%CONSOLIDADOR_RC%"=="0" (
  echo.
  echo El proceso termino con errores. Revisa los mensajes anteriores.
  pause
  exit /b %CONSOLIDADOR_RC%
)

echo.
pause
exit /b 0

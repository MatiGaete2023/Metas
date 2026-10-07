@echo off
setlocal EnableExtensions DisableDelayedExpansion
title Instalar dependencias - Consolidador Excel
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

%CONSOLIDADOR_PYTHON% -m pip install -r "%~dp0requirements.txt"
set "CONSOLIDADOR_RC=%ERRORLEVEL%"
if not "%CONSOLIDADOR_RC%"=="0" (
  echo.
  echo No fue posible instalar las dependencias.
  pause
  exit /b %CONSOLIDADOR_RC%
)

echo.
echo Dependencias instaladas correctamente.
pause
exit /b 0

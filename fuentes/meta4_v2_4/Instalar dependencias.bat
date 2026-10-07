@echo off
chcp 65001 >nul
setlocal EnableExtensions DisableDelayedExpansion
title Instalar dependencias - META 4

set "META4_DIR=%~dp0"
set "META4_REQUIREMENTS=%META4_DIR%requirements.txt"
set "META4_ENTORNO=%META4_DIR%scripts\_entorno.bat"

if not exist "%META4_REQUIREMENTS%" (
  echo ERROR: no se encontro requirements.txt.
  echo Descomprime el paquete completo y conserva su estructura.
  pause
  exit /b 2
)
if not exist "%META4_ENTORNO%" (
  echo ERROR: no se encontro scripts\_entorno.bat.
  echo Descomprime el paquete completo y conserva la carpeta scripts.
  pause
  exit /b 2
)

call "%META4_ENTORNO%"
if errorlevel 2 goto install
if errorlevel 1 goto no_python
goto install

:no_python
echo.
echo No encontre Python 3.10 o posterior en este equipo.
echo Instala Python 3 desde python.org y vuelve a intentar.
pause
exit /b 1

:install
echo.
echo Instalando o actualizando dependencias...
if defined META4_LAUNCHER (
  py -3 -m pip install --upgrade -r "%META4_REQUIREMENTS%"
) else (
  "%META4_PYTHON%" -m pip install --upgrade -r "%META4_REQUIREMENTS%"
)

if errorlevel 1 (
  echo.
  echo No fue posible instalar las dependencias.
  echo Revisa la conexion, el proxy o los permisos de Python.
  pause
  exit /b 1
)

echo.
echo Dependencias instaladas correctamente.
pause
exit /b 0

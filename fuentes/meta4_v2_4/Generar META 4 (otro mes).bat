@echo off
chcp 65001 >nul
setlocal EnableExtensions DisableDelayedExpansion
title Generador META 4 v2.4 - otro mes

set "META4_DIR=%~dp0"
set "META4_SCRIPT=%META4_DIR%scripts\generate_meta4.py"
set "META4_ENTORNO=%META4_DIR%scripts\_entorno.bat"

if not exist "%META4_SCRIPT%" goto missing_program
if not exist "%META4_ENTORNO%" goto missing_program

call "%META4_ENTORNO%"
if errorlevel 2 (
  echo.
  echo Faltan pandas, openpyxl o xlrd.
  echo Ejecuta primero "Instalar dependencias.bat".
  pause
  exit /b 1
)
if errorlevel 1 (
  echo.
  echo No encontre Python 3.10 o posterior en este equipo.
  echo Instala Python 3 y luego ejecuta "Instalar dependencias.bat".
  pause
  exit /b 1
)

echo.
echo Escribe el mes que quieres generar, en formato AAAA-MM.
echo Ejemplo: 2026-07
echo.
set "META4_MONTH="
set /p "META4_MONTH=Mes: "

if "%META4_MONTH%"=="" (
  echo No escribiste un mes. Proceso cancelado.
  pause
  exit /b 0
)

echo.
echo Elige la carpeta con los informes en la ventana que aparecera.
if defined META4_LAUNCHER (
  py -3 "%META4_SCRIPT%" --pick-folder --month "%META4_MONTH%" --open-output
) else (
  "%META4_PYTHON%" "%META4_SCRIPT%" --pick-folder --month "%META4_MONTH%" --open-output
)

if errorlevel 3 (
  echo.
  echo Proceso cancelado. No se genero ningun archivo.
  pause
  exit /b 0
)
if errorlevel 2 (
  echo.
  echo El mes escrito no es valido. Debe ser AAAA-MM, por ejemplo 2026-07.
  pause
  exit /b 1
)
if errorlevel 1 (
  echo.
  echo Hubo un problema durante la generacion. Revisa el mensaje anterior.
  pause
  exit /b 1
)

echo.
echo Proceso terminado.
pause
exit /b 0

:missing_program
echo.
echo ERROR: no se encontro el programa principal.
echo.
echo Este archivo BAT no funciona si se copia por separado.
echo Descomprime el paquete completo y conserva la carpeta scripts.
echo.
echo Ruta esperada:
echo "%META4_SCRIPT%"
pause
exit /b 2

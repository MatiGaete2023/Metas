@echo off
chcp 65001 >nul
setlocal EnableExtensions DisableDelayedExpansion
title Generador META 4 v2.4

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

set "SOURCE_DIR=%~1"

if "%SOURCE_DIR%"=="" goto pick_folder
if "%SOURCE_DIR:~-1%"=="\" set "SOURCE_DIR=%SOURCE_DIR%."

if not exist "%SOURCE_DIR%\." (
  echo.
  echo No encontre la carpeta que se debe examinar:
  echo "%SOURCE_DIR%"
  pause
  exit /b 1
)

echo.
echo Examinando:
echo "%SOURCE_DIR%"
echo.
if defined META4_LAUNCHER (
  py -3 "%META4_SCRIPT%" --folder "%SOURCE_DIR%" --open-output
) else (
  "%META4_PYTHON%" "%META4_SCRIPT%" --folder "%SOURCE_DIR%" --open-output
)
goto finished

:pick_folder
echo.
echo Elige la carpeta con los informes en la ventana que aparecera.
if defined META4_LAUNCHER (
  py -3 "%META4_SCRIPT%" --pick-folder --open-output
) else (
  "%META4_PYTHON%" "%META4_SCRIPT%" --pick-folder --open-output
)

:finished
if errorlevel 3 (
  echo.
  echo Proceso cancelado. No se genero ningun archivo.
  pause
  exit /b 0
)
if errorlevel 1 (
  echo.
  echo Hubo un problema durante la generacion. Revisa el mensaje anterior.
  pause
  exit /b 1
)

echo.
echo Proceso terminado.
echo Los archivos quedaron en la carpeta examinada.
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

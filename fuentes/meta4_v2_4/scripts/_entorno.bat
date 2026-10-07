@echo off
rem Este archivo NO se ejecuta directamente: lo invocan los demas BAT con
rem "call". A proposito no usa setlocal, para que META4_LAUNCHER y
rem META4_PYTHON queden visibles en quien lo invoca una vez que termina.
rem
rem Codigo de salida:
rem   0 = listo (META4_LAUNCHER=1 para "py -3", o META4_PYTHON=interprete)
rem   1 = no se encontro Python 3.10 o superior
rem   2 = Python esta, pero faltan pandas, openpyxl o xlrd
rem
rem Unico punto de deteccion de Python para los tres lanzadores (M-03): antes
rem cada BAT probaba un orden distinto y solo uno de los tres validaba la
rem version, por lo que un alias de Microsoft Store o un Python antiguo podia
rem quedar seleccionado sin que el usuario lo notara.

set "META4_LAUNCHER="
set "META4_PYTHON="

where py >nul 2>nul
if not errorlevel 1 (
  py -3 -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>nul
  if not errorlevel 1 (
    set "META4_LAUNCHER=1"
    goto :meta4_python_found
  )
)

where python >nul 2>nul
if not errorlevel 1 (
  python -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>nul
  if not errorlevel 1 (
    set "META4_PYTHON=python"
    goto :meta4_python_found
  )
)

exit /b 1

:meta4_python_found
if defined META4_LAUNCHER (
  py -3 -c "import pandas, openpyxl, xlrd" >nul 2>nul
) else (
  "%META4_PYTHON%" -c "import pandas, openpyxl, xlrd" >nul 2>nul
)
if errorlevel 1 exit /b 2
exit /b 0

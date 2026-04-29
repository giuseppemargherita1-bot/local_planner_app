@echo off
setlocal
cd /d "%~dp0"

set "PLANNER_PORT=8876"
set "PORT=%PLANNER_PORT%"

echo =============================================
echo  DEBUG AVVIO MANPOWER PLANNER REV
echo =============================================
echo Cartella: %CD%
echo Porta richiesta: %PLANNER_PORT%
echo.

if exist "%~dp0pyembed\python.exe" (
  echo Uso Python embedded: %~dp0pyembed\python.exe
  "%~dp0pyembed\python.exe" "%~dp0app.py"
  goto fine
)

if exist "%~dp0python.exe" (
  echo Uso python.exe locale: %~dp0python.exe
  "%~dp0python.exe" "%~dp0app.py"
  goto fine
)

echo Uso py da Windows PATH
py "%~dp0app.py"

:fine
echo.
echo =============================================
echo  IL SERVER SI E' CHIUSO O NON E' PARTITO
echo  Copia/incolla qui il messaggio sopra.
echo =============================================
pause

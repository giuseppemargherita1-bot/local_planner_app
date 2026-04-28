@echo off
setlocal
cd /d "%~dp0"

set "PLANNER_URL=http://127.0.0.1:8765/fast.html"
set "EDGE_ARGS=--app=%PLANNER_URL% --new-window"

REM Avvia il server del planner senza usare il launcher EXE, cosi evitiamo l'apertura automatica di Chrome.
if exist "%~dp0pyembed\python.exe" (
  start "Manpower Planner Server" /min "%~dp0pyembed\python.exe" "%~dp0app.py"
) else if exist "%~dp0python.exe" (
  start "Manpower Planner Server" /min "%~dp0python.exe" "%~dp0app.py"
) else (
  start "Manpower Planner Server" /min py "%~dp0app.py"
)

REM Lascia qualche secondo al server locale per partire.
timeout /t 3 /nobreak >nul

REM Apri Edge in modalita app. Non usa Chrome e non mostra una normale scheda browser.
if exist "%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe" (
  start "Manpower Planner FAST" "%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe" %EDGE_ARGS%
  exit /b 0
)

if exist "%ProgramFiles%\Microsoft\Edge\Application\msedge.exe" (
  start "Manpower Planner FAST" "%ProgramFiles%\Microsoft\Edge\Application\msedge.exe" %EDGE_ARGS%
  exit /b 0
)

where msedge >nul 2>nul
if %ERRORLEVEL%==0 (
  start "Manpower Planner FAST" msedge %EDGE_ARGS%
  exit /b 0
)

echo Microsoft Edge non trovato.
echo Apri manualmente: %PLANNER_URL%
pause

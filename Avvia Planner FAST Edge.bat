@echo off
setlocal
cd /d "%~dp0"

set "PLANNER_PORT=8876"
set "PLANNER_URL=http://127.0.0.1:%PLANNER_PORT%/fast.html"
set "EDGE_ARGS=--app=%PLANNER_URL% --new-window"

REM Avvia il server del planner su una porta alternativa per non interferire con altre app.
if exist "%~dp0pyembed\python.exe" (
  start "Manpower Planner Server REV" /min cmd /c "set PORT=%PLANNER_PORT%&& "%~dp0pyembed\python.exe" "%~dp0app.py""
) else if exist "%~dp0python.exe" (
  start "Manpower Planner Server REV" /min cmd /c "set PORT=%PLANNER_PORT%&& "%~dp0python.exe" "%~dp0app.py""
) else (
  start "Manpower Planner Server REV" /min cmd /c "set PORT=%PLANNER_PORT%&& py "%~dp0app.py""
)

REM Lascia qualche secondo al server locale per partire.
timeout /t 5 /nobreak >nul

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

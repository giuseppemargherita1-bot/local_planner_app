@echo off
setlocal

set APP_DIR=%~dp0
set PY_EMBEDW=%APP_DIR%pyembed\pythonw.exe
set PY_EMBED=%APP_DIR%pyembed\python.exe

if exist "%PY_EMBEDW%" (
  set PY_BIN=%PY_EMBEDW%
) else if exist "%PY_EMBED%" (
  set PY_BIN=%PY_EMBED%
) else (
  set PY_BIN=python
)

cd /d "%APP_DIR%"
start "" http://127.0.0.1:8765
"%PY_BIN%" app.py

endlocal

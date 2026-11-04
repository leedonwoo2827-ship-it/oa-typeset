@echo off
rem oa-typeset studio - http://127.0.0.1:7090  (other port: set PORT=7091 then run)
setlocal
cd /d "%~dp0"
set PYTHONUTF8=1
if "%PORT%"=="" set PORT=7090
if not exist ".venv\Scripts\python.exe" (
  echo .venv not found - run setup.bat first.
  pause
  exit /b 1
)
start "" http://127.0.0.1:%PORT%
".venv\Scripts\python.exe" -m oa_typeset studio --port %PORT%
pause

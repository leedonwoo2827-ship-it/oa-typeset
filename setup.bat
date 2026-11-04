@echo off
rem oa-typeset setup - run once. Creates .venv and installs everything.
setlocal
cd /d "%~dp0"
set PY=python
where py >nul 2>&1 && set PY=py -3
if not exist ".venv\Scripts\python.exe" (
  echo [1/3] Creating .venv ...
  %PY% -m venv .venv
  if errorlevel 1 goto err
)
echo [2/3] Installing packages (a few minutes) ...
".venv\Scripts\python.exe" -m pip install -q -U pip
".venv\Scripts\python.exe" -m pip install -e ".[all]"
if errorlevel 1 goto err
echo [3/3] LLM login status:
set PYTHONUTF8=1
".venv\Scripts\python.exe" -m oa_typeset setup-check
echo.
echo Done. Next: run-studio.bat
pause
exit /b 0
:err
echo.
echo SETUP FAILED - see the messages above.
pause
exit /b 1

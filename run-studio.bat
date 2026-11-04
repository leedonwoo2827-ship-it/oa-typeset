@echo off
chcp 65001 >nul
rem 로컬 웹 스튜디오 — http://127.0.0.1:7090  (PORT 를 바꾸려면: set PORT=7091 ^&^& run-studio.bat)
setlocal
cd /d "%~dp0"
if exist .venv\Scripts\activate.bat call .venv\Scripts\activate.bat
if "%PORT%"=="" set PORT=7090
start "" http://127.0.0.1:%PORT%
python -m oa_typeset studio --port %PORT%
pause

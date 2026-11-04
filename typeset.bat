@echo off
rem Drag .html (or .md) files onto this file -> hwpx, docx, pptx next to each.
rem   Options: set OA_FORM=<form folder name>   set OA_FORMATS=hwpx,docx
setlocal
cd /d "%~dp0"
set PYTHONUTF8=1
if "%OA_FORMATS%"=="" set OA_FORMATS=hwpx,docx,pptx
set FORMARG=
if not "%OA_FORM%"=="" set FORMARG=--form "%OA_FORM%"
set PYEXE=python
if exist ".venv\Scripts\python.exe" set PYEXE=".venv\Scripts\python.exe"
if "%~1"=="" (
  echo Drag .html files onto this .bat file.
  pause
  exit /b 1
)
for %%F in (%*) do %PYEXE% -m oa_typeset build "%%~fF" %FORMARG% -f %OA_FORMATS%
pause

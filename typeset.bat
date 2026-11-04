@echo off
chcp 65001 >nul
rem 끌어다 놓기: .html(.md) 파일을 이 파일 위에 놓으면 옆에 hwpx · docx · pptx 가 생긴다.
rem   양식 바꾸기: set OA_FORM=보고서형-A4   /  형식 바꾸기: set OA_FORMATS=hwpx
setlocal
cd /d "%~dp0"
if exist .venv\Scripts\activate.bat call .venv\Scripts\activate.bat
if "%OA_FORM%"=="" set OA_FORM=교재형-46배판
if "%OA_FORMATS%"=="" set OA_FORMATS=hwpx,docx,pptx
if "%~1"=="" (
  echo .html 파일을 이 .bat 위에 끌어다 놓으세요.
  pause
  exit /b 1
)
for %%F in (%*) do python -m oa_typeset build "%%~fF" --form "%OA_FORM%" -f %OA_FORMATS%
pause

@echo off
chcp 65001 >nul
rem oa-typeset 설치 — 처음 한 번. 가상환경(.venv)을 만들고 모두 설치한다.
setlocal
cd /d "%~dp0"
where py >nul 2>&1 && (set PY=py -3) || (set PY=python)
if not exist .venv (%PY% -m venv .venv || goto :err)
call .venv\Scripts\activate.bat
python -m pip install -U pip >nul
python -m pip install -e ".[all]" || goto :err
echo.
echo 설치 끝. LLM 집필을 쓰려면 한 번 로그인하세요:
echo   Codex  : npm i -g @openai/codex  그리고  codex login     (ChatGPT 계정)
echo   Claude : Claude Code 설치 후  claude auth login           (Claude 구독)
echo 조판만 쓸 거면 로그인은 필요 없습니다.
pause
exit /b 0
:err
echo 설치 실패. 위 메시지를 확인하세요.
pause
exit /b 1

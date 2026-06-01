@echo off
:: 고객지표 RPA — 초기 설치 및 설정 마법사
:: 더블클릭으로 실행하세요. 완료 후 설정 창이 열립니다.

cd /d "%~dp0"
title 고객지표 RPA 설치 중...

echo.
echo ================================================
echo   고객지표 RPA 설치 및 설정
echo ================================================
echo.

:: Python 설치 확인
python --version >nul 2>&1
if errorlevel 1 (
    echo [오류] Python이 설치되어 있지 않습니다.
    echo        https://www.python.org 에서 Python 3.10 이상을 설치하세요.
    pause
    exit /b 1
)

echo [1/4] 가상환경 생성 중...
if not exist ".venv\" (
    python -m venv .venv
    echo       완료
) else (
    echo       이미 존재함 - 건너뜀
)

echo [2/4] 패키지 설치 중... (처음 실행 시 1-2분 소요)
.venv\Scripts\pip install -r requirements.txt --quiet --disable-pip-version-check
if errorlevel 1 (
    echo [오류] 패키지 설치 실패. 인터넷 연결을 확인하세요.
    pause
    exit /b 1
)
echo       완료

echo [3/4] 브라우저 설치 중... (처음 실행 시 1-2분 소요)
.venv\Scripts\python -m playwright install chromium >nul 2>&1
echo       완료

echo [4/4] 설정 마법사 실행 중...
if not exist ".env" (
    copy .env.example .env >nul
)

echo.
.venv\Scripts\python -m src.setup_gui

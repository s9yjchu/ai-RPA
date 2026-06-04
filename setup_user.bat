@echo off
chcp 65001 > nul
echo ============================================================
echo   B2C 고객지표 RPA 설치 및 설정
echo ============================================================
echo.

cd /d "%~dp0"

REM ── 1. Python 설치 확인 ─────────────────────────────────────────────
python --version > nul 2>&1
if errorlevel 1 (
    echo [필요] Python 이 설치되어 있지 않습니다.
    echo        https://python.org 에서 Python 3.12 를 먼저 설치하세요.
    echo        설치 시 "Add python.exe to PATH" 체크박스를 반드시 선택하세요.
    pause
    exit /b 1
)
echo [OK] Python 확인 완료

REM ── 2. 가상환경 생성 및 패키지 설치 ─────────────────────────────────
if not exist ".venv\Scripts\python.exe" (
    echo [설치] 가상환경 생성 중...
    python -m venv .venv
)
echo [설치] 패키지 설치 중 (최초 실행 시 3~5분 소요)...
.venv\Scripts\python.exe -m pip install -r requirements.txt -q
if errorlevel 1 (
    echo [오류] 패키지 설치 실패. 인터넷 연결을 확인하세요.
    pause
    exit /b 1
)
echo [OK] 패키지 설치 완료

REM ── 3. Playwright Chromium 설치 ──────────────────────────────────────
echo [설치] 브라우저 설치 중...
.venv\Scripts\python.exe -m playwright install chromium
echo [OK] 브라우저 설치 완료

REM ── 4. SPC Hub 로그인 정보 입력 ─────────────────────────────────────
echo.
echo ── SPC Hub 로그인 정보 입력 ────────────────────────────────────────
echo    (본인의 SPC 사내 아이디/비밀번호를 입력하세요)
echo.
set /p HUB_ID=SPC Hub 아이디:
set /p HUB_PW=SPC Hub 비밀번호:

REM .env 파일 업데이트
powershell -Command "(Get-Content .env) -replace 'SPCHUB_ID=.*', 'SPCHUB_ID=%HUB_ID%' | Set-Content .env"
powershell -Command "(Get-Content .env) -replace 'SPCHUB_PW=.*', 'SPCHUB_PW=%HUB_PW%' | Set-Content .env"
echo [OK] SPC Hub 정보 저장 완료

REM ── 5. Google 인증 (Sheets 쓰기용) ──────────────────────────────────
echo.
echo ── Google 계정 인증 ────────────────────────────────────────────────
echo    브라우저가 열리면 본인의 Google 계정으로 로그인하고 [허용]을 클릭하세요.
echo    (Google Sheets 에 접근하는 계정: B2C사업본부 고객지표 시트 편집 권한 필요)
echo.
if not exist "token.json" (
    .venv\Scripts\python.exe -c "
from src.sheets_writer import get_client
from pathlib import Path
try:
    get_client(Path('./credentials.json'), Path('./token.json'))
    print('[OK] Google 인증 완료')
except Exception as e:
    print('[오류]', e)
"
) else (
    echo [OK] Google 인증 이미 완료됨 (token.json 존재)
)

REM ── 6. 동작 테스트 ───────────────────────────────────────────────────
echo.
echo ── 연결 테스트 ─────────────────────────────────────────────────────
.venv\Scripts\python.exe -c "
import sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
from src.config import load_config
try:
    c = load_config()
    print('[OK] 설정 로드:', c.hub.user_id)
    print('[OK] GCP 프로젝트:', c.gcp.project_id)
    print('[OK] Sheets ID:', c.sheets.spreadsheet_id[:20] + '...')
except Exception as e:
    print('[오류]', e)
    sys.exit(1)
"
if errorlevel 1 (
    echo [오류] 설정 확인 실패. .env 파일을 확인하세요.
    pause
    exit /b 1
)

REM ── 7. Windows 에이전트 서비스 등록 (관리자 권한 필요) ───────────────
echo.
echo ── 자동 실행 등록 ──────────────────────────────────────────────────
echo    GCP 에이전트를 시작 프로그램으로 등록합니다.
echo    (더 안정적인 Windows 서비스 등록을 원하면 install_agent_service.bat 을 관리자 권한으로 실행하세요)
echo.

REM Task Scheduler 등록 (현재 사용자 로그인 시 자동 시작)
set TASKNAME=B2C_RPA_GCP_Agent
set PYTHON_PATH=%CD%\.venv\Scripts\python.exe
set WORK_DIR=%CD%

schtasks /query /tn "%TASKNAME%" > nul 2>&1
if not errorlevel 1 (
    echo [OK] 에이전트 작업이 이미 등록되어 있습니다.
) else (
    schtasks /create /tn "%TASKNAME%" /tr "\"%PYTHON_PATH%\" -m src.gcp_agent" /sc onlogon /ru "%USERNAME%" /f > nul 2>&1
    if errorlevel 1 (
        echo [경고] Task Scheduler 등록 실패. install_agent_service.bat 을 관리자 권한으로 실행하세요.
    ) else (
        echo [OK] 에이전트 등록 완료 (로그인 시 자동 시작)
    )
)

REM 에이전트 즉시 시작
schtasks /run /tn "%TASKNAME%" > nul 2>&1
echo [OK] 에이전트 시작됨

echo.
echo ============================================================
echo   설치 완료!
echo.
echo   자동 실행: 매일 08:30 GCP 에서 실행 신호를 보내며
echo              이 PC 가 켜져 있을 때 자동으로 RPA 가 실행됩니다.
echo.
echo   수동 실행: run_rpa.bat 더블클릭
echo   사용 가이드: 사용법.md 참조
echo ============================================================
pause

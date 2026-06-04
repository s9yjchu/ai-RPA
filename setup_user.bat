@echo off
chcp 65001 > nul
echo ============================================================
echo   B2C 고객지표 RPA 설치 프로그램
echo ============================================================
echo.

cd /d "%~dp0"

REM ── 0. credentials.json 확인 ────────────────────────────────────────
if not exist "credentials.json" (
    echo.
    echo [오류] credentials.json 파일이 없습니다.
    echo.
    echo        관리자에게 credentials.json 파일을 요청하여
    echo        이 폴더에 복사한 뒤 다시 실행하세요.
    echo.
    pause
    exit /b 1
)
echo [OK] credentials.json 확인

REM ── 1. Python 확인 ──────────────────────────────────────────────────
python --version > nul 2>&1
if errorlevel 1 (
    echo [오류] Python 이 설치되어 있지 않습니다.
    echo.
    echo        1. https://python.org 접속
    echo        2. Python 3.12 다운로드 및 설치
    echo        3. 설치 시 "Add python.exe to PATH" 반드시 체크
    echo        4. 설치 완료 후 이 파일을 다시 실행
    echo.
    pause
    exit /b 1
)
for /f "tokens=2" %%v in ('python --version') do echo [OK] Python %%v 확인 완료

REM ── 2. 가상환경 생성 ────────────────────────────────────────────────
if not exist ".venv\Scripts\python.exe" (
    echo [설치] 가상환경 생성 중...
    python -m venv .venv
    if errorlevel 1 (
        echo [오류] 가상환경 생성 실패
        pause & exit /b 1
    )
)
echo [OK] 가상환경 확인

REM ── 3. 패키지 설치 ──────────────────────────────────────────────────
echo [설치] 필요 패키지 설치 중 (최초 실행 시 3~5분 소요)...
.venv\Scripts\python.exe -m pip install -r requirements.txt -q --no-warn-script-location
if errorlevel 1 (
    echo [오류] 패키지 설치 실패 — 인터넷 연결을 확인하세요.
    pause & exit /b 1
)
echo [OK] 패키지 설치 완료

REM ── 4. 브라우저 설치 ────────────────────────────────────────────────
echo [설치] Chromium 브라우저 설치 중 (최초 실행 시 2~3분 소요)...
.venv\Scripts\python.exe -m playwright install chromium
if errorlevel 1 (
    echo [오류] 브라우저 설치 실패
    pause & exit /b 1
)
echo [OK] 브라우저 설치 완료

REM ── 5. SPC Hub 로그인 정보 입력 ─────────────────────────────────────
echo.
echo ── SPC Hub 로그인 정보 ─────────────────────────────────────────────
echo    본인의 SPC 사내 아이디와 비밀번호를 입력하세요.
echo    (비밀번호는 입력해도 화면에 표시되지 않습니다)
echo.
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$id = Read-Host 'SPC Hub 아이디 '; if (-not $id) { Write-Host '[오류] 아이디를 입력하세요.'; exit 1 }; $sec = Read-Host 'SPC Hub 비밀번호' -AsSecureString; $pw = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)); if (-not $pw) { Write-Host '[오류] 비밀번호를 입력하세요.'; exit 1 }; $env = Get-Content '.env' -Encoding UTF8; $env = $env -replace '^SPCHUB_ID=.*', ('SPCHUB_ID=' + $id); $env = $env -replace '^SPCHUB_PW=.*', ('SPCHUB_PW=' + $pw); $env | Set-Content '.env' -Encoding UTF8; Write-Host '[OK] SPC Hub 정보 저장 완료'"
if %errorlevel% neq 0 (
    echo [오류] 로그인 정보 저장 실패. 다시 시도하세요.
    pause & exit /b 1
)

REM ── 6. Google 계정 인증 ─────────────────────────────────────────────
echo.
echo ── Google 계정 인증 ────────────────────────────────────────────────
if exist "token.json" (
    echo [OK] Google 인증 파일 확인 완료
) else (
    echo    브라우저가 열립니다. Google 계정으로 로그인 후 [허용]을 클릭하세요.
    echo    (B2C사업본부 고객지표 스프레드시트 편집 권한이 있는 계정)
    echo.
    .venv\Scripts\python.exe -c "
import sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
from src.sheets_writer import get_client
from pathlib import Path
try:
    get_client(Path('./credentials.json'), Path('./token.json'))
    print('[OK] Google 인증 완료')
except Exception as e:
    print('[오류]', e)
    sys.exit(1)
"
    if errorlevel 1 (
        echo [오류] Google 인증 실패. 다시 시도해주세요.
        pause & exit /b 1
    )
)

REM ── 7. 설정 확인 ────────────────────────────────────────────────────
echo.
echo [확인] 설정 검증 중...
.venv\Scripts\python.exe -c "
import sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
from src.config import load_config
try:
    c = load_config()
    print('[OK] SPC Hub 계정:', c.hub.user_id)
    print('[OK] Sheets 연결 설정 확인')
except Exception as e:
    print('[오류]', e)
    sys.exit(1)
"
if errorlevel 1 (
    echo [오류] 설정 확인 실패
    pause & exit /b 1
)

REM ── 8. 자동 실행 등록 (Task Scheduler) ─────────────────────────────
echo.
echo [등록] 자동 실행 일정 등록 중...

set PYEXE=%CD%\.venv\Scripts\python.exe
set TASKBASE=B2C_RPA_Daily

REM 기존 작업 제거
for %%H in (0830 0900 0930 1000 1030 1100 1130) do (
    schtasks /delete /tn "%TASKBASE%_%%H" /f > nul 2>&1
)

REM 매일 08:30~11:30 (30분 간격) 작업 등록
set TIMES=08:30 09:00 09:30 10:00 10:30 11:00 11:30
set CODES=0830 0900 0930 1000 1030 1100 1130

setlocal enabledelayedexpansion
set i=0
for %%T in (%TIMES%) do (
    set /a i+=1
    set IDX=0000!i!
    for %%H in (!CODES!) do (
        if !i!==1 (
            schtasks /create /tn "%TASKBASE%_%%H" /tr "\"%PYEXE%\" -m src.main daily" /sc daily /st %%T /mo 1 /ru "%USERNAME%" /f > nul 2>&1
        )
        set /a i-=1
        if !i!==0 goto :next
    )
    :next
)
endlocal

REM 간단하게 7개 작업 직접 등록
schtasks /create /tn "%TASKBASE%_0830" /tr "\"%PYEXE%\" -m src.main daily" /sc daily /st 08:30 /ru "%USERNAME%" /f > nul 2>&1
schtasks /create /tn "%TASKBASE%_0900" /tr "\"%PYEXE%\" -m src.main daily" /sc daily /st 09:00 /ru "%USERNAME%" /f > nul 2>&1
schtasks /create /tn "%TASKBASE%_0930" /tr "\"%PYEXE%\" -m src.main daily" /sc daily /st 09:30 /ru "%USERNAME%" /f > nul 2>&1
schtasks /create /tn "%TASKBASE%_1000" /tr "\"%PYEXE%\" -m src.main daily" /sc daily /st 10:00 /ru "%USERNAME%" /f > nul 2>&1
schtasks /create /tn "%TASKBASE%_1030" /tr "\"%PYEXE%\" -m src.main daily" /sc daily /st 10:30 /ru "%USERNAME%" /f > nul 2>&1
schtasks /create /tn "%TASKBASE%_1100" /tr "\"%PYEXE%\" -m src.main daily" /sc daily /st 11:00 /ru "%USERNAME%" /f > nul 2>&1
schtasks /create /tn "%TASKBASE%_1130" /tr "\"%PYEXE%\" -m src.main daily" /sc daily /st 11:30 /ru "%USERNAME%" /f > nul 2>&1

REM 월별 — 매월 1~4일 08:30
schtasks /create /tn "B2C_RPA_Monthly" /tr "\"%PYEXE%\" -m src.main monthly" /sc monthly /d 1 /st 08:30 /ru "%USERNAME%" /f > nul 2>&1

echo [OK] 자동 실행 등록 완료 (08:30~11:30 매일, 월별 매월 1일)

echo.
echo ============================================================
echo   설치 완료!
echo.
echo   - 매일 08:30 자동 실행됩니다 (PC 가 켜져 있어야 합니다)
echo   - 수동 실행: run_rpa.bat 더블클릭
echo   - 자세한 사용법: 설치_사용_가이드.md 참조
echo ============================================================
pause

@echo off
chcp 65001 > nul
echo ============================================================
echo   B2C 고객지표 RPA 설치 프로그램
echo ============================================================
echo.

cd /d "%~dp0"

REM ── 0-a. 설치 경로 ASCII 확인 ────────────────────────────────────────
REM Node.js 가 한글/특수문자 경로에서 충돌하므로 ASCII 경로 필수
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$p='%~dp0'; if($p -match '[^\x00-\x7F]'){Write-Host '[오류] 설치 경로에 한글/특수문자가 포함되어 있습니다.'; Write-Host ''; Write-Host '       C:\RPA\ 같은 영문 경로에 압축을 풀고 다시 실행하세요.'; Write-Host '       현재 경로:' $p; exit 1}else{exit 0}" > nul 2>&1
if errorlevel 1 (
    echo.
    echo [오류] 설치 경로에 한글 또는 특수문자가 포함되어 있습니다.
    echo.
    echo        예^) C:\RPA\ 또는 C:\Users\andrew\RPA\ 같은
    echo            영문 경로에 압축을 풀고 다시 실행하세요.
    echo.
    pause & exit /b 1
)
echo [OK] 설치 경로 확인

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

REM ── 1. Python 3.12 확인/설치 ────────────────────────────────────────────
echo [확인] Python 버전 확인 중...

REM py 런처로 3.12 직접 확인
py -3.12 --version > nul 2>&1
if not errorlevel 1 (
    for /f "tokens=2" %%v in ('py -3.12 --version') do echo [OK] Python %%v 확인 완료
    set PYTHON_CMD=py -3.12
    goto :python_ok
)

REM python 명령 버전 확인 (3.11 또는 3.12만 허용)
powershell -NoProfile -ExecutionPolicy Bypass -Command "try{$v=(python --version 2>&1);if($v -match 'Python 3\.(\d+)\.'){$m=[int]$Matches[1];if($m -ge 11 -and $m -le 12){exit 0}}exit 1}catch{exit 1}" > nul 2>&1
if not errorlevel 1 (
    for /f "tokens=2" %%v in ('python --version') do echo [OK] Python %%v 확인 완료
    set PYTHON_CMD=python
    goto :python_ok
)

REM Python 3.12 설치 필요 — 번들 설치파일 사용
echo [설치] Python 3.12 설치 중 (약 1~2분 소요, 화면이 멈춘 것처럼 보여도 정상)...
if not exist "%~dp0python-3.12.9-amd64.exe" (
    echo [오류] python-3.12.9-amd64.exe 파일이 없습니다. 관리자에게 문의하세요.
    pause & exit /b 1
)
"%~dp0python-3.12.9-amd64.exe" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0
if errorlevel 1 (
    echo [오류] Python 3.12 설치 실패
    pause & exit /b 1
)
echo [OK] Python 3.12 설치 완료
set PYTHON_CMD=py -3.12

:python_ok

REM ── 2. 가상환경 생성 ────────────────────────────────────────────────
if not exist ".venv\Scripts\python.exe" (
    echo [설치] 가상환경 생성 중...
    %PYTHON_CMD% -m venv .venv
    if errorlevel 1 (
        echo [오류] 가상환경 생성 실패
        pause & exit /b 1
    )
)
REM pip 누락 시 bootstrap (Python 3.12+ 일부 환경에서 venv에 pip 미포함)
.venv\Scripts\python.exe -m pip --version > nul 2>&1
if errorlevel 1 (
    echo [설치] pip 초기화 중...
    .venv\Scripts\python.exe -m ensurepip --upgrade
    if errorlevel 1 (
        echo [오류] pip 초기화 실패
        pause & exit /b 1
    )
)
echo [OK] 가상환경 확인

REM ── 3. 패키지 설치 ──────────────────────────────────────────────────
echo [설치] 필요 패키지 설치 중 (최초 실행 시 3~5분 소요)...
.venv\Scripts\python.exe -m pip install -r requirements_user.txt -q --no-warn-script-location
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
echo    B2C사업본부 고객지표 스프레드시트에 접근하는 Google 계정으로 인증합니다.
echo    SPC Hub 아이디와 다를 수 있습니다. 시트 편집 권한이 있는 이메일을 입력하세요.
echo.
set /p GOOGLE_EMAIL="   Google 이메일 주소: "
echo.

if exist "token.json" (
    echo [OK] Google 인증 파일 확인 완료
) else (
    echo    브라우저가 열립니다. %GOOGLE_EMAIL% 계정으로 로그인 후 [허용]을 클릭하세요.
    echo.
    .venv\Scripts\python.exe -m src.setup_helper auth "%GOOGLE_EMAIL%"
    if errorlevel 1 (
        echo [오류] Google 인증 실패. 다시 시도해주세요.
        pause & exit /b 1
    )
)

REM ── 7. 설정 확인 ────────────────────────────────────────────────────
echo.
echo [확인] 설정 검증 중...
.venv\Scripts\python.exe -m src.setup_helper validate
if errorlevel 1 (
    echo [오류] 설정 확인 실패
    pause & exit /b 1
)

REM ── 8. 자동 실행 등록 (Task Scheduler) ─────────────────────────────
echo.
echo [등록] 자동 실행 일정 등록 중...

set TASKBASE=B2C_RPA_Daily
set RUNBAT=%~dp0run_rpa.bat

REM 기존 작업 제거 (이전 버전 0830 슬롯 포함)
for %%H in (0830 0900 0930 1000 1030 1100 1130 1200) do (
    schtasks /delete /tn "%TASKBASE%_%%H" /f > nul 2>&1
)

REM 매일 09:00~12:00 (30분 간격) 작업 등록 — OLAP 09:00 오픈 기준
REM run_rpa.bat 이 cd /d 로 작업 디렉토리를 올바르게 설정함
schtasks /create /tn "%TASKBASE%_0900" /tr "\"%RUNBAT%\" daily /auto" /sc daily /st 09:00 /ru "%USERNAME%" /f > nul 2>&1
schtasks /create /tn "%TASKBASE%_0930" /tr "\"%RUNBAT%\" daily /auto" /sc daily /st 09:30 /ru "%USERNAME%" /f > nul 2>&1
schtasks /create /tn "%TASKBASE%_1000" /tr "\"%RUNBAT%\" daily /auto" /sc daily /st 10:00 /ru "%USERNAME%" /f > nul 2>&1
schtasks /create /tn "%TASKBASE%_1030" /tr "\"%RUNBAT%\" daily /auto" /sc daily /st 10:30 /ru "%USERNAME%" /f > nul 2>&1
schtasks /create /tn "%TASKBASE%_1100" /tr "\"%RUNBAT%\" daily /auto" /sc daily /st 11:00 /ru "%USERNAME%" /f > nul 2>&1
schtasks /create /tn "%TASKBASE%_1130" /tr "\"%RUNBAT%\" daily /auto" /sc daily /st 11:30 /ru "%USERNAME%" /f > nul 2>&1
schtasks /create /tn "%TASKBASE%_1200" /tr "\"%RUNBAT%\" daily /auto" /sc daily /st 12:00 /ru "%USERNAME%" /f > nul 2>&1

REM 월별 — 매월 1~4일 09:00
schtasks /create /tn "B2C_RPA_Monthly" /tr "\"%RUNBAT%\" monthly /auto" /sc monthly /d 1 /st 09:00 /ru "%USERNAME%" /f > nul 2>&1

echo [OK] 자동 실행 등록 완료 (09:00~12:00 매일, 월별 매월 1일)

REM ── 9. 바탕화면 바로가기 생성 ────────────────────────────────────────
echo [등록] 바탕화면 바로가기 생성 중...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$s=(New-Object -ComObject WScript.Shell).CreateShortcut([Environment]::GetFolderPath('Desktop')+'\B2C 고객지표 수동실행.lnk'); $s.TargetPath='%~dp0run_rpa.bat'; $s.WorkingDirectory='%~dp0'; $s.Description='B2C 고객지표 RPA 수동 실행'; $s.Save()" > nul 2>&1
echo [OK] 바탕화면에 "B2C 고객지표 수동실행" 바로가기 생성 완료

echo.
echo ============================================================
echo   설치 완료!
echo.
echo   - 매일 09:00 자동 실행됩니다 (PC 가 켜져 있어야 합니다)
echo   - 수동 실행: 바탕화면의 "B2C 고객지표 수동실행" 더블클릭
echo   - 자세한 사용법: 설치_사용_가이드.md 참조
echo ============================================================
pause

@echo off
REM ============================================================
REM  B2C RPA GCP 에이전트 — Windows 서비스 등록 (NSSM 사용)
REM  관리자 권한으로 실행하세요 (우클릭 → 관리자 권한으로 실행)
REM
REM  NSSM 다운로드: https://nssm.cc/download
REM  nssm.exe 를 이 배치 파일과 같은 폴더(프로젝트 루트)에 놓으세요.
REM ============================================================

setlocal
set PROJECT_DIR=%~dp0
set SERVICE_NAME=B2C_RPA_Agent
set PYTHON=%PROJECT_DIR%.venv\Scripts\python.exe

REM NSSM 존재 확인
if not exist "%PROJECT_DIR%nssm.exe" (
    echo [오류] nssm.exe 가 없습니다. https://nssm.cc/download 에서 받아서
    echo       %PROJECT_DIR% 에 복사하세요.
    pause
    exit /b 1
)

echo [설치] 서비스 '%SERVICE_NAME%' 등록 중...

nssm install %SERVICE_NAME% "%PYTHON%"
nssm set %SERVICE_NAME% AppParameters "-m src.gcp_agent"
nssm set %SERVICE_NAME% AppDirectory "%PROJECT_DIR%"
nssm set %SERVICE_NAME% AppStdout "%PROJECT_DIR%logs\agent_stdout.log"
nssm set %SERVICE_NAME% AppStderr "%PROJECT_DIR%logs\agent_stderr.log"
nssm set %SERVICE_NAME% AppRotateFiles 1
nssm set %SERVICE_NAME% AppRotateOnline 1
nssm set %SERVICE_NAME% AppRotateSeconds 86400
nssm set %SERVICE_NAME% Start SERVICE_AUTO_START
nssm set %SERVICE_NAME% DisplayName "B2C RPA GCP Agent"
nssm set %SERVICE_NAME% Description "GCP Pub/Sub 메시지를 수신하여 B2C 고객지표 RPA를 실행합니다."

echo [시작] 서비스 시작 중...
nssm start %SERVICE_NAME%

echo.
echo [완료] 서비스가 등록되었습니다.
echo.
echo  상태 확인 : nssm status %SERVICE_NAME%
echo  재시작    : nssm restart %SERVICE_NAME%
echo  중지      : nssm stop %SERVICE_NAME%
echo  제거      : nssm remove %SERVICE_NAME% confirm
echo.
pause

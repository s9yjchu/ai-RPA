@echo off
chcp 65001 > nul
:: 고객지표 RPA — 수동 실행 / Task Scheduler 진입점

setlocal enabledelayedexpansion
cd /d "%~dp0"

set MODE=%~1
if "%MODE%"=="" set MODE=daily

:: /auto 플래그 감지 (Task Scheduler 에서 전달) — Python 에는 넘기지 않음
set AUTO=0
set PY_ARGS=
for %%A in (%~2 %~3 %~4 %~5) do (
    if /I "%%A"=="/auto" (
        set AUTO=1
    ) else if not "%%A"=="" (
        set PY_ARGS=!PY_ARGS! %%A
    )
)

if not exist ".venv\Scripts\activate.bat" (
    echo.
    echo [오류] 설치가 완료되지 않았습니다.
    echo.
    echo        setup_user.bat 을 먼저 실행해주세요.
    echo.
    if %AUTO%==0 pause
    exit /b 1
)
call .venv\Scripts\activate.bat

echo.
echo ============================================================
echo   B2C 고객지표 RPA 실행 중...  [%MODE% 모드]
echo   실행이 끝날 때까지 이 창을 닫지 마세요.
echo ============================================================
echo.

python -m src.main %MODE% %PY_ARGS%
set EXIT_CODE=%ERRORLEVEL%

echo.
if %EXIT_CODE%==0 (
    echo ============================================================
    echo   [완료] 정상적으로 실행되었습니다.
    echo.
    echo   Google Sheets "B2C사업본부 고객지표" 에서 결과를 확인하세요.
    echo   확인 후 이 창을 닫으셔도 됩니다.
    echo ============================================================
) else (
    echo ============================================================
    echo   [오류] 실행 중 문제가 발생했습니다.
    echo.
    echo   logs\ 폴더의 최근 .log 파일을 관리자에게 전달해주세요.
    echo   담당: andrew.chu@secta9ine.co.kr
    echo ============================================================
)
echo.

:: 수동 실행: 최신 로그를 메모장으로 열고 창 유지
if %AUTO%==0 (
    for /f "delims=" %%L in ('dir /b /o-d "logs\*.log" 2^>nul') do (
        start notepad "logs\%%L"
        goto :show_log_done
    )
    :show_log_done
    pause
)

endlocal
exit /b %EXIT_CODE%

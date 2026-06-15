@echo off
:: 고객지표 RPA — 수동 실행 / Task Scheduler 진입점

setlocal
cd /d "%~dp0"

set MODE=%~1
if "%MODE%"=="" set MODE=daily

if exist ".venv\Scripts\activate.bat" (
    call .venv\Scripts\activate.bat
)

echo.
echo ============================================================
echo   B2C 고객지표 RPA 실행 중...  [%MODE% 모드]
echo   실행이 끝날 때까지 이 창을 닫지 마세요.
echo ============================================================
echo.

python -m src.main %MODE% %~2 %~3 %~4
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
pause

endlocal
exit /b %EXIT_CODE%

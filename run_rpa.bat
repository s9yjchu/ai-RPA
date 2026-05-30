@echo off
:: 고객지표 RPA — Windows Task Scheduler 진입점
:: 인수 없이 실행 시 daily 모드로 동작

setlocal
cd /d "%~dp0"

set MODE=%~1
if "%MODE%"=="" set MODE=daily

:: 가상환경 활성화 (없으면 시스템 Python 사용)
if exist ".venv\Scripts\activate.bat" (
    call .venv\Scripts\activate.bat
)

python -m src.main %MODE% %~2 %~3 %~4
set EXIT_CODE=%ERRORLEVEL%

endlocal
exit /b %EXIT_CODE%

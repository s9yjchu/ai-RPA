@echo off
REM GCP Pub/Sub 에이전트 수동 실행
REM 테스트: run_agent.bat --test daily
cd /d "%~dp0"
.venv\Scripts\python -m src.gcp_agent %*
pause

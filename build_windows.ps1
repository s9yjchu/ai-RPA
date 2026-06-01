# 고객지표 RPA — Windows .exe 빌드 (PyInstaller)
# 실행: .\build_windows.ps1

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location $ScriptDir

Write-Host "== 고객지표 RPA Windows 빌드 시작 =="

# Playwright 브라우저를 번들에 포함시키기 위해 경로를 로컬로 설정
$env:PLAYWRIGHT_BROWSERS_PATH = "0"
python -m playwright install chromium

pyinstaller `
    --clean --noconfirm `
    --onedir --windowed `
    --name "고객지표_RPA" `
    --paths . `
    --collect-all playwright `
    --hidden-import src.config `
    --hidden-import src.logger `
    --hidden-import src.browser `
    --hidden-import src.state_manager `
    --hidden-import src.olap_scraper `
    --hidden-import src.excel_parser `
    --hidden-import src.sheets_writer `
    --hidden-import src.notifier `
    --hidden-import src.daily_runner `
    --hidden-import src.monthly_runner `
    --hidden-import src.log_report_scraper `
    --hidden-import src.visual_report_scraper `
    --hidden-import src.setup_gui `
    --hidden-import src.main `
    src\main.py

Write-Host ""
Write-Host "== 빌드 완료: dist\고객지표_RPA\ =="
Write-Host "배포 시 아래 파일을 함께 복사하세요:"
Write-Host "  .env          (환경변수)"
Write-Host "  credentials.json  (Google OAuth)"
Write-Host "  token.json    (최초 인증 후 자동 생성)"
